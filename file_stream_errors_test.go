// Copyright 2016 - 2026 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"bytes"
	"encoding/binary"
	"errors"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func TestBufferedArchiveFailures(t *testing.T) {
	for _, failClose := range []bool{false, true} {
		f := NewFile()
		f.SetZipWriter(func(io.Writer) ZipWriter {
			if failClose {
				return &errZipWriter{closeErr: errors.New("archive close failed")}
			}
			return &errZipWriter{createFunc: func(string) (io.Writer, error) { return nil, errors.New("entry create failed") }}
		})
		_, err := f.WriteToBuffer()
		if failClose {
			assert.EqualError(t, err, "archive close failed")
		} else {
			assert.EqualError(t, err, "entry create failed")
		}
		require.NoError(t, f.Close())
	}
}

func TestDiskPartFailures(t *testing.T) {
	for _, missing := range []bool{false, true} {
		f := NewFile()
		root := t.TempDir()
		filename := filepath.Join(root, "part.xml")
		f.tempFiles.Store("custom/disk.xml", filename)
		if !missing {
			require.NoError(t, os.WriteFile(filename, []byte("disk part"), 0600))
		}
		f.SetZipWriter(func(io.Writer) ZipWriter {
			return &errZipWriter{createFunc: func(name string) (io.Writer, error) {
				if name == "custom/disk.xml" {
					return &errWriter{err: errors.New("entry write failed")}, nil
				}
				return io.Discard, nil
			}}
		})
		_, err := f.WriteTo(io.Discard)
		if missing {
			require.NoError(t, err)
			f.tempFiles.Delete("custom/disk.xml")
		} else {
			assert.EqualError(t, err, "entry write failed")
		}
		require.NoError(t, f.Close())
	}
}

func TestWorkbookPartCloseFailures(t *testing.T) {
	for _, failWrite := range []bool{false, true} {
		input := &workbookPartCloseError{Reader: strings.NewReader("data")}
		var output io.Writer = io.Discard
		if failWrite {
			output = &errWriter{err: errors.New("write failed")}
		}
		_, err := copyWorkbookPart(output, input)
		assert.True(t, input.closed)
		if failWrite {
			assert.EqualError(t, err, "write failed")
		} else {
			assert.EqualError(t, err, "part close failed")
		}
	}
}

// workbookPartCloseError verifies that the copy path always closes its reader.
type workbookPartCloseError struct {
	io.Reader
	closed bool
}

func (r *workbookPartCloseError) Close() error {
	r.closed = true
	return errors.New("part close failed")
}

func TestStreamPartReadFailure(t *testing.T) {
	f := NewFile()
	file, err := os.CreateTemp(t.TempDir(), "stream-*")
	require.NoError(t, err)
	require.NoError(t, file.Close())
	f.streams = map[string]*StreamWriter{"stream": {rawData: bufferedWriter{tmp: file}}}
	archive := &errZipWriter{}
	require.Error(t, f.writeToZip(archive))
	require.NoError(t, f.Close())
}

func TestSpooledArchiveFailureCleanup(t *testing.T) {
	for _, phase := range []string{"write", "patch", "seek"} {
		t.Run(phase, func(t *testing.T) {
			root := t.TempDir()
			f := NewFile(Options{TmpDir: root})
			defer func() { require.NoError(t, f.Close()) }()
			f.prepareWriteToZip()
			f.SetZipWriter(func(output io.Writer) ZipWriter {
				counter := output.(*workbookCountWriter)
				file := counter.writer.(*os.File)
				return &errZipWriter{createFunc: func(string) (io.Writer, error) {
					if phase == "write" {
						return nil, errors.New("ZIP entry failed")
					}
					if phase == "patch" {
						f.zip64Entries = []string{"large.xml"}
					}
					// Closing the output models a destination disappearing during ZIP output.
					// Leave the fake ZIP writer usable so the following file operation runs.
					if _, err := file.Stat(); err == nil {
						require.NoError(t, file.Close())
					}
					return io.Discard, nil
				}}
			})
			_, err := f.writeToSpooled(io.Discard)
			require.Error(t, err)
			if phase == "write" {
				assert.EqualError(t, err, "ZIP entry failed")
			}
			entries, err := os.ReadDir(root)
			require.NoError(t, err)
			assert.Empty(t, entries)
		})
	}
}

func TestZip64RandomAccessFailures(t *testing.T) {
	f := NewFile()
	defer func() { require.NoError(t, f.Close()) }()
	f.zip64Entries = []string{"large.xml"}
	header := make([]byte, 30)
	copy(header, []byte{'P', 'K', 3, 4})
	binary.LittleEndian.PutUint16(header[26:28], uint16(len("large.xml")))
	data := append(header, []byte("large.xml")...)
	for _, failAt := range []int64{0, 30} {
		reader := &zip64OffsetErrorReader{ReaderAt: bytes.NewReader(data), offset: failAt}
		assert.EqualError(t, f.patchZip64LFH(reader, zip64ErrorWriter{}), "offset read failed")
	}
	assert.ErrorIs(t, f.patchZip64LFH(bytes.NewReader(data), zip64ShortWriter{}), io.ErrShortWrite)
}

// zip64OffsetErrorReader fails on the exact header/name reads, not the scan.
type zip64OffsetErrorReader struct {
	io.ReaderAt
	offset int64
}

func (r *zip64OffsetErrorReader) ReadAt(p []byte, offset int64) (int, error) {
	if offset == r.offset && (len(p) == 30 || len(p) == len("large.xml")) {
		return 0, errors.New("offset read failed")
	}
	return r.ReaderAt.ReadAt(p, offset)
}

// zip64ShortWriter checks the random-access short-write contract.
type zip64ShortWriter struct{}

func (zip64ShortWriter) WriteAt(p []byte, _ int64) (int, error) { return len(p) - 1, nil }
