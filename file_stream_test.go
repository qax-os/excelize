// Copyright 2016 - 2026 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"errors"
	"io"
	"math"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func TestWriteToDirect(t *testing.T) {
	f := NewFile(Options{TmpDir: filepath.Join(t.TempDir(), "missing")})
	defer func() { require.NoError(t, f.Close()) }()
	require.NoError(t, f.SetCellStr("Sheet1", "A1", "中文<&>"))
	var output bytes.Buffer
	direct := false
	f.SetZipWriter(func(w io.Writer) ZipWriter {
		_, direct = w.(*workbookCountWriter)
		return zip.NewWriter(w)
	})
	n, err := f.WriteTo(&output)
	require.NoError(t, err)
	assert.True(t, direct)
	assert.Equal(t, int64(output.Len()), n)
	book, err := OpenReader(bytes.NewReader(output.Bytes()))
	require.NoError(t, err)
	defer func() { require.NoError(t, book.Close()) }()
	value, err := book.GetCellValue("Sheet1", "A1")
	require.NoError(t, err)
	assert.Equal(t, "中文<&>", value)
}

func TestWriteToDirectEncryptionAndErrors(t *testing.T) {
	f := NewFile()
	defer func() { require.NoError(t, f.Close()) }()
	var output bytes.Buffer
	_, err := f.WriteTo(&output, Options{Password: "password"})
	require.NoError(t, err)
	book, err := OpenReader(bytes.NewReader(output.Bytes()), Options{Password: "password"})
	require.NoError(t, err)
	require.NoError(t, book.Close())
	n, err := f.WriteTo(&errWriter{err: errors.New("destination failed")}, Options{})
	assert.EqualError(t, err, "destination failed")
	assert.Zero(t, n)
	counter := &workbookCountWriter{writer: shortWorkbookWriter{}}
	written, err := counter.Write([]byte("1234"))
	assert.Equal(t, 2, written)
	assert.Equal(t, int64(2), counter.written)
	assert.ErrorIs(t, err, io.ErrShortWrite)
}

// shortWorkbookWriter exercises the invalid short-write case of io.Writer.
type shortWorkbookWriter struct{}

func (shortWorkbookWriter) Write(p []byte) (int, error) { return len(p) / 2, nil }

func TestWriteDiskPartWithoutCaching(t *testing.T) {
	f := NewFile()
	defer func() { require.NoError(t, f.Close()) }()
	name := "custom/disk.xml"
	file, err := os.CreateTemp(t.TempDir(), "part-*")
	require.NoError(t, err)
	_, err = io.WriteString(file, "<data>streamed</data>")
	require.NoError(t, err)
	require.NoError(t, file.Close())
	f.tempFiles.Store(name, file.Name())
	var output bytes.Buffer
	_, err = f.WriteTo(&output)
	require.NoError(t, err)
	_, exists := f.Pkg.Load(name)
	assert.False(t, exists)
	archive, err := zip.NewReader(bytes.NewReader(output.Bytes()), int64(output.Len()))
	require.NoError(t, err)
	for _, entry := range archive.File {
		if entry.Name != name {
			continue
		}
		reader, err := entry.Open()
		require.NoError(t, err)
		data, err := io.ReadAll(reader)
		require.NoError(t, err)
		require.NoError(t, reader.Close())
		assert.Equal(t, "<data>streamed</data>", string(data))
		return
	}
	t.Fatal("missing streamed part")
}

func TestZip64SpoolNeeded(t *testing.T) {
	f := NewFile()
	defer func() { require.NoError(t, f.Close()) }()
	needed, err := f.zip64SpoolNeeded()
	require.NoError(t, err)
	assert.False(t, needed)
	file, err := os.CreateTemp(t.TempDir(), "large-*")
	require.NoError(t, err)
	require.NoError(t, file.Truncate(math.MaxUint32+1))
	f.tempFiles.Store("custom/large.xml", file.Name())
	needed, err = f.zip64SpoolNeeded()
	require.NoError(t, err)
	assert.True(t, needed)
	f.Pkg.Store("custom/large.xml", []byte("overridden"))
	needed, err = f.zip64SpoolNeeded()
	require.NoError(t, err)
	assert.False(t, needed)
	f.streams = map[string]*StreamWriter{"sheet": {rawData: bufferedWriter{tmp: file}}}
	needed, err = f.zip64SpoolNeeded()
	require.NoError(t, err)
	assert.True(t, needed)
	require.NoError(t, file.Close())
	_, err = f.zip64SpoolNeeded()
	assert.Error(t, err)
	f.streams = nil
}

func TestWriteToSpooled(t *testing.T) {
	root := t.TempDir()
	f := NewFile(Options{TmpDir: root})
	defer func() { require.NoError(t, f.Close()) }()
	require.NoError(t, f.SetCellStr("Sheet1", "A1", "spooled"))
	f.prepareWriteToZip()
	var output bytes.Buffer
	n, err := f.writeToSpooled(&output)
	require.NoError(t, err)
	assert.Equal(t, int64(output.Len()), n)
	book, err := OpenReader(bytes.NewReader(output.Bytes()))
	require.NoError(t, err)
	require.NoError(t, book.Close())
	files, err := os.ReadDir(root)
	require.NoError(t, err)
	assert.Empty(t, files)
	_, err = f.writeToSpooled(&errWriter{err: errors.New("copy failure")})
	assert.EqualError(t, err, "copy failure")
	files, err = os.ReadDir(root)
	require.NoError(t, err)
	assert.Empty(t, files)
	f.options.TmpDir = filepath.Join(root, "missing")
	_, err = f.writeToSpooled(io.Discard)
	assert.Error(t, err)
}

func TestPatchZip64LFH(t *testing.T) {
	f := NewFile()
	defer func() { require.NoError(t, f.Close()) }()
	name := "xl/worksheets/large.xml"
	f.zip64Entries = []string{name}
	header := make([]byte, 30)
	copy(header, []byte{'P', 'K', 3, 4})
	binary.LittleEndian.PutUint16(header[4:6], 20)
	binary.LittleEndian.PutUint16(header[26:28], uint16(len(name)))
	for _, offset := range []int{0, (64 << 10) - 2, (64 << 10) - 29} {
		file, err := os.CreateTemp(t.TempDir(), "zip64-*")
		require.NoError(t, err)
		_, err = io.Copy(file, strings.NewReader(strings.Repeat("x", offset)))
		require.NoError(t, err)
		_, err = file.Write(header)
		require.NoError(t, err)
		_, err = io.WriteString(file, name)
		require.NoError(t, err)
		require.NoError(t, f.patchZip64LFH(file, file))
		version := make([]byte, 2)
		_, err = file.ReadAt(version, int64(offset+4))
		require.NoError(t, err)
		assert.Equal(t, uint16(45), binary.LittleEndian.Uint16(version))
		require.NoError(t, file.Close())
	}
	for _, data := range [][]byte{header[:20], append(header, []byte("short")...)} {
		reader := bytes.NewReader(data)
		require.NoError(t, f.patchZip64LFH(reader, zip64ErrorWriter{}))
	}
	assert.EqualError(t, f.patchZip64LFH(zip64ErrorReader{}, zip64ErrorWriter{}), "read failed")
	full := append(append([]byte{}, header...), []byte(name)...)
	assert.EqualError(t, f.patchZip64LFH(bytes.NewReader(full), zip64ErrorWriter{}), "patch failed")
}

// zip64ErrorReader reports a random-access read failure.
type zip64ErrorReader struct{}

func (zip64ErrorReader) ReadAt([]byte, int64) (int, error) { return 0, errors.New("read failed") }

// zip64ErrorWriter reports a random-access write failure.
type zip64ErrorWriter struct{}

func (zip64ErrorWriter) WriteAt([]byte, int64) (int, error) { return 0, errors.New("patch failed") }

func BenchmarkWriteToStreaming(b *testing.B) {
	for _, buffered := range []bool{false, true} {
		name := "direct"
		if buffered {
			name = "buffered"
		}
		b.Run(name, func(b *testing.B) {
			f := NewFile()
			defer f.Close()
			sw, err := f.NewStreamWriter("Sheet1")
			require.NoError(b, err)
			for i := 1; i <= 10000; i++ {
				axis, _ := CoordinatesToCellName(1, i)
				require.NoError(b, sw.SetRow(axis, []interface{}{i, strings.Repeat("中文内容", 20), "https://example.com/"}))
			}
			require.NoError(b, sw.Flush())
			b.ReportAllocs()
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				if buffered {
					_, err = f.WriteToBuffer()
				} else {
					_, err = f.WriteTo(io.Discard)
				}
				require.NoError(b, err)
			}
		})
	}
}
