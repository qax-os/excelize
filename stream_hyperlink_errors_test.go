// Copyright 2016 - 2026 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"bufio"
	"errors"
	"io"
	"math"
	"os"
	"runtime"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

// newHyperlinkErrorStream gives each error case independent files and cleanup.
func newHyperlinkErrorStream(t *testing.T) (*File, *StreamWriter) {
	t.Helper()
	root := t.TempDir()
	f := NewFile(Options{TmpDir: root})
	sw, err := f.NewStreamWriter("Sheet1")
	require.NoError(t, err)
	require.NoError(t, sw.SetCellHyperLink("A1", "https://example.com/first", "External"))
	t.Cleanup(func() {
		// Some cases deliberately close or remove an open file first.
		_ = f.Close()
		files, err := os.ReadDir(root)
		require.NoError(t, err)
		assert.Empty(t, files)
	})
	return f, sw
}

func TestStreamHyperlinkWriteErrors(t *testing.T) {
	for _, phase := range []string{"links", "relationships"} {
		t.Run(phase, func(t *testing.T) {
			_, sw := newHyperlinkErrorStream(t)
			failure := errors.New("spool write failed")
			if phase == "links" {
				sw.hyperlinks.linkWriter = bufio.NewWriterSize(&errWriter{err: failure}, 1)
			} else {
				sw.hyperlinks.relWriter = bufio.NewWriterSize(&errWriter{err: failure}, 1)
			}
			assert.ErrorIs(t, sw.SetCellHyperLink("A2", "https://example.com/next", "External"), failure)
			// Once output is partially written, subsequent operations must fail rather
			// than silently generating an incomplete relationship pair.
			assert.ErrorIs(t, sw.SetCellHyperLink("A3", "https://example.com/later", "External"), failure)
			assert.ErrorIs(t, sw.Flush(), failure)
			assert.ErrorIs(t, sw.writeHyperlinkRelationships(&errZipWriter{}), failure)
		})
	}
}

func TestStreamHyperlinkInvalidTemplateMetadata(t *testing.T) {
	f := NewFile()
	defer func() { require.NoError(t, f.Close()) }()
	sw, err := f.NewStreamWriter("Sheet1")
	require.NoError(t, err)
	sw.worksheet.Hyperlinks = &xlsxHyperlinks{Hyperlink: []xlsxHyperlink{{Ref: "A1:invalid"}}}
	require.Error(t, sw.SetCellHyperLink("A1", "https://example.com", "External"))
	sw.worksheet.Hyperlinks = nil
	f.Pkg.Store(sw.hyperlinkRelsPath(), []byte("<invalid"))
	require.Error(t, sw.SetCellHyperLink("A1", "https://example.com", "External"))
}

func TestStreamHyperlinkFileCreationCleanup(t *testing.T) {
	root := t.TempDir()
	f := NewFile(Options{TmpDir: root})
	defer func() { require.NoError(t, f.Close()) }()
	sw, err := f.NewStreamWriter("Sheet1")
	require.NoError(t, err)
	sw.streamHyperlinks()
	failure := errors.New("second temporary file failed")
	created := 0
	err = sw.openHyperlinkFiles(func(directory, pattern string) (*os.File, error) {
		created++
		if created == 2 {
			return nil, failure
		}
		return os.CreateTemp(directory, pattern)
	})
	assert.ErrorIs(t, err, failure)
	assert.Nil(t, sw.hyperlinks.links)
	files, err := os.ReadDir(root)
	require.NoError(t, err)
	assert.Empty(t, files)
	// A retry must work after a partial allocation was cleaned up.
	require.NoError(t, sw.SetCellHyperLink("A1", "https://example.com/retry", "External"))
	require.NoError(t, sw.Flush())
}

func TestStreamHyperlinkFlushErrors(t *testing.T) {
	t.Run("limit without links", func(t *testing.T) {
		f := NewFile()
		defer func() { require.NoError(t, f.Close()) }()
		sw, err := f.NewStreamWriter("Sheet1")
		require.NoError(t, err)
		require.NoError(t, sw.SetHyperlinkLimit(0))
		require.NoError(t, sw.Flush())
	})
	for _, phase := range []string{"relationship flush", "link seek", "worksheet write", "template worksheet write"} {
		t.Run(phase, func(t *testing.T) {
			_, sw := newHyperlinkErrorStream(t)
			switch phase {
			case "relationship flush":
				require.NoError(t, sw.hyperlinks.rels.Close())
			case "link seek":
				require.NoError(t, sw.hyperlinks.linkWriter.Flush())
				require.NoError(t, sw.hyperlinks.links.Close())
			default:
				file, err := os.CreateTemp(t.TempDir(), "closed-sheet-*")
				require.NoError(t, err)
				require.NoError(t, file.Close())
				sw.rawData.tmp = file
				if phase == "template worksheet write" {
					sw.worksheet.Hyperlinks = &xlsxHyperlinks{Hyperlink: []xlsxHyperlink{{Ref: "Z1", Location: "Sheet1!A1"}}}
					sw.rawData.buf.WriteString(strings.Repeat("x", StreamChunkSize))
				}
			}
			require.Error(t, sw.Flush())
		})
	}
}

func TestStreamHyperlinkRelationshipsErrors(t *testing.T) {
	for _, phase := range []string{"flush", "create", "header", "metadata", "collision", "template write", "seek", "copy", "trailer"} {
		t.Run(phase, func(t *testing.T) {
			f, sw := newHyperlinkErrorStream(t)
			failure := errors.New("relationship output failed")
			switch phase {
			case "flush":
				require.NoError(t, sw.hyperlinks.rels.Close())
			case "metadata":
				f.Pkg.Store(sw.hyperlinkRelsPath(), []byte("<invalid"))
			case "collision":
				f.Relationships.Store(sw.hyperlinkRelsPath(), &xlsxRelationships{Relationships: []xlsxRelationship{{ID: sw.hyperlinks.prefix + "1"}}})
			case "template write":
				f.Relationships.Store(sw.hyperlinkRelsPath(), &xlsxRelationships{Relationships: []xlsxRelationship{{ID: "rId1", Target: "old", Type: SourceRelationshipHyperLink}}})
			case "seek":
				require.NoError(t, sw.hyperlinks.relWriter.Flush())
				require.NoError(t, sw.hyperlinks.rels.Close())
			}
			calls := 0
			archive := &errZipWriter{createFunc: func(string) (io.Writer, error) {
				if phase == "create" {
					return nil, failure
				}
				return hyperlinkCallbackWriter{write: func(data []byte) (int, error) {
					calls++
					if (phase == "header" && calls == 1) || (phase == "template write" && calls == 2) || (phase == "copy" && calls == 2) || (phase == "trailer" && calls == 3) {
						return 0, failure
					}
					return len(data), nil
				}}, nil
			}}
			err := sw.writeHyperlinkRelationships(archive)
			require.Error(t, err)
			switch phase {
			case "create", "header", "template write", "copy", "trailer":
				assert.ErrorIs(t, err, failure)
			case "collision":
				assert.Contains(t, err.Error(), "ID collision")
			}
		})
	}
}

func TestStreamHyperlinkZIPErrorPropagation(t *testing.T) {
	f, sw := newHyperlinkErrorStream(t)
	require.NoError(t, sw.Flush())
	failure := errors.New("create relationship ZIP entry failed")
	f.SetZipWriter(func(io.Writer) ZipWriter {
		return &errZipWriter{createFunc: func(name string) (io.Writer, error) {
			if name == sw.hyperlinkRelsPath() {
				return nil, failure
			}
			return io.Discard, nil
		}}
	})
	_, err := f.WriteToBuffer()
	assert.ErrorIs(t, err, failure)
}

func TestStreamHyperlinkRemovalError(t *testing.T) {
	if runtime.GOOS == "windows" {
		t.Skip("Windows does not permit unlinking an open spool")
	}
	_, sw := newHyperlinkErrorStream(t)
	// The handle can still be closed even if another process removed the path.
	require.NoError(t, os.Remove(sw.hyperlinks.links.Name()))
	err := sw.closeHyperlinks()
	assert.ErrorIs(t, err, os.ErrNotExist)
}

func TestStreamHyperlinkLargeRelationshipPart(t *testing.T) {
	_, sw := newHyperlinkErrorStream(t)
	require.NoError(t, sw.hyperlinks.relWriter.Flush())
	// A sparse file covers the ZIP64 size boundary without allocating a 4 GiB
	// object or spending CPU on compression. The fake ZIP sink discards bytes.
	require.NoError(t, sw.hyperlinks.rels.Truncate(math.MaxUint32+1))
	archive := &errZipWriter{createFunc: func(string) (io.Writer, error) { return io.Discard, nil }}
	require.NoError(t, sw.writeHyperlinkRelationships(archive))
	assert.Contains(t, sw.file.zip64Entries, sw.hyperlinkRelsPath())
}

// hyperlinkCallbackWriter injects failures at individual ZIP write boundaries.
type hyperlinkCallbackWriter struct{ write func([]byte) (int, error) }

func (w hyperlinkCallbackWriter) Write(p []byte) (int, error) { return w.write(p) }
