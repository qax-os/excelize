// Copyright 2016 - 2026 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func TestStreamNativeHyperlinks(t *testing.T) {
	root := t.TempDir()
	f := NewFile(Options{TmpDir: root})
	require.NoError(t, f.SetCellHyperLink("Sheet1", "Z1", "https://example.com/existing", "External"))
	// Force a prefix collision with a pre-existing, nonnumeric relationship.
	rels, err := f.relsReader("xl/worksheets/_rels/sheet1.xml.rels")
	require.NoError(t, err)
	rels.Relationships = append(rels.Relationships, xlsxRelationship{ID: "rId_stream_hyperlink_1", Type: SourceRelationshipHyperLink, Target: "https://example.com/reserved", TargetMode: "External"})
	sw, err := f.NewStreamWriter("Sheet1")
	require.NoError(t, err)
	target := `https://example.com/?x="中文"&long=` + strings.Repeat("a", 400)
	display, tooltip := "查看原文", "提示<&>"
	style, err := f.NewStyle(&Style{Font: &Font{Color: "FF0000", Bold: true}})
	require.NoError(t, err)
	require.NoError(t, sw.SetRow("A1", []interface{}{Cell{Value: target, StyleID: style}, "内部定位"}))
	require.NoError(t, sw.SetCellHyperLink("a1", target, "External", HyperlinkOpts{Display: &display, Tooltip: &tooltip}))
	require.NoError(t, sw.SetCellHyperLink("B1", "Sheet1!A1", "Location"))
	// Numeric relationship allocation after streamed links must not collide.
	require.NoError(t, sw.AddTable(&Table{Range: "A1:B2", Name: "Links"}))
	require.NoError(t, sw.Flush())
	for attempt := 0; attempt < 2; attempt++ {
		output, err := f.WriteToBuffer()
		require.NoError(t, err)
		book, err := OpenReader(bytes.NewReader(output.Bytes()))
		require.NoError(t, err)
		value, err := book.GetCellValue("Sheet1", "A1")
		require.NoError(t, err)
		assert.Equal(t, target, value)
		formula, err := book.GetCellFormula("Sheet1", "A1")
		require.NoError(t, err)
		assert.Empty(t, formula)
		gotStyle, err := book.GetCellStyle("Sheet1", "A1")
		require.NoError(t, err)
		assert.Equal(t, style, gotStyle)
		for axis, want := range map[string]string{"A1": target, "B1": "Sheet1!A1", "Z1": "https://example.com/existing"} {
			ok, got, err := book.GetCellHyperLink("Sheet1", axis)
			require.NoError(t, err)
			assert.True(t, ok, axis)
			assert.Equal(t, want, got, axis)
		}
		ws, err := book.workSheetReader("Sheet1")
		require.NoError(t, err)
		assert.Len(t, ws.Hyperlinks.Hyperlink, 3)
		assert.Equal(t, display, ws.Hyperlinks.Hyperlink[1].Display)
		assert.Equal(t, tooltip, ws.Hyperlinks.Hyperlink[1].Tooltip)
		tables, err := book.GetTables("Sheet1")
		require.NoError(t, err)
		assert.Len(t, tables, 1)
		require.NoError(t, book.Close())
		archive, err := zip.NewReader(bytes.NewReader(output.Bytes()), int64(output.Len()))
		require.NoError(t, err)
		seen := map[string]bool{}
		for _, entry := range archive.File {
			assert.False(t, seen[entry.Name], entry.Name)
			seen[entry.Name] = true
		}
	}
	require.NoError(t, f.Close())
	files, err := os.ReadDir(root)
	require.NoError(t, err)
	assert.Empty(t, files)
}

func TestStreamHyperlinkValidation(t *testing.T) {
	f := NewFile()
	defer func() { require.NoError(t, f.Close()) }()
	require.NoError(t, f.SetCellHyperLink("Sheet1", "Z1", "https://example.com/old", "External"))
	sw, err := f.NewStreamWriter("Sheet1")
	require.NoError(t, err)
	require.Error(t, sw.SetHyperlinkLimit(-1))
	require.Error(t, sw.SetCellHyperLink("bad", "https://example.com", "External"))
	require.Error(t, sw.SetCellHyperLink("A1", "https://example.com", "None"))
	require.Error(t, sw.SetCellHyperLink("Z1", "https://example.com", "External"))
	require.NoError(t, sw.SetHyperlinkLimit(3))
	require.NoError(t, sw.SetCellHyperLink("A2", "https://example.com/1", "External"))
	require.Error(t, sw.SetHyperlinkLimit(0))
	require.Error(t, sw.SetCellHyperLink("A1", "https://example.com/descending", "External"))
	require.Error(t, sw.SetCellHyperLink("A2", "https://example.com/duplicate", "External"))
	require.NoError(t, sw.SetCellHyperLink("B2", "Sheet1!A1", "Location"))
	assert.ErrorIs(t, sw.SetCellHyperLink("C2", "https://example.com/limit", "External"), ErrTotalSheetHyperlinks)
	require.NoError(t, sw.Flush())
	require.Error(t, sw.SetCellHyperLink("A3", "https://example.com/late", "External"))
	require.Error(t, sw.SetHyperlinkLimit(0))
}

func TestStreamHyperlinkTemporaryErrors(t *testing.T) {
	f := NewFile(Options{TmpDir: filepath.Join(t.TempDir(), "missing")})
	sw, err := f.NewStreamWriter("Sheet1")
	require.NoError(t, err)
	require.Error(t, sw.SetCellHyperLink("A1", "https://example.com", "External"))
	require.NoError(t, f.Close())
	root := t.TempDir()
	f = NewFile(Options{TmpDir: root})
	sw, err = f.NewStreamWriter("Sheet1")
	require.NoError(t, err)
	require.NoError(t, sw.SetCellHyperLink("A1", "https://example.com", "External"))
	require.NoError(t, sw.hyperlinks.links.Close())
	require.Error(t, sw.Flush())
	require.Error(t, f.Close())
	files, err := os.ReadDir(root)
	require.NoError(t, err)
	assert.Empty(t, files)
}

func TestStreamHyperlinkDefaultLimit(t *testing.T) {
	f := NewFile()
	defer func() { require.NoError(t, f.Close()) }()
	sw, err := f.NewStreamWriter("Sheet1")
	require.NoError(t, err)
	assert.Equal(t, TotalSheetHyperlinks+1, sw.streamHyperlinks().limit)
	sw.hyperlinks.count = TotalSheetHyperlinks + 1
	assert.ErrorIs(t, sw.SetCellHyperLink("A1", "https://example.com", "External"), ErrTotalSheetHyperlinks)
}

func TestStreamHyperlinksBeyondDefaultLimit(t *testing.T) {
	root := t.TempDir()
	f := NewFile(Options{TmpDir: root})
	sw, err := f.NewStreamWriter("Sheet1")
	require.NoError(t, err)
	require.NoError(t, sw.SetHyperlinkLimit(0))
	const rows = 70000
	for row := 1; row <= rows; row++ {
		axis := "A" + strconv.Itoa(row)
		target := "https://example.com/" + strconv.Itoa(row) + "?x=1&y=中文"
		require.NoError(t, sw.SetRow(axis, []interface{}{target}))
		require.NoError(t, sw.SetCellHyperLink(axis, target, "External"))
	}
	assert.Nil(t, sw.worksheet.Hyperlinks, "links must not accumulate in the worksheet object model")
	rels, err := f.relsReader(sw.hyperlinkRelsPath())
	require.NoError(t, err)
	assert.Nil(t, rels, "URLs must not accumulate in relationships")
	assert.LessOrEqual(t, sw.hyperlinks.linkWriter.Buffered(), 64<<10)
	assert.LessOrEqual(t, sw.hyperlinks.relWriter.Buffered(), 64<<10)
	require.NoError(t, sw.Flush())
	output, err := f.WriteToBuffer()
	require.NoError(t, err)
	archive, err := zip.NewReader(bytes.NewReader(output.Bytes()), int64(output.Len()))
	require.NoError(t, err)
	links, relationships, sheets := 0, 0, 0
	for _, entry := range archive.File {
		worksheet := entry.Name == "xl/worksheets/sheet1.xml"
		relationship := entry.Name == "xl/worksheets/_rels/sheet1.xml.rels"
		if !worksheet && !relationship {
			continue
		}
		if worksheet {
			sheets++
		}
		input, err := entry.Open()
		require.NoError(t, err)
		decoder := xml.NewDecoder(input)
		for {
			token, err := decoder.Token()
			if err == io.EOF {
				break
			}
			require.NoError(t, err)
			start, ok := token.(xml.StartElement)
			if !ok {
				continue
			}
			switch start.Name.Local {
			case "hyperlink":
				links++
			case "Relationship":
				relationships++
			case "f":
				t.Fatal("native link became a formula")
			}
		}
		require.NoError(t, input.Close())
	}
	assert.Equal(t, 1, sheets)
	assert.Equal(t, rows, links)
	assert.Equal(t, rows, relationships)
	require.NoError(t, f.Close())
	files, err := os.ReadDir(root)
	require.NoError(t, err)
	assert.Empty(t, files)
}

func BenchmarkStreamNativeHyperlinks(b *testing.B) {
	for _, count := range []int{1000, 10000} {
		b.Run(strconv.Itoa(count), func(b *testing.B) {
			b.ReportAllocs()
			for n := 0; n < b.N; n++ {
				f := NewFile()
				sw, err := f.NewStreamWriter("Sheet1")
				require.NoError(b, err)
				for row := 1; row <= count; row++ {
					axis := "A" + strconv.Itoa(row)
					target := "https://example.com/" + strconv.Itoa(row)
					require.NoError(b, sw.SetRow(axis, []interface{}{target}))
					require.NoError(b, sw.SetCellHyperLink(axis, target, "External"))
				}
				require.NoError(b, sw.Flush())
				require.NoError(b, f.Close())
			}
		})
	}
}
