// Copyright 2016 - 2026 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"bufio"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"math"
	"os"
	"path"
	"strconv"
	"strings"
)

// streamHyperlinks keeps only counters and fixed-size buffers in memory.
// Link XML and external relationship XML are stored in temporary files.
type streamHyperlinks struct {
	links, rels                         *os.File
	linkWriter, relWriter               *bufio.Writer
	record                              bytes.Buffer
	prefix                              string
	count, external, row, column, limit int
	err                                 error
}

// SetHyperlinkLimit sets the maximum number of hyperlinks in this stream,
// including hyperlinks already present in the worksheet. Call it before adding
// stream hyperlinks. The default is 65,530, the same as File.SetCellHyperLink.
// Zero disables the count check; files exceeding Excel's documented worksheet
// limit may not be supported by spreadsheet applications. A negative limit is
// invalid. Disabling this check does not split the worksheet.
func (sw *StreamWriter) SetHyperlinkLimit(limit int) error {
	if sw.flushed {
		return errors.New("cannot set hyperlinks after stream flush")
	}
	if limit < 0 {
		return errors.New("hyperlink limit must not be negative")
	}
	links := sw.streamHyperlinks()
	if links.count != 0 {
		return errors.New("must set hyperlink limit before adding stream hyperlinks")
	}
	links.limit = limit
	return nil
}

// SetCellHyperLink appends a native hyperlink without changing the cell value
// or style. It supports External and Location links and HyperlinkOpts display
// and tooltip attributes. Call it before Flush, in strictly increasing row and
// column order, independently of SetRow calls. Replacing or removing hyperlinks
// and overlapping existing template hyperlinks are not supported in stream
// mode. Links and relationships spill directly to temporary files, avoiding
// per-link objects and repeated scans of the accumulated links.
//
// For example:
//
//	sw, err := f.NewStreamWriter("Sheet1")
//	if err != nil { return err }
//	if err = sw.SetRow("A1", []interface{}{ "https://example.com" }); err != nil { return err }
//	if err = sw.SetCellHyperLink("A1", "https://example.com", "External"); err != nil { return err }
//	return sw.Flush()
func (sw *StreamWriter) SetCellHyperLink(cell, link, linkType string, opts ...HyperlinkOpts) error {
	if sw.flushed {
		return errors.New("cannot set hyperlinks after stream flush")
	}
	column, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return err
	}
	axis, err := CoordinatesToCellName(column, row)
	if err != nil {
		return err
	}
	if linkType != "External" && linkType != "Location" {
		return newInvalidLinkTypeError(linkType)
	}
	links := sw.streamHyperlinks()
	if links.err != nil {
		return links.err
	}
	if row < links.row || (row == links.row && column <= links.column) {
		return errors.New("stream hyperlinks must be in ascending cell order")
	}
	existing := 0
	if sw.worksheet.Hyperlinks != nil {
		existing = len(sw.worksheet.Hyperlinks.Hyperlink)
		for _, item := range sw.worksheet.Hyperlinks.Hyperlink {
			overlap, err := sw.file.checkCellInRangeRef(axis, item.Ref)
			if err != nil {
				return err
			}
			if axis == item.Ref || overlap {
				return errors.New("stream hyperlink overlaps an existing hyperlink")
			}
		}
	}
	if links.limit > 0 && links.count+existing >= links.limit {
		return ErrTotalSheetHyperlinks
	}
	if links.links == nil {
		if err := sw.openHyperlinkFiles(); err != nil {
			return err
		}
	}
	data := xlsxHyperlink{Ref: axis}
	relationship := xlsxRelationship{Type: SourceRelationshipHyperLink, Target: link, TargetMode: "External"}
	if linkType == "External" {
		relationship.ID = links.prefix + strconv.Itoa(links.external+1)
		data.RID = relationship.ID
	} else {
		data.Location = link
	}
	for _, opt := range opts {
		if opt.Display != nil {
			data.Display = *opt.Display
		}
		if opt.Tooltip != nil {
			data.Tooltip = *opt.Tooltip
		}
	}
	// Reuse a single record buffer. xml.EscapeText retains the standard XML
	// escaping rules without reflection and encoder allocation per hyperlink.
	links.record.Reset()
	links.record.WriteString(`<hyperlink ref="` + data.Ref + `"`)
	if data.RID != "" {
		links.record.WriteString(` r:id="` + data.RID + `"`)
	}
	if data.Location != "" {
		appendHyperlinkAttribute(&links.record, "location", data.Location)
	}
	if data.Display != "" {
		appendHyperlinkAttribute(&links.record, "display", data.Display)
	}
	if data.Tooltip != "" {
		appendHyperlinkAttribute(&links.record, "tooltip", data.Tooltip)
	}
	links.record.WriteString("/>")
	if _, links.err = links.linkWriter.Write(links.record.Bytes()); links.err != nil {
		return links.err
	}
	if linkType == "External" {
		links.record.Reset()
		links.record.WriteString(`<Relationship Id="` + relationship.ID + `" Type="` + SourceRelationshipHyperLink + `"`)
		appendHyperlinkAttribute(&links.record, "Target", relationship.Target)
		links.record.WriteString(` TargetMode="External"/>`)
		if _, links.err = links.relWriter.Write(links.record.Bytes()); links.err != nil {
			return links.err
		}
		links.external++
	}
	links.count++
	links.row, links.column = row, column
	return nil
}

// streamHyperlinks initializes stream-local metadata without opening files.
func (sw *StreamWriter) streamHyperlinks() *streamHyperlinks {
	if sw.hyperlinks == nil {
		sw.hyperlinks = &streamHyperlinks{limit: TotalSheetHyperlinks + 1}
	}
	return sw.hyperlinks
}

// hyperlinkRelsPath returns the worksheet relationship part name.
func (sw *StreamWriter) hyperlinkRelsPath() string {
	sheetPath := sw.file.sheetMap[sw.Sheet]
	return path.Join(path.Dir(sheetPath), "_rels", path.Base(sheetPath)+".rels")
}

// openHyperlinkFiles reserves a relationship namespace independent of numeric
// IDs used by tables, drawings, and other ordinary worksheet relationships.
func (sw *StreamWriter) openHyperlinkFiles() error {
	links := sw.hyperlinks
	rels, err := sw.file.relsReader(sw.hyperlinkRelsPath())
	if err != nil {
		return err
	}
	links.prefix = "rId_stream_hyperlink_"
	if rels != nil {
		for {
			collision := false
			for _, rel := range rels.Relationships {
				if strings.HasPrefix(rel.ID, links.prefix) {
					collision = true
					break
				}
			}
			if !collision {
				break
			}
			links.prefix += "_"
		}
	}
	links.links, err = os.CreateTemp(sw.file.options.TmpDir, "excelize-hyperlinks-*")
	if err != nil {
		return err
	}
	links.rels, err = os.CreateTemp(sw.file.options.TmpDir, "excelize-hyperlink-rels-*")
	if err != nil {
		_ = links.links.Close()
		_ = os.Remove(links.links.Name())
		links.links = nil
		return err
	}
	links.linkWriter = bufio.NewWriterSize(links.links, 64<<10)
	links.relWriter = bufio.NewWriterSize(links.rels, 64<<10)
	return nil
}

// flushHyperlinks appends native links at their prescribed worksheet position.
func (sw *StreamWriter) flushHyperlinks() error {
	links := sw.hyperlinks
	if links == nil {
		bulkAppendFields(&sw.rawData, sw.worksheet, 21, 21)
		return nil
	}
	if links.err != nil {
		return links.err
	}
	if links.count == 0 {
		bulkAppendFields(&sw.rawData, sw.worksheet, 21, 21)
		return nil
	}
	if err := links.linkWriter.Flush(); err != nil {
		return err
	}
	if err := links.relWriter.Flush(); err != nil {
		return err
	}
	_, _ = sw.rawData.WriteString(`<hyperlinks xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`)
	if sw.worksheet.Hyperlinks != nil {
		for _, link := range sw.worksheet.Hyperlinks.Hyperlink {
			data, err := xml.Marshal(struct {
				XMLName xml.Name `xml:"hyperlink"`
				xlsxHyperlink
			}{xlsxHyperlink: link})
			if err != nil {
				return err
			}
			_, _ = sw.rawData.Write(replaceRelationshipsBytes(data))
			if err := sw.rawData.Sync(); err != nil {
				return err
			}
		}
	}
	if _, err := links.links.Seek(0, io.SeekStart); err != nil {
		return err
	}
	if _, err := io.Copy(streamHyperlinkWriter{sw}, links.links); err != nil {
		return err
	}
	_, _ = sw.rawData.WriteString("</hyperlinks>")
	return nil
}

// streamHyperlinkWriter bounds worksheet memory while copying link fragments.
type streamHyperlinkWriter struct{ stream *StreamWriter }

// Write periodically spills copied metadata just like SetRow spills cell XML.
func (w streamHyperlinkWriter) Write(data []byte) (int, error) {
	n, err := w.stream.rawData.Write(data)
	if err != nil {
		return n, err
	}
	if w.stream.rawData.tmp != nil {
		return n, w.stream.rawData.Flush()
	}
	return n, w.stream.rawData.Sync()
}

// writeHyperlinkRelationships writes existing and streamed relationships once,
// without adding streamed URLs to File.Relationships or File.Pkg.
func (sw *StreamWriter) writeHyperlinkRelationships(archive ZipWriter) error {
	links := sw.hyperlinks
	if links == nil || links.external == 0 {
		return nil
	}
	if links.err != nil {
		return links.err
	}
	if err := links.relWriter.Flush(); err != nil {
		return err
	}
	output, err := archive.Create(sw.hyperlinkRelsPath())
	if err != nil {
		return err
	}
	counter := &hyperlinkPartWriter{writer: output}
	output = counter
	if _, err = io.WriteString(output, xml.Header+`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`); err != nil {
		return err
	}
	rels, err := sw.file.relsReader(sw.hyperlinkRelsPath())
	if err != nil {
		return err
	}
	if rels != nil {
		encoder := xml.NewEncoder(output)
		for _, rel := range rels.Relationships {
			if strings.HasPrefix(rel.ID, links.prefix) {
				return fmt.Errorf("stream hyperlink relationship ID collision: %s", rel.ID)
			}
			if err = encoder.EncodeElement(rel, xml.StartElement{Name: xml.Name{Local: "Relationship"}}); err != nil {
				return err
			}
		}
	}
	if _, err = links.rels.Seek(0, io.SeekStart); err != nil {
		return err
	}
	if _, err = io.Copy(output, links.rels); err != nil {
		return err
	}
	_, err = io.WriteString(output, "</Relationships>")
	if counter.written > math.MaxUint32 {
		sw.file.zip64Entries = append(sw.file.zip64Entries, sw.hyperlinkRelsPath())
	}
	return err
}

// closeHyperlinks removes stream-local temporary files, including on errors.
func (sw *StreamWriter) closeHyperlinks() error {
	if sw.hyperlinks == nil {
		return nil
	}
	var firstErr error
	for _, file := range []*os.File{sw.hyperlinks.links, sw.hyperlinks.rels} {
		if file == nil {
			continue
		}
		if err := file.Close(); err != nil && firstErr == nil {
			firstErr = err
		}
		if err := os.Remove(file.Name()); err != nil && firstErr == nil {
			firstErr = err
		}
	}
	sw.hyperlinks = nil
	return firstErr
}

// appendHyperlinkAttribute escapes attribute values into the reusable record
// buffer. bytes.Buffer writes cannot fail.
func appendHyperlinkAttribute(record *bytes.Buffer, name, value string) {
	record.WriteByte(' ')
	record.WriteString(name)
	record.WriteString(`="`)
	_ = xml.EscapeText(record, []byte(value))
	record.WriteByte('"')
}

// hyperlinkPartWriter records relationship-part sizes for ZIP64 header fixup.
type hyperlinkPartWriter struct {
	writer  io.Writer
	written int64
}

// Write counts actual bytes written to the relationship ZIP entry.
func (w *hyperlinkPartWriter) Write(data []byte) (int, error) {
	n, err := w.writer.Write(data)
	w.written += int64(n)
	return n, err
}
