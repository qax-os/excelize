package excelize

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"math/rand"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
)

func BenchmarkStreamWriter(b *testing.B) {
	file := NewFile()
	defer func() {
		if err := file.Close(); err != nil {
			b.Error(err)
		}
	}()
	row := make([]interface{}, 10)
	for colID := 0; colID < 10; colID++ {
		row[colID] = colID
	}

	for n := 0; n < b.N; n++ {
		streamWriter, _ := file.NewStreamWriter("Sheet1")
		for rowID := 10; rowID <= 110; rowID++ {
			cell, _ := CoordinatesToCellName(1, rowID)
			_ = streamWriter.SetRow(cell, row)
		}
	}

	b.ReportAllocs()
}

func TestStreamWriter(t *testing.T) {
	file := NewFile()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)

	// Test max characters in a cell
	row := make([]interface{}, 1)
	row[0] = strings.Repeat("c", TotalCellChars+2)
	assert.NoError(t, streamWriter.SetRow("A1", row))

	// Test leading and ending space(s) character characters in a cell
	row = make([]interface{}, 1)
	row[0] = " characters"
	assert.NoError(t, streamWriter.SetRow("A2", row))

	row = make([]interface{}, 1)
	row[0] = []byte("Word")
	assert.NoError(t, streamWriter.SetRow("A3", row))

	// Test set cell with style and rich text
	styleID, err := file.NewStyle(&Style{Font: &Font{Color: "777777"}})
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.SetRow("A4", []interface{}{
		Cell{StyleID: styleID},
		Cell{Formula: "SUM(A10,B10)", Value: " preserve space "},
	},
		RowOpts{Height: 45, StyleID: styleID}))
	assert.NoError(t, streamWriter.SetRow("A5", []interface{}{
		&Cell{StyleID: styleID, Value: "cell <>&'\""},
		&Cell{Formula: "SUM(A10,B10)"},
		[]RichTextRun{
			{Text: "Rich ", Font: &Font{Color: "2354E8"}},
			{Text: "Text", Font: &Font{Color: "E83723"}},
		},
	}))
	assert.NoError(t, streamWriter.SetRow("A6", []interface{}{time.Now()}))
	assert.NoError(t, streamWriter.SetRow("A7", nil, RowOpts{Height: 20, Hidden: true, StyleID: styleID}))
	assert.Equal(t, ErrMaxRowHeight, streamWriter.SetRow("A8", nil, RowOpts{Height: MaxRowHeight + 1}))

	assert.NoError(t, streamWriter.SetRow("A9", []interface{}{math.NaN(), math.Inf(0), math.Inf(-1)}))

	for rowID := 10; rowID <= 51200; rowID++ {
		row := make([]interface{}, 50)
		for colID := 0; colID < 50; colID++ {
			row[colID] = rand.Intn(640000)
		}
		cell, _ := CoordinatesToCellName(1, rowID)
		assert.NoError(t, streamWriter.SetRow(cell, row))
	}

	assert.NoError(t, streamWriter.Flush())
	// Save spreadsheet by the given path
	assert.NoError(t, file.SaveAs(filepath.Join("test", "TestStreamWriter.xlsx")))

	// Test set cell column overflow
	assert.ErrorIs(t, streamWriter.SetRow("XFD51201", []interface{}{"A", "B", "C"}), ErrColumnNumber)
	assert.NoError(t, file.Close())

	// Test close temporary file error
	file = NewFile(Options{TmpDir: os.TempDir()})
	streamWriter, err = file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	for rowID := 10; rowID <= 25600; rowID++ {
		row := make([]interface{}, 50)
		for colID := 0; colID < 50; colID++ {
			row[colID] = rand.Intn(640000)
		}
		cell, _ := CoordinatesToCellName(1, rowID)
		assert.NoError(t, streamWriter.SetRow(cell, row))
	}
	assert.NoError(t, streamWriter.rawData.Close())
	assert.Error(t, streamWriter.Flush())

	streamWriter.rawData.tmp, err = os.CreateTemp(os.TempDir(), "excelize-")
	assert.NoError(t, err)
	_, err = streamWriter.rawData.Reader()
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.rawData.tmp.Close())
	assert.NoError(t, os.Remove(streamWriter.rawData.tmp.Name()))

	// Test create stream writer with unsupported charset
	file = NewFile()
	file.Sheet.Delete("xl/worksheets/sheet1.xml")
	file.Pkg.Store("xl/worksheets/sheet1.xml", MacintoshCyrillicCharset)
	_, err = file.NewStreamWriter("Sheet1")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, file.Close())

	// Test read cell
	file = NewFile()
	streamWriter, err = file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.SetRow("A1", []interface{}{Cell{StyleID: styleID, Value: "Data"}}))
	assert.NoError(t, streamWriter.Flush())
	cellValue, err := file.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "Data", cellValue)

	// Test stream reader for a worksheet with huge amounts of data
	file, err = OpenFile(filepath.Join("test", "TestStreamWriter.xlsx"))
	assert.NoError(t, err)
	rows, err := file.Rows("Sheet1")
	assert.NoError(t, err)
	cells := 0
	for rows.Next() {
		row, err := rows.Columns()
		assert.NoError(t, err)
		cells += len(row)
	}
	assert.NoError(t, rows.Close())
	assert.Equal(t, 2559562, cells)
	// Save spreadsheet with password.
	assert.NoError(t, file.SaveAs(filepath.Join("test", "EncryptionTestStreamWriter.xlsx"), Options{Password: "password"}))
	assert.NoError(t, file.Close())
}

func TestStreamSetColVisible(t *testing.T) {
	file := NewFile()
	defer func() {
		assert.NoError(t, file.Close())
	}()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.SetColVisible(3, 2, false))
	assert.Equal(t, ErrColumnNumber, streamWriter.SetColVisible(0, 3, false))
	assert.Equal(t, ErrColumnNumber, streamWriter.SetColVisible(MaxColumns+1, 3, false))
	assert.NoError(t, streamWriter.SetRow("A1", []interface{}{"A", "B", "C"}))
	assert.Equal(t, newStreamSetRowOrderError("SetColVisible"), streamWriter.SetColVisible(2, 3, false))
	assert.NoError(t, streamWriter.Flush())
}

func TestStreamSetColOutlineLevel(t *testing.T) {
	file := NewFile()
	defer func() {
		assert.NoError(t, file.Close())
	}()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.SetColOutlineLevel(4, 2))
	assert.Equal(t, ErrOutlineLevel, streamWriter.SetColOutlineLevel(4, 0))
	assert.Equal(t, ErrOutlineLevel, streamWriter.SetColOutlineLevel(4, 8))
	assert.Equal(t, ErrColumnNumber, streamWriter.SetColOutlineLevel(0, 2))
	assert.Equal(t, ErrColumnNumber, streamWriter.SetColOutlineLevel(MaxColumns+1, 2))
	assert.NoError(t, streamWriter.SetRow("A1", []interface{}{"A", "B", "C"}))
	assert.Equal(t, newStreamSetRowOrderError("SetColOutlineLevel"), streamWriter.SetColOutlineLevel(4, 2))
	assert.NoError(t, streamWriter.Flush())
}

func TestStreamSetColStyle(t *testing.T) {
	file := NewFile()
	defer func() {
		assert.NoError(t, file.Close())
	}()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.SetColStyle(3, 2, 0))
	assert.Equal(t, ErrColumnNumber, streamWriter.SetColStyle(0, 3, 20))
	assert.Equal(t, ErrColumnNumber, streamWriter.SetColStyle(MaxColumns+1, 3, 20))
	assert.Equal(t, newInvalidStyleID(2), streamWriter.SetColStyle(1, 3, 2))
	assert.NoError(t, streamWriter.SetRow("A1", []interface{}{"A", "B", "C"}))
	assert.Equal(t, newStreamSetRowOrderError("SetColStyle"), streamWriter.SetColStyle(2, 3, 0))

	file = NewFile()
	defer func() {
		assert.NoError(t, file.Close())
	}()
	// Test set column style with unsupported charset style sheet
	file.Styles = nil
	file.Pkg.Store(defaultXMLPathStyles, MacintoshCyrillicCharset)
	streamWriter, err = file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.EqualError(t, streamWriter.SetColStyle(3, 2, 0), "XML syntax error on line 1: invalid UTF-8")
}

func TestStreamSetColWidth(t *testing.T) {
	file := NewFile()
	defer func() {
		assert.NoError(t, file.Close())
	}()
	styleID, err := file.NewStyle(&Style{
		Fill: Fill{Type: "pattern", Color: []string{"E0EBF5"}, Pattern: 1},
	})
	if err != nil {
		fmt.Println(err)
	}
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.SetColWidth(3, 2, 20))
	assert.NoError(t, streamWriter.SetColStyle(3, 2, styleID))
	assert.Equal(t, ErrColumnNumber, streamWriter.SetColWidth(0, 3, 20))
	assert.Equal(t, ErrColumnNumber, streamWriter.SetColWidth(MaxColumns+1, 3, 20))
	assert.Equal(t, ErrColumnWidth, streamWriter.SetColWidth(1, 3, MaxColumnWidth+1))
	assert.NoError(t, streamWriter.SetRow("A1", []interface{}{"A", "B", "C"}))
	assert.Equal(t, newStreamSetRowOrderError("SetColWidth"), streamWriter.SetColWidth(2, 3, 20))
	assert.NoError(t, streamWriter.Flush())
}

func TestStreamSetPanes(t *testing.T) {
	file, paneOpts := NewFile(), &Panes{
		Freeze:      true,
		Split:       false,
		XSplit:      1,
		YSplit:      0,
		TopLeftCell: "B1",
		ActivePane:  "topRight",
		Selection: []Selection{
			{SQRef: "K16", ActiveCell: "K16", Pane: "topRight"},
		},
	}
	defer func() {
		assert.NoError(t, file.Close())
	}()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.SetPanes(paneOpts))
	assert.Equal(t, ErrParameterInvalid, streamWriter.SetPanes(nil))
	assert.NoError(t, streamWriter.SetRow("A1", []interface{}{"A", "B", "C"}))
	assert.Equal(t, newStreamSetRowOrderError("SetPanes"), streamWriter.SetPanes(paneOpts))
}

func TestStreamTable(t *testing.T) {
	file := NewFile()
	defer func() {
		assert.NoError(t, file.Close())
	}()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	// Test add table without table header
	assert.EqualError(t, streamWriter.AddTable(&Table{Range: "A1:C2"}), "XML syntax error on line 2: unexpected EOF")
	// Write some rows. We want enough rows to force a temp file (>16MB)
	assert.NoError(t, streamWriter.SetRow("A1", []interface{}{"A", "B", "C"}))
	row := []interface{}{1, 2, 3}
	for r := 2; r < 10000; r++ {
		assert.NoError(t, streamWriter.SetRow(fmt.Sprintf("A%d", r), row))
	}

	// Write a table
	assert.NoError(t, streamWriter.AddTable(&Table{Range: "A1:C2"}))
	assert.NoError(t, streamWriter.Flush())

	// Verify the table has names
	var table xlsxTable
	val, ok := file.Pkg.Load("xl/tables/table1.xml")
	assert.True(t, ok)
	assert.NoError(t, xml.Unmarshal(val.([]byte), &table))
	assert.Equal(t, "A", table.TableColumns.TableColumn[0].Name)
	assert.Equal(t, "B", table.TableColumns.TableColumn[1].Name)
	assert.Equal(t, "C", table.TableColumns.TableColumn[2].Name)

	assert.NoError(t, streamWriter.AddTable(&Table{Range: "A1:C1"}))

	// Test add table with illegal cell reference
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), streamWriter.AddTable(&Table{Range: "A:B1"}))
	assert.Equal(t, newCellNameToCoordinatesError("B", newInvalidCellNameError("B")), streamWriter.AddTable(&Table{Range: "A1:B"}))
	// Test add table with invalid table name
	assert.Equal(t, newInvalidNameError("1Table"), streamWriter.AddTable(&Table{Range: "A:B1", Name: "1Table"}))
	// Test add table with row number exceeds maximum limit
	assert.Equal(t, ErrMaxRows, streamWriter.AddTable(&Table{Range: "A1048576:C1048576"}))
	// Test add table with unsupported charset content types
	file.ContentTypes = nil
	file.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, streamWriter.AddTable(&Table{Range: "A1:C2"}), "XML syntax error on line 1: invalid UTF-8")
}

func TestStreamMergeCells(t *testing.T) {
	file := NewFile()
	defer func() {
		assert.NoError(t, file.Close())
	}()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.MergeCell("A1", "D1"))
	// Test merge cells with illegal cell reference
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), streamWriter.MergeCell("A", "D1"))
	assert.NoError(t, streamWriter.Flush())
	// Save spreadsheet by the given path
	assert.NoError(t, file.SaveAs(filepath.Join("test", "TestStreamMergeCells.xlsx")))
}

func TestStreamInsertPageBreak(t *testing.T) {
	file := NewFile()
	defer func() {
		assert.NoError(t, file.Close())
	}()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.InsertPageBreak("A1"))
	assert.NoError(t, streamWriter.Flush())
	// Save spreadsheet by the given path
	assert.NoError(t, file.SaveAs(filepath.Join("test", "TestStreamInsertPageBreak.xlsx")))
}

func TestNewStreamWriter(t *testing.T) {
	// Test error exceptions
	file := NewFile()
	defer func() {
		assert.NoError(t, file.Close())
	}()
	_, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	_, err = file.NewStreamWriter("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Test new stream write with invalid sheet name
	_, err = file.NewStreamWriter("Sheet:1")
	assert.Equal(t, ErrSheetNameInvalid, err)
}

func TestStreamMarshalAttrs(t *testing.T) {
	var r *RowOpts
	attrs, err := r.marshalAttrs()
	assert.NoError(t, err)
	assert.Empty(t, attrs)
}

func TestStreamSetRow(t *testing.T) {
	// Test error exceptions
	file := NewFile()
	defer func() {
		assert.NoError(t, file.Close())
	}()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), streamWriter.SetRow("A", []interface{}{}))
	// Test set row with non-ascending row number
	assert.NoError(t, streamWriter.SetRow("A1", []interface{}{}))
	assert.Equal(t, newStreamSetRowError(1), streamWriter.SetRow("A1", []interface{}{}))
	// Test set row with unsupported charset workbook
	file.WorkBook = nil
	file.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	assert.EqualError(t, streamWriter.SetRow("A2", []interface{}{time.Now()}), "XML syntax error on line 1: invalid UTF-8")
}

func TestStreamSetRowNilValues(t *testing.T) {
	file := NewFile()
	defer func() {
		assert.NoError(t, file.Close())
	}()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.SetRow("A1", []interface{}{nil, nil, Cell{Value: "foo"}}))
	assert.NoError(t, streamWriter.Flush())
	ws, err := file.workSheetReader("Sheet1")
	assert.NoError(t, err)
	assert.NotEqual(t, ws.SheetData.Row[0].C[0].XMLName.Local, "c")
}

func TestStreamSetRowWithStyle(t *testing.T) {
	file := NewFile()
	defer func() {
		assert.NoError(t, file.Close())
	}()
	grayStyleID, err := file.NewStyle(&Style{Font: &Font{Color: "777777"}})
	assert.NoError(t, err)
	blueStyleID, err := file.NewStyle(&Style{Font: &Font{Color: "0000FF"}})
	assert.NoError(t, err)

	sheetName := "Sheet1"
	streamWriter, err := file.NewStreamWriter(sheetName)
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.SetColStyle(1, 1, grayStyleID))
	assert.NoError(t, streamWriter.SetColStyle(3, 3, blueStyleID))
	assert.NoError(t, streamWriter.SetRow("A1", []interface{}{
		"A1",
		Cell{Value: "B1"},
		&Cell{Value: "C1"},
		Cell{StyleID: blueStyleID, Value: "D1"},
		&Cell{StyleID: blueStyleID, Value: "E1"},
	}, RowOpts{StyleID: grayStyleID}))
	assert.NoError(t, streamWriter.SetRow("A2", []interface{}{
		"A2",
		Cell{Value: "B2"},
		&Cell{Value: "C2"},
		Cell{StyleID: grayStyleID, Value: "D2"},
		&Cell{StyleID: blueStyleID, Value: "E2"},
	}))
	assert.NoError(t, streamWriter.Flush())

	ws, err := file.workSheetReader(sheetName)
	assert.NoError(t, err)
	for colIdx, expected := range []int{grayStyleID, grayStyleID, grayStyleID, blueStyleID, blueStyleID} {
		assert.Equal(t, expected, ws.SheetData.Row[0].C[colIdx].S)
	}
	for colIdx, expected := range []int{grayStyleID, 0, blueStyleID, grayStyleID, blueStyleID} {
		assert.Equal(t, expected, ws.SheetData.Row[1].C[colIdx].S)
	}
}

func TestStreamSetCellValFunc(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()
	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	c := &xlsxC{}
	for _, val := range []interface{}{
		128,
		int8(-128),
		int16(-32768),
		int32(-2147483648),
		int64(-9223372036854775808),
		uint(128),
		uint8(255),
		uint16(65535),
		uint32(4294967295),
		uint64(18446744073709551615),
		float32(100.1588),
		100.1588,
		" Hello",
		[]byte(" Hello"),
		time.Now().UTC(),
		time.Duration(1e13),
		true,
		nil,
		complex64(5 + 10i),
	} {
		assert.NoError(t, sw.setCellValFunc(c, val))
	}
}

func TestStreamWriterOutlineLevel(t *testing.T) {
	file := NewFile()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)

	// Test set outlineLevel in row
	assert.NoError(t, streamWriter.SetRow("A1", nil, RowOpts{OutlineLevel: 1}))
	assert.NoError(t, streamWriter.SetRow("A2", nil, RowOpts{OutlineLevel: 7}))
	assert.ErrorIs(t, ErrOutlineLevel, streamWriter.SetRow("A3", nil, RowOpts{OutlineLevel: 8}))

	assert.NoError(t, streamWriter.Flush())
	// Save spreadsheet by the given path
	assert.NoError(t, file.SaveAs(filepath.Join("test", "TestStreamWriterSetRowOutlineLevel.xlsx")))

	file, err = OpenFile(filepath.Join("test", "TestStreamWriterSetRowOutlineLevel.xlsx"))
	assert.NoError(t, err)
	for rowIdx, expected := range []uint8{1, 7, 0} {
		level, err := file.GetRowOutlineLevel("Sheet1", rowIdx+1)
		assert.NoError(t, err)
		assert.Equal(t, expected, level)
	}
	assert.NoError(t, file.Close())
}

func TestStreamWriterReader(t *testing.T) {
	var (
		err error
		sw  = StreamWriter{
			rawData: bufferedWriter{},
		}
	)
	sw.rawData.tmp, err = os.CreateTemp(os.TempDir(), "excelize-")
	assert.NoError(t, err)
	assert.NoError(t, sw.rawData.tmp.Close())
	// Test reader stat a closed temp file
	_, err = sw.rawData.Reader()
	assert.Error(t, err)
	_, err = sw.getRowValues(1, 1, 1)
	assert.Error(t, err)
	assert.NoError(t, os.Remove(sw.rawData.tmp.Name()))

	sw = StreamWriter{
		file:    NewFile(),
		rawData: bufferedWriter{},
	}
	// Test getRowValues without expected row
	sw.rawData.buf.WriteString("<worksheet><row r=\"1\"><c r=\"B1\"></c></row><worksheet/>")
	_, err = sw.getRowValues(1, 1, 1)
	assert.NoError(t, err)
	sw.rawData.buf.Reset()
	// Test getRowValues with illegal cell reference
	sw.rawData.buf.WriteString("<worksheet><row r=\"1\"><c r=\"A\"></c></row><worksheet/>")
	_, err = sw.getRowValues(1, 1, 1)
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), err)
	sw.rawData.buf.Reset()
	// Test getRowValues with invalid c element characters
	sw.rawData.buf.WriteString("<worksheet><row r=\"1\"><c></row><worksheet/>")
	_, err = sw.getRowValues(1, 1, 1)
	assert.EqualError(t, err, "XML syntax error on line 1: element <c> closed by </row>")
	sw.rawData.buf.Reset()
}

func TestStreamWriterGetRowElement(t *testing.T) {
	// Test get row element without r attribute
	dec := xml.NewDecoder(strings.NewReader("<row ht=\"0\" />"))
	for {
		token, err := dec.Token()
		if err == io.EOF {
			break
		}
		_, ok := getRowElement(token, 0)
		assert.False(t, ok)
	}
}

func TestBufferedWriterWriteInt(t *testing.T) {
	// In-memory path
	bw := &bufferedWriter{flushSize: StreamChunkSize, bioSize: StreamingBufSizeDefault}
	bw.WriteInt(42)
	bw.WriteInt(-1234567890)
	assert.Equal(t, "42-1234567890", bw.buf.String())
	assert.Equal(t, int64(len("42-1234567890")), bw.written)

	// Temp file (bio) path
	bw2 := &bufferedWriter{flushSize: 1, bioSize: 4096}
	bw2.WriteString("x") // trigger sync
	_ = bw2.Sync()
	assert.NotNil(t, bw2.bio)
	bw2.WriteInt(99)
	_ = bw2.Flush()
	bw2.tmp.Seek(0, 0)
	data, _ := io.ReadAll(bw2.tmp)
	assert.Contains(t, string(data), "99")
	bw2.Close()
}

func TestBufferedWriterWriteUint(t *testing.T) {
	bw := &bufferedWriter{flushSize: StreamChunkSize, bioSize: StreamingBufSizeDefault}
	bw.WriteUint(12345)
	assert.Equal(t, "12345", bw.buf.String())

	// bio path
	bw2 := &bufferedWriter{flushSize: 1, bioSize: 4096}
	bw2.WriteString("x")
	_ = bw2.Sync()
	bw2.WriteUint(67890)
	_ = bw2.Flush()
	bw2.tmp.Seek(0, 0)
	data, _ := io.ReadAll(bw2.tmp)
	assert.Contains(t, string(data), "67890")
	bw2.Close()
}

func TestBufferedWriterWriteFloat(t *testing.T) {
	bw := &bufferedWriter{flushSize: StreamChunkSize, bioSize: StreamingBufSizeDefault}
	bw.WriteFloat(3.14, 'f', 2, 64)
	assert.Equal(t, "3.14", bw.buf.String())

	// bio path
	bw2 := &bufferedWriter{flushSize: 1, bioSize: 4096}
	bw2.WriteString("x")
	_ = bw2.Sync()
	bw2.WriteFloat(2.72, 'f', 2, 64)
	_ = bw2.Flush()
	bw2.tmp.Seek(0, 0)
	data, _ := io.ReadAll(bw2.tmp)
	assert.Contains(t, string(data), "2.72")
	bw2.Close()
}

func TestBufferedWriterBytes(t *testing.T) {
	// In-memory: returns buffer bytes
	bw := &bufferedWriter{flushSize: StreamChunkSize, bioSize: StreamingBufSizeDefault}
	bw.WriteString("hello")
	assert.Equal(t, []byte("hello"), bw.Bytes())

	// After temp file creation: returns nil
	bw2 := &bufferedWriter{flushSize: 1, bioSize: 4096}
	bw2.WriteString("x")
	_ = bw2.Sync()
	assert.Nil(t, bw2.Bytes())
	bw2.Close()
}

func TestBufferedWriterWriteAt(t *testing.T) {
	// In-memory WriteAt
	bw := &bufferedWriter{flushSize: StreamChunkSize, bioSize: StreamingBufSizeDefault}
	bw.WriteString("AAABBBCCC")
	err := bw.WriteAt([]byte("XXX"), 3)
	assert.NoError(t, err)
	assert.Equal(t, "AAAXXXCCC", bw.buf.String())

	// In-memory WriteAt out of bounds
	err = bw.WriteAt([]byte("TOOLONG"), 5)
	assert.Error(t, err)

	// Temp file WriteAt
	bw2 := &bufferedWriter{flushSize: 1, bioSize: 4096}
	bw2.WriteString("AAABBBCCC")
	_ = bw2.Sync()
	err = bw2.WriteAt([]byte("YYY"), 3)
	assert.NoError(t, err)
	// Verify by reading the file back
	var readBuf bytes.Buffer
	_, _ = bw2.CopyTo(&readBuf)
	assert.Equal(t, "AAAYYYCC", readBuf.String()[:8])
	bw2.Close()
}

func TestBufferedWriterCopyTo(t *testing.T) {
	// In-memory CopyTo
	bw := &bufferedWriter{flushSize: StreamChunkSize, bioSize: StreamingBufSizeDefault}
	bw.WriteString("hello world")
	var dst bytes.Buffer
	n, err := bw.CopyTo(&dst)
	assert.NoError(t, err)
	assert.Equal(t, int64(11), n)
	assert.Equal(t, "hello world", dst.String())

	// Temp file CopyTo
	bw2 := &bufferedWriter{flushSize: 1, bioSize: 4096}
	bw2.WriteString("file data here")
	_ = bw2.Sync()
	bw2.WriteString(" more") // this goes through bio
	var dst2 bytes.Buffer
	n2, err := bw2.CopyTo(&dst2)
	assert.NoError(t, err)
	assert.Equal(t, int64(19), n2)
	assert.Equal(t, "file data here more", dst2.String())
	bw2.Close()

	// Temp file CopyTo with large bioSize (> 256KB)
	bw3 := &bufferedWriter{flushSize: 1, bioSize: 512 * 1024}
	bw3.WriteString("large buffer test")
	_ = bw3.Sync()
	var dst3 bytes.Buffer
	n3, err := bw3.CopyTo(&dst3)
	assert.NoError(t, err)
	assert.Equal(t, int64(17), n3)
	assert.Equal(t, "large buffer test", dst3.String())
	bw3.Close()
}

func TestBufferedWriterReset(t *testing.T) {
	// Reset in-memory only
	bw := &bufferedWriter{flushSize: StreamChunkSize, bioSize: StreamingBufSizeDefault}
	bw.WriteString("data")
	bw.Reset()
	assert.Equal(t, 0, bw.buf.Len())

	// Reset after temp file creation
	bw2 := &bufferedWriter{flushSize: 1, bioSize: 4096}
	bw2.WriteString("data")
	_ = bw2.Sync()
	assert.NotNil(t, bw2.bio)
	bw2.Reset()
	assert.Nil(t, bw2.bio)
	assert.Equal(t, 0, bw2.buf.Len())
	bw2.Close()
}

func TestNewStreamWriterOptions(t *testing.T) {
	// Test StreamingChunkSize = -1 (never spill)
	f := NewFile()
	defer f.Close()
	f.options.StreamingChunkSize = -1
	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.True(t, sw.rawData.flushSize > StreamChunkSize)

	// Test StreamingBufSize custom value
	f2 := NewFile()
	defer f2.Close()
	f2.options.StreamingBufSize = 64 * 1024
	sw2, err := f2.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, 64*1024, sw2.rawData.bioSize)

	// Test StreamingChunkSize positive custom value
	f3 := NewFile()
	defer f3.Close()
	f3.options.StreamingChunkSize = 1024
	sw3, err := f3.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, 1024, sw3.rawData.flushSize)
}

func TestBufferedWriterWriteAtFlushError(t *testing.T) {
	// Test WriteAt temp file path where Flush fails (line 901)
	bw := &bufferedWriter{flushSize: 1, bioSize: 4096}
	bw.WriteString("AAABBBCCC")
	_ = bw.Sync()
	// Write more data so bio has unflushed content
	bw.bio.WriteString("extra")
	// Close the temp file to cause Flush (bio.Flush) to fail
	bw.tmp.Close()
	err := bw.WriteAt([]byte("YYY"), 3)
	assert.Error(t, err)
}

func TestBufferedWriterCopyToFlushError(t *testing.T) {
	// Test CopyTo temp file path where Flush fails (line 915)
	bw := &bufferedWriter{flushSize: 1, bioSize: 4096}
	bw.WriteString("test data")
	_ = bw.Sync()
	bw.WriteString(" more")
	// Close file to cause Flush to fail
	bw.tmp.Close()
	var dst bytes.Buffer
	_, err := bw.CopyTo(&dst)
	assert.Error(t, err)
}

func TestBufferedWriterCopyToSeekError(t *testing.T) {
	// Test CopyTo temp file path where Seek fails (line 918)
	bw := &bufferedWriter{flushSize: 1, bioSize: 4096}
	bw.WriteString("test data")
	_ = bw.Sync()
	// Close file so Flush succeeds (bio is nil after sync with no writes) but Seek fails
	// We need bio to be nil so Flush() is a no-op, then Seek will fail on closed file
	bw.bio = nil
	bw.tmp.Close()
	var dst bytes.Buffer
	_, err := bw.CopyTo(&dst)
	assert.Error(t, err)
}

func TestBufferedWriterSyncWriteToError(t *testing.T) {
	// Test Sync where buf.WriteTo(tmp) fails (line 970)
	bw := &bufferedWriter{flushSize: 1, bioSize: 4096}
	bw.WriteString("initial")
	// Sync to create temp file
	_ = bw.Sync()
	// Now reset state to have data in buf and tmp exists but is closed
	bw.bio = nil
	bw.buf.WriteString("more data")
	bw.tmp.Close()
	err := bw.Sync()
	assert.Error(t, err)
}
