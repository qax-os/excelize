package excelize

import (
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
	file = NewFile()
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
	assert.Equal(t, ErrStreamSetColStyle, streamWriter.SetColStyle(2, 3, 0))

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
	assert.Equal(t, ErrStreamSetColWidth, streamWriter.SetColWidth(2, 3, 20))
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
	assert.Equal(t, ErrStreamSetPanes, streamWriter.SetPanes(paneOpts))
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
	streamWriter.Flush()
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

	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.SetRow("A1", []interface{}{
		"value1",
		Cell{Value: "value2"},
		&Cell{Value: "value2"},
		Cell{StyleID: blueStyleID, Value: "value3"},
		&Cell{StyleID: blueStyleID, Value: "value3"},
	}, RowOpts{StyleID: grayStyleID}))
	assert.NoError(t, streamWriter.Flush())

	ws, err := file.workSheetReader("Sheet1")
	assert.NoError(t, err)
	for colIdx, expected := range []int{grayStyleID, grayStyleID, grayStyleID, blueStyleID, blueStyleID} {
		assert.Equal(t, expected, ws.SheetData.Row[0].C[colIdx].S)
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
	os.Remove(sw.rawData.tmp.Name())

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
