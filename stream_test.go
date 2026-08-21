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
	f := NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			b.Error(err)
		}
	}()
	row := make([]interface{}, 10)
	for colID := 0; colID < 10; colID++ {
		row[colID] = colID
	}

	for b.Loop() {
		sw, _ := f.NewStreamWriter("Sheet1")
		for rowID := 10; rowID <= 110; rowID++ {
			cell, _ := CoordinatesToCellName(1, rowID)
			_ = sw.SetRow(cell, row)
		}
	}

	b.ReportAllocs()
}

func TestStreamWriter(t *testing.T) {
	f := NewFile()
	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)

	// Test max characters in a cell
	row := make([]interface{}, 1)
	row[0] = strings.Repeat("c", TotalCellChars+2)
	assert.NoError(t, sw.SetRow("A1", row))

	// Test leading and ending space(s) character characters in a cell
	row = make([]interface{}, 1)
	row[0] = " characters"
	assert.NoError(t, sw.SetRow("A2", row))

	row = make([]interface{}, 1)
	row[0] = []byte("Word")
	assert.NoError(t, sw.SetRow("A3", row))

	// Test set cell with style and rich text
	styleID, err := f.NewStyle(&Style{Font: &Font{Color: "777777"}})
	assert.NoError(t, err)
	assert.NoError(t, sw.SetRow("A4", []interface{}{
		Cell{StyleID: styleID},
		Cell{Formula: "SUM(A10,B10)", Value: " preserve space "},
	},
		RowOpts{Height: 45, StyleID: styleID}))
	assert.NoError(t, sw.SetRow("A5", []interface{}{
		&Cell{StyleID: styleID, Value: "cell <>&'\""},
		&Cell{Formula: "SUM(A10,B10)"},
		[]RichTextRun{
			{Text: "Rich ", Font: &Font{Color: "2354E8"}},
			{Text: "Text", Font: &Font{Color: "E83723"}},
		},
	}))
	assert.NoError(t, sw.SetRow("A6", []interface{}{time.Now()}))
	assert.NoError(t, sw.SetRow("A7", nil, RowOpts{Height: 20, Hidden: true, StyleID: styleID}))
	assert.Equal(t, ErrMaxRowHeight, sw.SetRow("A8", nil, RowOpts{Height: MaxRowHeight + 1}))

	assert.NoError(t, sw.SetRow("A9", []interface{}{math.NaN(), math.Inf(0), math.Inf(-1)}))

	for rowID := 10; rowID <= 51200; rowID++ {
		row := make([]interface{}, 50)
		for colID := 0; colID < 50; colID++ {
			row[colID] = rand.Intn(640000)
		}
		cell, _ := CoordinatesToCellName(1, rowID)
		assert.NoError(t, sw.SetRow(cell, row))
	}

	assert.NoError(t, sw.Flush())
	// Save spreadsheet by the given path
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestStreamWriter.xlsx")))

	// Test set cell column overflow
	assert.ErrorIs(t, sw.SetRow("XFD51201", []interface{}{"A", "B", "C"}), ErrColumnNumber)
	assert.NoError(t, f.Close())

	// Test close temporary file error
	f = NewFile(Options{TmpDir: os.TempDir()})
	sw, err = f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	for rowID := 10; rowID <= 25600; rowID++ {
		row := make([]interface{}, 50)
		for colID := 0; colID < 50; colID++ {
			row[colID] = rand.Intn(640000)
		}
		cell, _ := CoordinatesToCellName(1, rowID)
		assert.NoError(t, sw.SetRow(cell, row))
	}
	assert.NoError(t, sw.rawData.Close())
	assert.Error(t, sw.Flush())

	sw.rawData.tmp, err = os.CreateTemp(os.TempDir(), "excelize-")
	assert.NoError(t, err)
	_, err = sw.rawData.Reader()
	assert.NoError(t, err)
	assert.NoError(t, sw.rawData.tmp.Close())
	assert.NoError(t, os.Remove(sw.rawData.tmp.Name()))

	// Test create stream writer with unsupported charset
	f = NewFile()
	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", MacintoshCyrillicCharset)
	_, err = f.NewStreamWriter("Sheet1")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())

	// Test read cell
	f = NewFile()
	sw, err = f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, sw.SetRow("A1", []interface{}{Cell{StyleID: styleID, Value: "Data"}}))
	assert.NoError(t, sw.Flush())
	cellValue, err := f.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "Data", cellValue)

	// Test stream reader for a worksheet with huge amounts of data
	f, err = OpenFile(filepath.Join("test", "TestStreamWriter.xlsx"))
	assert.NoError(t, err)
	rows, err := f.Rows("Sheet1")
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
	assert.NoError(t, f.SaveAs(filepath.Join("test", "EncryptionTestStreamWriter.xlsx"), Options{Password: "password"}))
	assert.NoError(t, f.Close())
}

func TestStreamSetColVisible(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()
	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, sw.SetColVisible(3, 2, false))
	assert.Equal(t, ErrColumnNumber, sw.SetColVisible(0, 3, false))
	assert.Equal(t, ErrColumnNumber, sw.SetColVisible(MaxColumns+1, 3, false))
	assert.NoError(t, sw.SetRow("A1", []interface{}{"A", "B", "C"}))
	assert.Equal(t, newStreamSetRowOrderError("SetColVisible"), sw.SetColVisible(2, 3, false))
	assert.NoError(t, sw.Flush())
}

func TestStreamSetColOutlineLevel(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()
	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, sw.SetColOutlineLevel(4, 2))
	assert.Equal(t, ErrOutlineLevel, sw.SetColOutlineLevel(4, 0))
	assert.Equal(t, ErrOutlineLevel, sw.SetColOutlineLevel(4, 8))
	assert.Equal(t, ErrColumnNumber, sw.SetColOutlineLevel(0, 2))
	assert.Equal(t, ErrColumnNumber, sw.SetColOutlineLevel(MaxColumns+1, 2))
	assert.NoError(t, sw.SetRow("A1", []interface{}{"A", "B", "C"}))
	assert.Equal(t, newStreamSetRowOrderError("SetColOutlineLevel"), sw.SetColOutlineLevel(4, 2))
	assert.NoError(t, sw.Flush())
}

func TestStreamSetColStyle(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()
	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, sw.SetColStyle(3, 2, 0))
	assert.Equal(t, ErrColumnNumber, sw.SetColStyle(0, 3, 20))
	assert.Equal(t, ErrColumnNumber, sw.SetColStyle(MaxColumns+1, 3, 20))
	assert.Equal(t, newInvalidStyleID(2), sw.SetColStyle(1, 3, 2))
	assert.NoError(t, sw.SetRow("A1", []interface{}{"A", "B", "C"}))
	assert.Equal(t, newStreamSetRowOrderError("SetColStyle"), sw.SetColStyle(2, 3, 0))

	f = NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()
	// Test set column style with unsupported charset style sheet
	f.Styles = nil
	f.Pkg.Store(defaultXMLPathStyles, MacintoshCyrillicCharset)
	sw, err = f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.EqualError(t, sw.SetColStyle(3, 2, 0), "XML syntax error on line 1: invalid UTF-8")
}

func TestStreamSetColWidth(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()
	styleID, err := f.NewStyle(&Style{
		Fill: Fill{Type: "pattern", Color: []string{"E0EBF5"}, Pattern: 1},
	})
	if err != nil {
		fmt.Println(err)
	}
	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, sw.SetColWidth(3, 2, 20))
	assert.NoError(t, sw.SetColStyle(3, 2, styleID))
	assert.Equal(t, ErrColumnNumber, sw.SetColWidth(0, 3, 20))
	assert.Equal(t, ErrColumnNumber, sw.SetColWidth(MaxColumns+1, 3, 20))
	assert.Equal(t, ErrColumnWidth, sw.SetColWidth(1, 3, MaxColumnWidth+1))
	assert.NoError(t, sw.SetRow("A1", []interface{}{"A", "B", "C"}))
	assert.Equal(t, newStreamSetRowOrderError("SetColWidth"), sw.SetColWidth(2, 3, 20))
	assert.NoError(t, sw.Flush())
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
	sw, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, sw.SetPanes(paneOpts))
	assert.Equal(t, ErrParameterInvalid, sw.SetPanes(nil))
	assert.NoError(t, sw.SetRow("A1", []interface{}{"A", "B", "C"}))
	assert.Equal(t, newStreamSetRowOrderError("SetPanes"), sw.SetPanes(paneOpts))
}

func TestStreamTable(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()
	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	// Test add table without table header
	assert.EqualError(t, sw.AddTable(&Table{Range: "A1:C2"}), "XML syntax error on line 2: unexpected EOF")
	// Write some rows. We want enough rows to force a temp file (>16MB)
	assert.NoError(t, sw.SetRow("A1", []interface{}{"A", "B", "C"}))
	row := []interface{}{1, 2, 3}
	for r := 2; r < 10000; r++ {
		assert.NoError(t, sw.SetRow(fmt.Sprintf("A%d", r), row))
	}

	// Write a table
	assert.NoError(t, sw.AddTable(&Table{Range: "A1:C2"}))
	assert.NoError(t, sw.Flush())

	// Verify the table has names
	var table xlsxTable
	val, ok := f.Pkg.Load("xl/tables/table1.xml")
	assert.True(t, ok)
	assert.NoError(t, xml.Unmarshal(val.([]byte), &table))
	assert.Equal(t, "A", table.TableColumns.TableColumn[0].Name)
	assert.Equal(t, "B", table.TableColumns.TableColumn[1].Name)
	assert.Equal(t, "C", table.TableColumns.TableColumn[2].Name)

	assert.NoError(t, sw.AddTable(&Table{Range: "A1:C1"}))

	// Test add table with illegal cell reference
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), sw.AddTable(&Table{Range: "A:B1"}))
	assert.Equal(t, newCellNameToCoordinatesError("B", newInvalidCellNameError("B")), sw.AddTable(&Table{Range: "A1:B"}))
	// Test add table with invalid table name
	assert.Equal(t, newInvalidNameError("1Table"), sw.AddTable(&Table{Range: "A:B1", Name: "1Table"}))
	// Test add table with row number exceeds maximum limit
	assert.Equal(t, ErrMaxRows, sw.AddTable(&Table{Range: "A1048576:C1048576"}))
	// Test add table with unsupported charset content types
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, sw.AddTable(&Table{Range: "A1:C2"}), "XML syntax error on line 1: invalid UTF-8")
}

func TestStreamMergeCells(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()
	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, sw.MergeCell("A1", "D1"))
	// Test merge cells with illegal cell reference
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), sw.MergeCell("A", "D1"))
	assert.NoError(t, sw.Flush())
	// Save spreadsheet by the given path
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestStreamMergeCells.xlsx")))
}

func TestStreamInsertPageBreak(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()
	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, sw.InsertPageBreak("A1"))
	assert.NoError(t, sw.Flush())
	// Save spreadsheet by the given path
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestStreamInsertPageBreak.xlsx")))
}

func TestNewStreamWriter(t *testing.T) {
	// Test error exceptions
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()
	_, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	_, err = f.NewStreamWriter("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Test new stream write with invalid sheet name
	_, err = f.NewStreamWriter("Sheet:1")
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
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()
	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), sw.SetRow("A", []interface{}{}))
	// Test set row with non-ascending row number
	assert.NoError(t, sw.SetRow("A1", []interface{}{}))
	assert.Equal(t, newStreamSetRowError(1), sw.SetRow("A1", []interface{}{}))
	// Test set row with unsupported charset workbook
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	assert.EqualError(t, sw.SetRow("A2", []interface{}{time.Now()}), "XML syntax error on line 1: invalid UTF-8")
}

func TestStreamSetRowNilValues(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()
	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, sw.SetRow("A1", []interface{}{nil, nil, Cell{Value: "foo"}}))
	assert.NoError(t, sw.Flush())
	ws, err := f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	assert.NotEqual(t, ws.SheetData.Row[0].C[0].XMLName.Local, "c")
}

func TestStreamSetRowWithStyle(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()
	grayStyleID, err := f.NewStyle(&Style{Font: &Font{Color: "777777"}})
	assert.NoError(t, err)
	blueStyleID, err := f.NewStyle(&Style{Font: &Font{Color: "0000FF"}})
	assert.NoError(t, err)

	sheetName := "Sheet1"
	sw, err := f.NewStreamWriter(sheetName)
	assert.NoError(t, err)
	assert.NoError(t, sw.SetColStyle(1, 1, grayStyleID))
	assert.NoError(t, sw.SetColStyle(3, 3, blueStyleID))
	assert.NoError(t, sw.SetRow("A1", []interface{}{
		"A1",
		Cell{Value: "B1"},
		&Cell{Value: "C1"},
		Cell{StyleID: blueStyleID, Value: "D1"},
		&Cell{StyleID: blueStyleID, Value: "E1"},
	}, RowOpts{StyleID: grayStyleID}))
	assert.NoError(t, sw.SetRow("A2", []interface{}{
		"A2",
		Cell{Value: "B2"},
		&Cell{Value: "C2"},
		Cell{StyleID: grayStyleID, Value: "D2"},
		&Cell{StyleID: blueStyleID, Value: "E2"},
	}))
	assert.NoError(t, sw.Flush())

	ws, err := f.workSheetReader(sheetName)
	assert.NoError(t, err)
	for colIdx, expected := range []int{grayStyleID, grayStyleID, grayStyleID, blueStyleID, blueStyleID} {
		assert.Equal(t, expected, ws.SheetData.Row[0].C[colIdx].S)
	}
	for colIdx, expected := range []int{grayStyleID, 0, blueStyleID, grayStyleID, blueStyleID} {
		assert.Equal(t, expected, ws.SheetData.Row[1].C[colIdx].S)
	}
}

func TestStreamWriterDurationFormat(t *testing.T) {
	f := NewFile()
	savePath := filepath.Join("test", "TestStreamWriterDuration.xlsx")
	val := 25*time.Hour + 30*time.Minute
	styleID, err := f.NewStyle(&Style{Font: &Font{Color: "777777"}})
	assert.NoError(t, err)
	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, sw.SetRow("A1", []interface{}{val, Cell{StyleID: styleID, Value: val}}))
	assert.NoError(t, sw.Flush())
	assert.NoError(t, f.SaveAs(savePath))
	assert.NoError(t, f.Close())
	f, err = OpenFile(savePath)
	assert.NoError(t, err)
	cell, err := f.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "25:30:00", cell)
	cell, err = f.GetCellValue("Sheet1", "B1")
	assert.NoError(t, err)
	assert.Equal(t, "1.0625", cell)
	assert.NoError(t, f.Close())
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
	f := NewFile()
	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)

	// Test set outlineLevel in row
	assert.NoError(t, sw.SetRow("A1", nil, RowOpts{OutlineLevel: 1}))
	assert.NoError(t, sw.SetRow("A2", nil, RowOpts{OutlineLevel: 7}))
	assert.ErrorIs(t, ErrOutlineLevel, sw.SetRow("A3", nil, RowOpts{OutlineLevel: 8}))

	assert.NoError(t, sw.Flush())
	// Save spreadsheet by the given path
	savePath := filepath.Join("test", "TestStreamWriterSetRowOutlineLevel.xlsx")
	assert.NoError(t, f.SaveAs(savePath))

	f, err = OpenFile(savePath)
	assert.NoError(t, err)
	for rowIdx, expected := range []uint8{1, 7, 0} {
		level, err := f.GetRowOutlineLevel("Sheet1", rowIdx+1)
		assert.NoError(t, err)
		assert.Equal(t, expected, level)
	}
	assert.NoError(t, f.Close())
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

func TestStreamWriterUseSharedStrings(t *testing.T) {
	// Test basic UseSharedStrings functionality
	f := NewFile(Options{UseSharedStrings: true})
	defer f.Close()

	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.True(t, sw.useSharedStrings)
	assert.NotNil(t, sw.sharedStringsMap)

	// Write rows with duplicate strings
	assert.NoError(t, sw.SetRow("A1", []interface{}{"hello", "world", "hello"}))
	assert.NoError(t, sw.SetRow("A2", []interface{}{"world", "new", "hello"}))

	assert.NoError(t, sw.Flush())

	// Verify shared strings were created
	data, ok := f.Pkg.Load(defaultXMLPathSharedStrings)
	assert.True(t, ok)
	sstXML := string(data.([]byte))
	assert.Contains(t, sstXML, `<si><t>hello</t></si>`)
	assert.Contains(t, sstXML, `<si><t>world</t></si>`)
	assert.Contains(t, sstXML, `<si><t>new</t></si>`)
	assert.Contains(t, sstXML, `uniqueCount="3"`)

	// Save and re-read to verify file is valid
	tmpPath := filepath.Join(t.TempDir(), "shared_strings.xlsx")
	assert.NoError(t, f.SaveAs(tmpPath))

	f2, err := OpenFile(tmpPath)
	assert.NoError(t, err)
	defer f2.Close()

	val, err := f2.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "hello", val)

	val, err = f2.GetCellValue("Sheet1", "B1")
	assert.NoError(t, err)
	assert.Equal(t, "world", val)

	val, err = f2.GetCellValue("Sheet1", "C1")
	assert.NoError(t, err)
	assert.Equal(t, "hello", val)

	val, err = f2.GetCellValue("Sheet1", "B2")
	assert.NoError(t, err)
	assert.Equal(t, "new", val)
}

func TestStreamWriterSharedStringsPreserveSpace(t *testing.T) {
	// Test that strings with leading/trailing whitespace get xml:space="preserve"
	f := NewFile(Options{UseSharedStrings: true})
	defer f.Close()

	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)

	assert.NoError(t, sw.SetRow("A1", []interface{}{" leading", "trailing ", "\tnewline\n"}))
	assert.NoError(t, sw.Flush())

	data, ok := f.Pkg.Load(defaultXMLPathSharedStrings)
	assert.True(t, ok)
	sstXML := string(data.([]byte))
	assert.Contains(t, sstXML, `xml:space="preserve"`)
}

func TestStreamWriterSharedStringsDedup(t *testing.T) {
	// Test that getOrAddSharedString deduplicates correctly
	f := NewFile(Options{UseSharedStrings: true})
	defer f.Close()

	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)

	// Same string should return same index
	idx1 := sw.getOrAddSharedString("test")
	idx2 := sw.getOrAddSharedString("test")
	idx3 := sw.getOrAddSharedString("other")
	assert.Equal(t, 0, idx1)
	assert.Equal(t, 0, idx2)
	assert.Equal(t, 1, idx3)
	assert.Equal(t, 2, len(sw.sharedStrings))
}

func TestStreamWriterSharedStringsDisabled(t *testing.T) {
	// Test that without UseSharedStrings, inline strings are used
	f := NewFile()
	defer f.Close()

	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.False(t, sw.useSharedStrings)

	assert.NoError(t, sw.SetRow("A1", []interface{}{"hello"}))
	assert.NoError(t, sw.Flush())

	// No shared strings file should be created
	_, ok := f.Pkg.Load(defaultXMLPathSharedStrings)
	assert.False(t, ok)
}

func TestStreamWriterSharedStringsEmptyStrings(t *testing.T) {
	// Test with empty string values (no shared strings generated)
	f := NewFile(Options{UseSharedStrings: true})
	defer f.Close()

	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)

	// Write only numeric values - no shared strings
	assert.NoError(t, sw.SetRow("A1", []interface{}{42, 3.14, true}))
	assert.NoError(t, sw.Flush())

	// No shared strings should be written (empty table)
	_, ok := f.Pkg.Load(defaultXMLPathSharedStrings)
	assert.False(t, ok)
}

func TestStreamWriterSharedStringsRelationship(t *testing.T) {
	// Test that writeSharedStrings adds content type and relationship
	f := NewFile(Options{UseSharedStrings: true})
	defer f.Close()

	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)

	assert.NoError(t, sw.SetRow("A1", []interface{}{"test"}))
	assert.NoError(t, sw.Flush())

	// Call Flush again with another stream writer to verify relationship dedup
	sw2, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, sw2.SetRow("A1", []interface{}{"test2"}))
	assert.NoError(t, sw2.Flush())
}

func TestStreamWriterSharedStringsError(t *testing.T) {
	// Test writeSharedStrings error path (corrupted content types)
	f := NewFile(Options{UseSharedStrings: true})
	defer f.Close()

	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)

	assert.NoError(t, sw.SetRow("A1", []interface{}{"trigger_sst"}))

	// Corrupt content types with invalid XML that triggers a decode error
	// (unclosed element causes syntax error, not EOF)
	f.Pkg.Store(defaultXMLPathContentTypes, []byte(`<Types xmlns="x"><Override PartName="/xl`))
	f.ContentTypes = nil

	err = sw.Flush()
	assert.Error(t, err)
}
