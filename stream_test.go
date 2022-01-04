package excelize

import (
	"encoding/xml"
	"fmt"
	"io/ioutil"
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

	// Test max characters in a cell.
	row := make([]interface{}, 1)
	row[0] = strings.Repeat("c", TotalCellChars+2)
	assert.NoError(t, streamWriter.SetRow("A1", row))

	// Test leading and ending space(s) character characters in a cell.
	row = make([]interface{}, 1)
	row[0] = " characters"
	assert.NoError(t, streamWriter.SetRow("A2", row))

	row = make([]interface{}, 1)
	row[0] = []byte("Word")
	assert.NoError(t, streamWriter.SetRow("A3", row))

	// Test set cell with style.
	styleID, err := file.NewStyle(&Style{Font: &Font{Color: "#777777"}})
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.SetRow("A4", []interface{}{Cell{StyleID: styleID}, Cell{Formula: "SUM(A10,B10)"}}), RowOpts{Height: 45, StyleID: styleID})
	assert.NoError(t, streamWriter.SetRow("A5", []interface{}{&Cell{StyleID: styleID, Value: "cell"}, &Cell{Formula: "SUM(A10,B10)"}}))
	assert.NoError(t, streamWriter.SetRow("A6", []interface{}{time.Now()}))
	assert.NoError(t, streamWriter.SetRow("A7", nil, RowOpts{Height: 20, Hidden: true, StyleID: styleID}))
	assert.EqualError(t, streamWriter.SetRow("A7", nil, RowOpts{Height: MaxRowHeight + 1}), ErrMaxRowHeight.Error())

	for rowID := 10; rowID <= 51200; rowID++ {
		row := make([]interface{}, 50)
		for colID := 0; colID < 50; colID++ {
			row[colID] = rand.Intn(640000)
		}
		cell, _ := CoordinatesToCellName(1, rowID)
		assert.NoError(t, streamWriter.SetRow(cell, row))
	}

	assert.NoError(t, streamWriter.Flush())
	// Save spreadsheet by the given path.
	assert.NoError(t, file.SaveAs(filepath.Join("test", "TestStreamWriter.xlsx")))

	// Test set cell column overflow.
	assert.EqualError(t, streamWriter.SetRow("XFD1", []interface{}{"A", "B", "C"}), ErrColumnNumber.Error())

	// Test close temporary file error.
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

	streamWriter.rawData.tmp, err = ioutil.TempFile(os.TempDir(), "excelize-")
	assert.NoError(t, err)
	_, err = streamWriter.rawData.Reader()
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.rawData.tmp.Close())
	assert.NoError(t, os.Remove(streamWriter.rawData.tmp.Name()))

	// Test unsupported charset
	file = NewFile()
	file.Sheet.Delete("xl/worksheets/sheet1.xml")
	file.Pkg.Store("xl/worksheets/sheet1.xml", MacintoshCyrillicCharset)
	_, err = file.NewStreamWriter("Sheet1")
	assert.EqualError(t, err, "xml decode error: XML syntax error on line 1: invalid UTF-8")

	// Test read cell.
	file = NewFile()
	streamWriter, err = file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.SetRow("A1", []interface{}{Cell{StyleID: styleID, Value: "Data"}}))
	assert.NoError(t, streamWriter.Flush())
	cellValue, err := file.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "Data", cellValue)

	// Test stream reader for a worksheet with huge amounts of data.
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
	assert.Equal(t, 2559558, cells)
	assert.NoError(t, file.Close())
}

func TestStreamSetColWidth(t *testing.T) {
	file := NewFile()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.SetColWidth(3, 2, 20))
	assert.EqualError(t, streamWriter.SetColWidth(0, 3, 20), ErrColumnNumber.Error())
	assert.EqualError(t, streamWriter.SetColWidth(TotalColumns+1, 3, 20), ErrColumnNumber.Error())
	assert.EqualError(t, streamWriter.SetColWidth(1, 3, MaxColumnWidth+1), ErrColumnWidth.Error())
	assert.NoError(t, streamWriter.SetRow("A1", []interface{}{"A", "B", "C"}))
	assert.EqualError(t, streamWriter.SetColWidth(2, 3, 20), ErrStreamSetColWidth.Error())
}

func TestStreamTable(t *testing.T) {
	file := NewFile()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)

	// Write some rows. We want enough rows to force a temp file (>16MB).
	assert.NoError(t, streamWriter.SetRow("A1", []interface{}{"A", "B", "C"}))
	row := []interface{}{1, 2, 3}
	for r := 2; r < 10000; r++ {
		assert.NoError(t, streamWriter.SetRow(fmt.Sprintf("A%d", r), row))
	}

	// Write a table.
	assert.NoError(t, streamWriter.AddTable("A1", "C2", ``))
	assert.NoError(t, streamWriter.Flush())

	// Verify the table has names.
	var table xlsxTable
	val, ok := file.Pkg.Load("xl/tables/table1.xml")
	assert.True(t, ok)
	assert.NoError(t, xml.Unmarshal(val.([]byte), &table))
	assert.Equal(t, "A", table.TableColumns.TableColumn[0].Name)
	assert.Equal(t, "B", table.TableColumns.TableColumn[1].Name)
	assert.Equal(t, "C", table.TableColumns.TableColumn[2].Name)

	assert.NoError(t, streamWriter.AddTable("A1", "C1", ``))

	// Test add table with illegal formatset.
	assert.EqualError(t, streamWriter.AddTable("B26", "A21", `{x}`), "invalid character 'x' looking for beginning of object key string")
	// Test add table with illegal cell coordinates.
	assert.EqualError(t, streamWriter.AddTable("A", "B1", `{}`), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	assert.EqualError(t, streamWriter.AddTable("A1", "B", `{}`), newCellNameToCoordinatesError("B", newInvalidCellNameError("B")).Error())
}

func TestStreamMergeCells(t *testing.T) {
	file := NewFile()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.MergeCell("A1", "D1"))
	// Test merge cells with illegal cell coordinates.
	assert.EqualError(t, streamWriter.MergeCell("A", "D1"), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	assert.NoError(t, streamWriter.Flush())
	// Save spreadsheet by the given path.
	assert.NoError(t, file.SaveAs(filepath.Join("test", "TestStreamMergeCells.xlsx")))
}

func TestNewStreamWriter(t *testing.T) {
	// Test error exceptions
	file := NewFile()
	_, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	_, err = file.NewStreamWriter("SheetN")
	assert.EqualError(t, err, "sheet SheetN is not exist")
}

func TestSetRow(t *testing.T) {
	// Test error exceptions
	file := NewFile()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.EqualError(t, streamWriter.SetRow("A", []interface{}{}), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
}

func TestSetCellValFunc(t *testing.T) {
	f := NewFile()
	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	c := &xlsxC{}
	assert.NoError(t, sw.setCellValFunc(c, 128))
	assert.NoError(t, sw.setCellValFunc(c, int8(-128)))
	assert.NoError(t, sw.setCellValFunc(c, int16(-32768)))
	assert.NoError(t, sw.setCellValFunc(c, int32(-2147483648)))
	assert.NoError(t, sw.setCellValFunc(c, int64(-9223372036854775808)))
	assert.NoError(t, sw.setCellValFunc(c, uint(128)))
	assert.NoError(t, sw.setCellValFunc(c, uint8(255)))
	assert.NoError(t, sw.setCellValFunc(c, uint16(65535)))
	assert.NoError(t, sw.setCellValFunc(c, uint32(4294967295)))
	assert.NoError(t, sw.setCellValFunc(c, uint64(18446744073709551615)))
	assert.NoError(t, sw.setCellValFunc(c, float32(100.1588)))
	assert.NoError(t, sw.setCellValFunc(c, float64(100.1588)))
	assert.NoError(t, sw.setCellValFunc(c, " Hello"))
	assert.NoError(t, sw.setCellValFunc(c, []byte(" Hello")))
	assert.NoError(t, sw.setCellValFunc(c, time.Now().UTC()))
	assert.NoError(t, sw.setCellValFunc(c, time.Duration(1e13)))
	assert.NoError(t, sw.setCellValFunc(c, true))
	assert.NoError(t, sw.setCellValFunc(c, nil))
	assert.NoError(t, sw.setCellValFunc(c, complex64(5+10i)))
}
