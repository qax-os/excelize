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
	row[0] = strings.Repeat("c", 32769)
	assert.NoError(t, streamWriter.SetRow("A1", row))

	// Test leading and ending space(s) character characters in a cell.
	row = make([]interface{}, 1)
	row[0] = " characters"
	assert.NoError(t, streamWriter.SetRow("A2", row))

	row = make([]interface{}, 1)
	row[0] = []byte("Word")
	assert.NoError(t, streamWriter.SetRow("A3", row))

	// Test set cell with style.
	styleID, err := file.NewStyle(`{"font":{"color":"#777777"}}`)
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.SetRow("A4", []interface{}{Cell{StyleID: styleID}, Cell{Formula: "SUM(A10,B10)"}}))
	assert.NoError(t, streamWriter.SetRow("A5", []interface{}{&Cell{StyleID: styleID, Value: "cell"}, &Cell{Formula: "SUM(A10,B10)"}}))
	assert.EqualError(t, streamWriter.SetRow("A6", []interface{}{time.Now()}), "only UTC time expected")

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

	// Test unsupport charset
	file = NewFile()
	delete(file.Sheet, "xl/worksheets/sheet1.xml")
	file.XLSX["xl/worksheets/sheet1.xml"] = MacintoshCyrillicCharset
	_, err = file.NewStreamWriter("Sheet1")
	assert.EqualError(t, err, "xml decode error: XML syntax error on line 1: invalid UTF-8")
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
	assert.NoError(t, xml.Unmarshal(file.XLSX["xl/tables/table1.xml"], &table))
	assert.Equal(t, "A", table.TableColumns.TableColumn[0].Name)
	assert.Equal(t, "B", table.TableColumns.TableColumn[1].Name)
	assert.Equal(t, "C", table.TableColumns.TableColumn[2].Name)

	assert.NoError(t, streamWriter.AddTable("A1", "C1", ``))

	// Test add table with illegal formatset.
	assert.EqualError(t, streamWriter.AddTable("B26", "A21", `{x}`), "invalid character 'x' looking for beginning of object key string")
	// Test add table with illegal cell coordinates.
	assert.EqualError(t, streamWriter.AddTable("A", "B1", `{}`), `cannot convert cell "A" to coordinates: invalid cell name "A"`)
	assert.EqualError(t, streamWriter.AddTable("A1", "B", `{}`), `cannot convert cell "B" to coordinates: invalid cell name "B"`)
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
	assert.EqualError(t, streamWriter.SetRow("A", []interface{}{}), `cannot convert cell "A" to coordinates: invalid cell name "A"`)
}

func TestSetCellValFunc(t *testing.T) {
	c := &xlsxC{}
	assert.NoError(t, setCellValFunc(c, 128))
	assert.NoError(t, setCellValFunc(c, int8(-128)))
	assert.NoError(t, setCellValFunc(c, int16(-32768)))
	assert.NoError(t, setCellValFunc(c, int32(-2147483648)))
	assert.NoError(t, setCellValFunc(c, int64(-9223372036854775808)))
	assert.NoError(t, setCellValFunc(c, uint(128)))
	assert.NoError(t, setCellValFunc(c, uint8(255)))
	assert.NoError(t, setCellValFunc(c, uint16(65535)))
	assert.NoError(t, setCellValFunc(c, uint32(4294967295)))
	assert.NoError(t, setCellValFunc(c, uint64(18446744073709551615)))
	assert.NoError(t, setCellValFunc(c, float32(100.1588)))
	assert.NoError(t, setCellValFunc(c, float64(100.1588)))
	assert.NoError(t, setCellValFunc(c, " Hello"))
	assert.NoError(t, setCellValFunc(c, []byte(" Hello")))
	assert.NoError(t, setCellValFunc(c, time.Now().UTC()))
	assert.NoError(t, setCellValFunc(c, time.Duration(1e13)))
	assert.NoError(t, setCellValFunc(c, true))
	assert.NoError(t, setCellValFunc(c, nil))
	assert.NoError(t, setCellValFunc(c, complex64(5+10i)))
}
