package excelize

import (
	"encoding/xml"
	"fmt"
	"math/rand"
	"path/filepath"
	"strings"
	"testing"

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
			streamWriter.SetRow(cell, row, nil)
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
	assert.NoError(t, streamWriter.SetRow("A1", row, nil))

	// Test leading and ending space(s) character characters in a cell.
	row = make([]interface{}, 1)
	row[0] = " characters"
	assert.NoError(t, streamWriter.SetRow("A2", row, nil))

	row = make([]interface{}, 1)
	row[0] = []byte("Word")
	assert.NoError(t, streamWriter.SetRow("A3", row, nil))

	for rowID := 10; rowID <= 51200; rowID++ {
		row := make([]interface{}, 50)
		for colID := 0; colID < 50; colID++ {
			row[colID] = rand.Intn(640000)
		}
		cell, _ := CoordinatesToCellName(1, rowID)
		assert.NoError(t, streamWriter.SetRow(cell, row, nil))
	}

	assert.NoError(t, streamWriter.Flush())
	// Save xlsx file by the given path.
	assert.NoError(t, file.SaveAs(filepath.Join("test", "TestStreamWriter.xlsx")))

	// Test close temporary file error
	file = NewFile()
	streamWriter, err = file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	for rowID := 10; rowID <= 51200; rowID++ {
		row := make([]interface{}, 50)
		for colID := 0; colID < 50; colID++ {
			row[colID] = rand.Intn(640000)
		}
		cell, _ := CoordinatesToCellName(1, rowID)
		assert.NoError(t, streamWriter.SetRow(cell, row, nil))
	}
	assert.NoError(t, streamWriter.rawData.Close())
	assert.Error(t, streamWriter.Flush())
}

func TestStreamTable(t *testing.T) {
	file := NewFile()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)

	// Write some rows. We want enough rows to force a temp file (>16MB).
	assert.NoError(t, streamWriter.SetRow("A1", []interface{}{"A", "B", "C"}, nil))
	row := []interface{}{1, 2, 3}
	for r := 2; r < 10000; r++ {
		assert.NoError(t, streamWriter.SetRow(fmt.Sprintf("A%d", r), row, nil))
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
	assert.EqualError(t, streamWriter.SetRow("A", []interface{}{}, nil), `cannot convert cell "A" to coordinates: invalid cell name "A"`)
}
