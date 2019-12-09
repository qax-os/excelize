package excelize

import (
	"math/rand"
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestStreamWriter(t *testing.T) {
	file := NewFile()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)

	// Test max characters in a cell.
	row := make([]interface{}, 1)
	row[0] = strings.Repeat("c", 32769)
	assert.NoError(t, streamWriter.SetRow("A1", &row))

	// Test leading and ending space(s) character characters in a cell.
	row = make([]interface{}, 1)
	row[0] = " characters"
	assert.NoError(t, streamWriter.SetRow("A2", &row))

	row = make([]interface{}, 1)
	row[0] = []byte("Word")
	assert.NoError(t, streamWriter.SetRow("A3", &row))

	for rowID := 10; rowID <= 51200; rowID++ {
		row := make([]interface{}, 50)
		for colID := 0; colID < 50; colID++ {
			row[colID] = rand.Intn(640000)
		}
		cell, _ := CoordinatesToCellName(1, rowID)
		assert.NoError(t, streamWriter.SetRow(cell, &row))
	}

	err = streamWriter.Flush()
	assert.NoError(t, err)
	// Save xlsx file by the given path.
	assert.NoError(t, file.SaveAs(filepath.Join("test", "TestStreamWriter.xlsx")))

	// Test error exceptions
	streamWriter, err = file.NewStreamWriter("SheetN")
	assert.EqualError(t, err, "sheet SheetN is not exist")
}

func TestFlush(t *testing.T) {
	// Test error exceptions
	file := NewFile()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	streamWriter.Sheet = "SheetN"
	assert.EqualError(t, streamWriter.Flush(), "sheet SheetN is not exist")
}

func TestSetRow(t *testing.T) {
	// Test error exceptions
	file := NewFile()
	streamWriter, err := file.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	assert.EqualError(t, streamWriter.SetRow("A", &[]interface{}{}), `cannot convert cell "A" to coordinates: invalid cell name "A"`)
	assert.EqualError(t, streamWriter.SetRow("A1", []interface{}{}), `pointer to slice expected`)
}
