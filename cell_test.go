package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestCheckCellInArea(t *testing.T) {
	expectedTrueCellInAreaList := [][2]string{
		{"c2", "A1:AAZ32"},
		{"B9", "A1:B9"},
		{"C2", "C2:C2"},
	}

	for _, expectedTrueCellInArea := range expectedTrueCellInAreaList {
		cell := expectedTrueCellInArea[0]
		area := expectedTrueCellInArea[1]

		assert.Truef(t, checkCellInArea(cell, area),
			"Expected cell %v to be in area %v, got false\n", cell, area)
	}

	expectedFalseCellInAreaList := [][2]string{
		{"c2", "A4:AAZ32"},
		{"C4", "D6:A1"}, // weird case, but you never know
		{"AEF42", "BZ40:AEF41"},
	}

	for _, expectedFalseCellInArea := range expectedFalseCellInAreaList {
		cell := expectedFalseCellInArea[0]
		area := expectedFalseCellInArea[1]

		assert.Falsef(t, checkCellInArea(cell, area),
			"Expected cell %v not to be inside of area %v, but got true\n", cell, area)
	}

	assert.Panics(t, func() {
		checkCellInArea("AA0", "Z0:AB1")
	})
}

func TestSetCellFloat(t *testing.T) {
	sheet := "Sheet1"
	t.Run("with no decimal", func(t *testing.T) {
		f := NewFile()
		f.SetCellFloat(sheet, "A1", 123.0, -1, 64)
		f.SetCellFloat(sheet, "A2", 123.0, 1, 64)
		assert.Equal(t, "123", f.GetCellValue(sheet, "A1"), "A1 should be 123")
		assert.Equal(t, "123.0", f.GetCellValue(sheet, "A2"), "A2 should be 123.0")
	})

	t.Run("with a decimal and precision limit", func(t *testing.T) {
		f := NewFile()
		f.SetCellFloat(sheet, "A1", 123.42, 1, 64)
		assert.Equal(t, "123.4", f.GetCellValue(sheet, "A1"), "A1 should be 123.4")
	})

	t.Run("with a decimal and no limit", func(t *testing.T) {
		f := NewFile()
		f.SetCellFloat(sheet, "A1", 123.42, -1, 64)
		assert.Equal(t, "123.42", f.GetCellValue(sheet, "A1"), "A1 should be 123.42")
	})
}

func ExampleFile_SetCellFloat() {
	f := NewFile()
	var x float64 = 3.14159265
	f.SetCellFloat("Sheet1", "A1", x, 2, 64)
	fmt.Println(f.GetCellValue("Sheet1", "A1"))
	// Output: 3.14
}
