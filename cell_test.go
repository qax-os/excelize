package excelize

import (
	"fmt"
	"path/filepath"
	"testing"
	"time"

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
		ok, err := checkCellInArea(cell, area)
		assert.NoError(t, err)
		assert.Truef(t, ok,
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
		ok, err := checkCellInArea(cell, area)
		assert.NoError(t, err)
		assert.Falsef(t, ok,
			"Expected cell %v not to be inside of area %v, but got true\n", cell, area)
	}

	ok, err := checkCellInArea("AA0", "Z0:AB1")
	assert.EqualError(t, err, `cannot convert cell "AA0" to coordinates: invalid cell name "AA0"`)
	assert.False(t, ok)
}

func TestSetCellFloat(t *testing.T) {
	sheet := "Sheet1"
	t.Run("with no decimal", func(t *testing.T) {
		f := NewFile()
		f.SetCellFloat(sheet, "A1", 123.0, -1, 64)
		f.SetCellFloat(sheet, "A2", 123.0, 1, 64)
		val, err := f.GetCellValue(sheet, "A1")
		assert.NoError(t, err)
		assert.Equal(t, "123", val, "A1 should be 123")
		val, err = f.GetCellValue(sheet, "A2")
		assert.NoError(t, err)
		assert.Equal(t, "123.0", val, "A2 should be 123.0")
	})

	t.Run("with a decimal and precision limit", func(t *testing.T) {
		f := NewFile()
		f.SetCellFloat(sheet, "A1", 123.42, 1, 64)
		val, err := f.GetCellValue(sheet, "A1")
		assert.NoError(t, err)
		assert.Equal(t, "123.4", val, "A1 should be 123.4")
	})

	t.Run("with a decimal and no limit", func(t *testing.T) {
		f := NewFile()
		f.SetCellFloat(sheet, "A1", 123.42, -1, 64)
		val, err := f.GetCellValue(sheet, "A1")
		assert.NoError(t, err)
		assert.Equal(t, "123.42", val, "A1 should be 123.42")
	})
	f := NewFile()
	assert.EqualError(t, f.SetCellFloat(sheet, "A", 123.42, -1, 64), `cannot convert cell "A" to coordinates: invalid cell name "A"`)
}

func TestSetCellValue(t *testing.T) {
	f := NewFile()
	assert.EqualError(t, f.SetCellValue("Sheet1", "A", time.Now().UTC()), `cannot convert cell "A" to coordinates: invalid cell name "A"`)
	assert.EqualError(t, f.SetCellValue("Sheet1", "A", time.Duration(1e13)), `cannot convert cell "A" to coordinates: invalid cell name "A"`)
}

func TestSetCellBool(t *testing.T) {
	f := NewFile()
	assert.EqualError(t, f.SetCellBool("Sheet1", "A", true), `cannot convert cell "A" to coordinates: invalid cell name "A"`)
}

func TestGetCellFormula(t *testing.T) {
	f := NewFile()
	f.GetCellFormula("Sheet", "A1")
}

func TestMergeCell(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.EqualError(t, f.MergeCell("Sheet1", "A", "B"), `cannot convert cell "A" to coordinates: invalid cell name "A"`)
	f.MergeCell("Sheet1", "D9", "D9")
	f.MergeCell("Sheet1", "D9", "E9")
	f.MergeCell("Sheet1", "H14", "G13")
	f.MergeCell("Sheet1", "C9", "D8")
	f.MergeCell("Sheet1", "F11", "G13")
	f.MergeCell("Sheet1", "H7", "B15")
	f.MergeCell("Sheet1", "D11", "F13")
	f.MergeCell("Sheet1", "G10", "K12")
	f.SetCellValue("Sheet1", "G11", "set value in merged cell")
	f.SetCellInt("Sheet1", "H11", 100)
	f.SetCellValue("Sheet1", "I11", float64(0.5))
	f.SetCellHyperLink("Sheet1", "J11", "https://github.com/360EntSecGroup-Skylar/excelize", "External")
	f.SetCellFormula("Sheet1", "G12", "SUM(Sheet1!B19,Sheet1!C19)")
	f.GetCellValue("Sheet1", "H11")
	f.GetCellValue("Sheet2", "A6") // Merged cell ref is single coordinate.
	f.GetCellFormula("Sheet1", "G12")

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestMergeCell.xlsx")))
}

func ExampleFile_SetCellFloat() {
	f := NewFile()
	var x = 3.14159265
	f.SetCellFloat("Sheet1", "A1", x, 2, 64)
	val, _ := f.GetCellValue("Sheet1", "A1")
	fmt.Println(val)
	// Output: 3.14
}

func BenchmarkSetCellValue(b *testing.B) {
	values := []string{"First", "Second", "Third", "Fourth", "Fifth", "Sixth"}
	cols := []string{"A", "B", "C", "D", "E", "F"}
	f := NewFile()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		for j := 0; j < len(values); j++ {
			f.SetCellValue("Sheet1", fmt.Sprint(cols[j], i), values[j])
		}
	}
}

func TestOverflowNumericCell(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "OverflowNumericCell.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	val, err := f.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	// GOARCH=amd64 - all ok; GOARCH=386 - actual: "-2147483648"
	assert.Equal(t, "8595602512225", val, "A1 should be 8595602512225")
}
