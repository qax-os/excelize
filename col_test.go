package excelize

import (
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestColumnVisibility(t *testing.T) {
	t.Run("TestBook1", func(t *testing.T) {
		f, err := prepareTestBook1()
		if !assert.NoError(t, err) {
			t.FailNow()
		}

		assert.NoError(t, f.SetColVisible("Sheet1", "F", false))
		assert.NoError(t, f.SetColVisible("Sheet1", "F", true))
		visible, err := f.GetColVisible("Sheet1", "F")
		assert.Equal(t, true, visible)
		assert.NoError(t, err)

		// Test get column visiable on not exists worksheet.
		_, err = f.GetColVisible("SheetN", "F")
		assert.EqualError(t, err, "sheet SheetN is not exist")

		// Test get column visiable with illegal cell coordinates.
		_, err = f.GetColVisible("Sheet1", "*")
		assert.EqualError(t, err, `invalid column name "*"`)
		assert.EqualError(t, f.SetColVisible("Sheet1", "*", false), `invalid column name "*"`)

		f.NewSheet("Sheet3")
		assert.NoError(t, f.SetColVisible("Sheet3", "E", false))

		assert.EqualError(t, f.SetColVisible("SheetN", "E", false), "sheet SheetN is not exist")
		assert.NoError(t, f.SaveAs(filepath.Join("test", "TestColumnVisibility.xlsx")))
	})

	t.Run("TestBook3", func(t *testing.T) {
		f, err := prepareTestBook3()
		if !assert.NoError(t, err) {
			t.FailNow()
		}
		f.GetColVisible("Sheet1", "B")
	})
}

func TestOutlineLevel(t *testing.T) {
	f := NewFile()
	f.GetColOutlineLevel("Sheet1", "D")
	f.NewSheet("Sheet2")
	f.SetColOutlineLevel("Sheet1", "D", 4)
	f.GetColOutlineLevel("Sheet1", "D")
	f.GetColOutlineLevel("Shee2", "A")
	f.SetColWidth("Sheet2", "A", "D", 13)
	f.SetColOutlineLevel("Sheet2", "B", 2)
	f.SetRowOutlineLevel("Sheet1", 2, 250)

	// Test set and get column outline level with illegal cell coordinates.
	assert.EqualError(t, f.SetColOutlineLevel("Sheet1", "*", 1), `invalid column name "*"`)
	_, err := f.GetColOutlineLevel("Sheet1", "*")
	assert.EqualError(t, err, `invalid column name "*"`)

	// Test set column outline level on not exists worksheet.
	assert.EqualError(t, f.SetColOutlineLevel("SheetN", "E", 2), "sheet SheetN is not exist")

	assert.EqualError(t, f.SetRowOutlineLevel("Sheet1", 0, 1), "invalid row number 0")
	level, err := f.GetRowOutlineLevel("Sheet1", 2)
	assert.NoError(t, err)
	assert.Equal(t, uint8(250), level)

	_, err = f.GetRowOutlineLevel("Sheet1", 0)
	assert.EqualError(t, err, `invalid row number 0`)

	level, err = f.GetRowOutlineLevel("Sheet1", 10)
	assert.NoError(t, err)
	assert.Equal(t, uint8(0), level)

	err = f.SaveAs(filepath.Join("test", "TestOutlineLevel.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	f, err = OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	f.SetColOutlineLevel("Sheet2", "B", 2)
}

func TestSetColStyle(t *testing.T) {
	f := NewFile()
	style, err := f.NewStyle(`{"fill":{"type":"pattern","color":["#94d3a2"],"pattern":1}}`)
	assert.NoError(t, err)
	// Test set column style on not exists worksheet.
	assert.EqualError(t, f.SetColStyle("SheetN", "E", style), "sheet SheetN is not exist")
	// Test set column style with illegal cell coordinates.
	assert.EqualError(t, f.SetColStyle("Sheet1", "*", style), `invalid column name "*"`)
	assert.EqualError(t, f.SetColStyle("Sheet1", "A:*", style), `invalid column name "*"`)

	assert.NoError(t, f.SetColStyle("Sheet1", "B", style))
	// Test set column style with already exists column with style.
	assert.NoError(t, f.SetColStyle("Sheet1", "B", style))
	assert.NoError(t, f.SetColStyle("Sheet1", "D:C", style))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetColStyle.xlsx")))
}

func TestColWidth(t *testing.T) {
	f := NewFile()
	f.SetColWidth("Sheet1", "B", "A", 12)
	f.SetColWidth("Sheet1", "A", "B", 12)
	f.GetColWidth("Sheet1", "A")
	f.GetColWidth("Sheet1", "C")

	// Test set and get column width with illegal cell coordinates.
	_, err := f.GetColWidth("Sheet1", "*")
	assert.EqualError(t, err, `invalid column name "*"`)
	assert.EqualError(t, f.SetColWidth("Sheet1", "*", "B", 1), `invalid column name "*"`)
	assert.EqualError(t, f.SetColWidth("Sheet1", "A", "*", 1), `invalid column name "*"`)

	// Test set column width on not exists worksheet.
	assert.EqualError(t, f.SetColWidth("SheetN", "B", "A", 12), "sheet SheetN is not exist")

	// Test get column width on not exists worksheet.
	_, err = f.GetColWidth("SheetN", "A")
	assert.EqualError(t, err, "sheet SheetN is not exist")

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestColWidth.xlsx")))
	convertRowHeightToPixels(0)
}

func TestInsertCol(t *testing.T) {
	f := NewFile()
	sheet1 := f.GetSheetName(1)

	fillCells(f, sheet1, 10, 10)

	f.SetCellHyperLink(sheet1, "A5", "https://github.com/360EntSecGroup-Skylar/excelize", "External")
	f.MergeCell(sheet1, "A1", "C3")

	err := f.AutoFilter(sheet1, "A2", "B2", `{"column":"B","expression":"x != blanks"}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.InsertCol(sheet1, "A"))

	// Test insert column with illegal cell coordinates.
	assert.EqualError(t, f.InsertCol("Sheet1", "*"), `invalid column name "*"`)

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestInsertCol.xlsx")))
}

func TestRemoveCol(t *testing.T) {
	f := NewFile()
	sheet1 := f.GetSheetName(1)

	fillCells(f, sheet1, 10, 15)

	f.SetCellHyperLink(sheet1, "A5", "https://github.com/360EntSecGroup-Skylar/excelize", "External")
	f.SetCellHyperLink(sheet1, "C5", "https://github.com", "External")

	f.MergeCell(sheet1, "A1", "B1")
	f.MergeCell(sheet1, "A2", "B2")

	assert.NoError(t, f.RemoveCol(sheet1, "A"))
	assert.NoError(t, f.RemoveCol(sheet1, "A"))

	// Test remove column with illegal cell coordinates.
	assert.EqualError(t, f.RemoveCol("Sheet1", "*"), `invalid column name "*"`)

	// Test remove column on not exists worksheet.
	assert.EqualError(t, f.RemoveCol("SheetN", "B"), "sheet SheetN is not exist")

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestRemoveCol.xlsx")))
}
