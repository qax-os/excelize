package excelize

import (
	"fmt"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestRows(t *testing.T) {
	const sheet2 = "Sheet2"

	xlsx, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	rows, err := xlsx.Rows(sheet2)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	collectedRows := make([][]string, 0)
	for rows.Next() {
		columns, err := rows.Columns()
		assert.NoError(t, err)
		collectedRows = append(collectedRows, trimSliceSpace(columns))
	}
	if !assert.NoError(t, rows.Error()) {
		t.FailNow()
	}

	returnedRows, err := xlsx.GetRows(sheet2)
	assert.NoError(t, err)
	for i := range returnedRows {
		returnedRows[i] = trimSliceSpace(returnedRows[i])
	}
	if !assert.Equal(t, collectedRows, returnedRows) {
		t.FailNow()
	}

	r := Rows{}
	r.Columns()
}

func TestRowsError(t *testing.T) {
	xlsx, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	_, err = xlsx.Rows("SheetN")
	assert.EqualError(t, err, "sheet SheetN is not exist")
}

func TestRowHeight(t *testing.T) {
	xlsx := NewFile()
	sheet1 := xlsx.GetSheetName(1)

	assert.EqualError(t, xlsx.SetRowHeight(sheet1, 0, defaultRowHeightPixels+1.0), "invalid row number 0")

	_, err := xlsx.GetRowHeight("Sheet1", 0)
	assert.EqualError(t, err, "invalid row number 0")

	assert.NoError(t, xlsx.SetRowHeight(sheet1, 1, 111.0))
	height, err := xlsx.GetRowHeight(sheet1, 1)
	assert.NoError(t, err)
	assert.Equal(t, 111.0, height)

	assert.NoError(t, xlsx.SetRowHeight(sheet1, 4, 444.0))
	height, err = xlsx.GetRowHeight(sheet1, 4)
	assert.NoError(t, err)
	assert.Equal(t, 444.0, height)

	// Test get row height that rows index over exists rows.
	height, err = xlsx.GetRowHeight(sheet1, 5)
	assert.NoError(t, err)
	assert.Equal(t, defaultRowHeightPixels, height)

	// Test get row height that rows heights haven't changed.
	height, err = xlsx.GetRowHeight(sheet1, 3)
	assert.NoError(t, err)
	assert.Equal(t, defaultRowHeightPixels, height)

	// Test set and get row height on not exists worksheet.
	assert.EqualError(t, xlsx.SetRowHeight("SheetN", 1, 111.0), "sheet SheetN is not exist")
	_, err = xlsx.GetRowHeight("SheetN", 3)
	assert.EqualError(t, err, "sheet SheetN is not exist")

	err = xlsx.SaveAs(filepath.Join("test", "TestRowHeight.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	convertColWidthToPixels(0)
}

func TestRowVisibility(t *testing.T) {
	xlsx, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	xlsx.NewSheet("Sheet3")
	assert.NoError(t, xlsx.SetRowVisible("Sheet3", 2, false))
	assert.NoError(t, xlsx.SetRowVisible("Sheet3", 2, true))
	xlsx.GetRowVisible("Sheet3", 2)
	xlsx.GetRowVisible("Sheet3", 25)
	assert.EqualError(t, xlsx.SetRowVisible("Sheet3", 0, true), "invalid row number 0")

	visible, err := xlsx.GetRowVisible("Sheet3", 0)
	assert.Equal(t, false, visible)
	assert.EqualError(t, err, "invalid row number 0")

	assert.NoError(t, xlsx.SaveAs(filepath.Join("test", "TestRowVisibility.xlsx")))
}

func TestRemoveRow(t *testing.T) {
	xlsx := NewFile()
	sheet1 := xlsx.GetSheetName(1)
	r, err := xlsx.workSheetReader(sheet1)
	assert.NoError(t, err)
	const (
		colCount = 10
		rowCount = 10
	)
	fillCells(xlsx, sheet1, colCount, rowCount)

	xlsx.SetCellHyperLink(sheet1, "A5", "https://github.com/360EntSecGroup-Skylar/excelize", "External")

	assert.EqualError(t, xlsx.RemoveRow(sheet1, -1), "invalid row number -1")

	assert.EqualError(t, xlsx.RemoveRow(sheet1, 0), "invalid row number 0")

	assert.NoError(t, xlsx.RemoveRow(sheet1, 4))
	if !assert.Len(t, r.SheetData.Row, rowCount-1) {
		t.FailNow()
	}

	xlsx.MergeCell(sheet1, "B3", "B5")

	assert.NoError(t, xlsx.RemoveRow(sheet1, 2))
	if !assert.Len(t, r.SheetData.Row, rowCount-2) {
		t.FailNow()
	}

	assert.NoError(t, xlsx.RemoveRow(sheet1, 4))
	if !assert.Len(t, r.SheetData.Row, rowCount-3) {
		t.FailNow()
	}

	err = xlsx.AutoFilter(sheet1, "A2", "A2", `{"column":"A","expression":"x != blanks"}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, xlsx.RemoveRow(sheet1, 1))
	if !assert.Len(t, r.SheetData.Row, rowCount-4) {
		t.FailNow()
	}

	assert.NoError(t, xlsx.RemoveRow(sheet1, 2))
	if !assert.Len(t, r.SheetData.Row, rowCount-5) {
		t.FailNow()
	}

	assert.NoError(t, xlsx.RemoveRow(sheet1, 1))
	if !assert.Len(t, r.SheetData.Row, rowCount-6) {
		t.FailNow()
	}

	assert.NoError(t, xlsx.RemoveRow(sheet1, 10))
	assert.NoError(t, xlsx.SaveAs(filepath.Join("test", "TestRemoveRow.xlsx")))
}

func TestInsertRow(t *testing.T) {
	xlsx := NewFile()
	sheet1 := xlsx.GetSheetName(1)
	r, err := xlsx.workSheetReader(sheet1)
	assert.NoError(t, err)
	const (
		colCount = 10
		rowCount = 10
	)
	fillCells(xlsx, sheet1, colCount, rowCount)

	xlsx.SetCellHyperLink(sheet1, "A5", "https://github.com/360EntSecGroup-Skylar/excelize", "External")

	assert.EqualError(t, xlsx.InsertRow(sheet1, -1), "invalid row number -1")

	assert.EqualError(t, xlsx.InsertRow(sheet1, 0), "invalid row number 0")

	assert.NoError(t, xlsx.InsertRow(sheet1, 1))
	if !assert.Len(t, r.SheetData.Row, rowCount+1) {
		t.FailNow()
	}

	assert.NoError(t, xlsx.InsertRow(sheet1, 4))
	if !assert.Len(t, r.SheetData.Row, rowCount+2) {
		t.FailNow()
	}

	assert.NoError(t, xlsx.SaveAs(filepath.Join("test", "TestInsertRow.xlsx")))
}

// Testing internal sructure state after insert operations.
// It is important for insert workflow to be constant to avoid side effect with functions related to internal structure.
func TestInsertRowInEmptyFile(t *testing.T) {
	xlsx := NewFile()
	sheet1 := xlsx.GetSheetName(1)
	r, err := xlsx.workSheetReader(sheet1)
	assert.NoError(t, err)
	assert.NoError(t, xlsx.InsertRow(sheet1, 1))
	assert.Len(t, r.SheetData.Row, 0)
	assert.NoError(t, xlsx.InsertRow(sheet1, 2))
	assert.Len(t, r.SheetData.Row, 0)
	assert.NoError(t, xlsx.InsertRow(sheet1, 99))
	assert.Len(t, r.SheetData.Row, 0)
	assert.NoError(t, xlsx.SaveAs(filepath.Join("test", "TestInsertRowInEmptyFile.xlsx")))
}

func TestDuplicateRowFromSingleRow(t *testing.T) {
	const sheet = "Sheet1"
	outFile := filepath.Join("test", "TestDuplicateRow.%s.xlsx")

	cells := map[string]string{
		"A1": "A1 Value",
		"A2": "A2 Value",
		"A3": "A3 Value",
		"B1": "B1 Value",
		"B2": "B2 Value",
		"B3": "B3 Value",
	}

	t.Run("FromSingleRow", func(t *testing.T) {
		xlsx := NewFile()
		xlsx.SetCellStr(sheet, "A1", cells["A1"])
		xlsx.SetCellStr(sheet, "B1", cells["B1"])

		assert.NoError(t, xlsx.DuplicateRow(sheet, 1))
		if !assert.NoError(t, xlsx.SaveAs(fmt.Sprintf(outFile, "TestDuplicateRow.FromSingleRow_1"))) {
			t.FailNow()
		}
		expect := map[string]string{
			"A1": cells["A1"], "B1": cells["B1"],
			"A2": cells["A1"], "B2": cells["B1"],
		}
		for cell, val := range expect {
			v, err := xlsx.GetCellValue(sheet, cell)
			assert.NoError(t, err)
			if !assert.Equal(t, val, v, cell) {
				t.FailNow()
			}
		}

		assert.NoError(t, xlsx.DuplicateRow(sheet, 2))
		if !assert.NoError(t, xlsx.SaveAs(fmt.Sprintf(outFile, "TestDuplicateRow.FromSingleRow_2"))) {
			t.FailNow()
		}
		expect = map[string]string{
			"A1": cells["A1"], "B1": cells["B1"],
			"A2": cells["A1"], "B2": cells["B1"],
			"A3": cells["A1"], "B3": cells["B1"],
		}
		for cell, val := range expect {
			v, err := xlsx.GetCellValue(sheet, cell)
			assert.NoError(t, err)
			if !assert.Equal(t, val, v, cell) {
				t.FailNow()
			}
		}
	})
}

func TestDuplicateRowUpdateDuplicatedRows(t *testing.T) {
	const sheet = "Sheet1"
	outFile := filepath.Join("test", "TestDuplicateRow.%s.xlsx")

	cells := map[string]string{
		"A1": "A1 Value",
		"A2": "A2 Value",
		"A3": "A3 Value",
		"B1": "B1 Value",
		"B2": "B2 Value",
		"B3": "B3 Value",
	}

	t.Run("UpdateDuplicatedRows", func(t *testing.T) {
		xlsx := NewFile()
		xlsx.SetCellStr(sheet, "A1", cells["A1"])
		xlsx.SetCellStr(sheet, "B1", cells["B1"])

		assert.NoError(t, xlsx.DuplicateRow(sheet, 1))

		xlsx.SetCellStr(sheet, "A2", cells["A2"])
		xlsx.SetCellStr(sheet, "B2", cells["B2"])

		if !assert.NoError(t, xlsx.SaveAs(fmt.Sprintf(outFile, "TestDuplicateRow.UpdateDuplicatedRows"))) {
			t.FailNow()
		}
		expect := map[string]string{
			"A1": cells["A1"], "B1": cells["B1"],
			"A2": cells["A2"], "B2": cells["B2"],
		}
		for cell, val := range expect {
			v, err := xlsx.GetCellValue(sheet, cell)
			assert.NoError(t, err)
			if !assert.Equal(t, val, v, cell) {
				t.FailNow()
			}
		}
	})
}

func TestDuplicateRowFirstOfMultipleRows(t *testing.T) {
	const sheet = "Sheet1"
	outFile := filepath.Join("test", "TestDuplicateRow.%s.xlsx")

	cells := map[string]string{
		"A1": "A1 Value",
		"A2": "A2 Value",
		"A3": "A3 Value",
		"B1": "B1 Value",
		"B2": "B2 Value",
		"B3": "B3 Value",
	}

	newFileWithDefaults := func() *File {
		f := NewFile()
		for cell, val := range cells {
			f.SetCellStr(sheet, cell, val)

		}
		return f
	}

	t.Run("FirstOfMultipleRows", func(t *testing.T) {
		xlsx := newFileWithDefaults()

		assert.NoError(t, xlsx.DuplicateRow(sheet, 1))

		if !assert.NoError(t, xlsx.SaveAs(fmt.Sprintf(outFile, "TestDuplicateRow.FirstOfMultipleRows"))) {
			t.FailNow()
		}
		expect := map[string]string{
			"A1": cells["A1"], "B1": cells["B1"],
			"A2": cells["A1"], "B2": cells["B1"],
			"A3": cells["A2"], "B3": cells["B2"],
			"A4": cells["A3"], "B4": cells["B3"],
		}
		for cell, val := range expect {
			v, err := xlsx.GetCellValue(sheet, cell)
			assert.NoError(t, err)
			if !assert.Equal(t, val, v, cell) {
				t.FailNow()
			}
		}
	})
}

func TestDuplicateRowZeroWithNoRows(t *testing.T) {
	const sheet = "Sheet1"
	outFile := filepath.Join("test", "TestDuplicateRow.%s.xlsx")

	t.Run("ZeroWithNoRows", func(t *testing.T) {
		xlsx := NewFile()

		assert.EqualError(t, xlsx.DuplicateRow(sheet, 0), "invalid row number 0")

		if !assert.NoError(t, xlsx.SaveAs(fmt.Sprintf(outFile, "TestDuplicateRow.ZeroWithNoRows"))) {
			t.FailNow()
		}

		val, err := xlsx.GetCellValue(sheet, "A1")
		assert.NoError(t, err)
		assert.Equal(t, "", val)
		val, err = xlsx.GetCellValue(sheet, "B1")
		assert.NoError(t, err)
		assert.Equal(t, "", val)
		val, err = xlsx.GetCellValue(sheet, "A2")
		assert.NoError(t, err)
		assert.Equal(t, "", val)
		val, err = xlsx.GetCellValue(sheet, "B2")
		assert.NoError(t, err)
		assert.Equal(t, "", val)

		assert.NoError(t, err)
		expect := map[string]string{
			"A1": "", "B1": "",
			"A2": "", "B2": "",
		}

		for cell, val := range expect {
			v, err := xlsx.GetCellValue(sheet, cell)
			assert.NoError(t, err)
			if !assert.Equal(t, val, v, cell) {
				t.FailNow()
			}
		}
	})
}

func TestDuplicateRowMiddleRowOfEmptyFile(t *testing.T) {
	const sheet = "Sheet1"
	outFile := filepath.Join("test", "TestDuplicateRow.%s.xlsx")

	t.Run("MiddleRowOfEmptyFile", func(t *testing.T) {
		xlsx := NewFile()

		assert.NoError(t, xlsx.DuplicateRow(sheet, 99))

		if !assert.NoError(t, xlsx.SaveAs(fmt.Sprintf(outFile, "TestDuplicateRow.MiddleRowOfEmptyFile"))) {
			t.FailNow()
		}
		expect := map[string]string{
			"A98":  "",
			"A99":  "",
			"A100": "",
		}
		for cell, val := range expect {
			v, err := xlsx.GetCellValue(sheet, cell)
			assert.NoError(t, err)
			if !assert.Equal(t, val, v, cell) {
				t.FailNow()
			}
		}
	})
}

func TestDuplicateRowWithLargeOffsetToMiddleOfData(t *testing.T) {
	const sheet = "Sheet1"
	outFile := filepath.Join("test", "TestDuplicateRow.%s.xlsx")

	cells := map[string]string{
		"A1": "A1 Value",
		"A2": "A2 Value",
		"A3": "A3 Value",
		"B1": "B1 Value",
		"B2": "B2 Value",
		"B3": "B3 Value",
	}

	newFileWithDefaults := func() *File {
		f := NewFile()
		for cell, val := range cells {
			f.SetCellStr(sheet, cell, val)

		}
		return f
	}

	t.Run("WithLargeOffsetToMiddleOfData", func(t *testing.T) {
		xlsx := newFileWithDefaults()

		assert.NoError(t, xlsx.DuplicateRowTo(sheet, 1, 3))

		if !assert.NoError(t, xlsx.SaveAs(fmt.Sprintf(outFile, "TestDuplicateRow.WithLargeOffsetToMiddleOfData"))) {
			t.FailNow()
		}
		expect := map[string]string{
			"A1": cells["A1"], "B1": cells["B1"],
			"A2": cells["A2"], "B2": cells["B2"],
			"A3": cells["A1"], "B3": cells["B1"],
			"A4": cells["A3"], "B4": cells["B3"],
		}
		for cell, val := range expect {
			v, err := xlsx.GetCellValue(sheet, cell)
			assert.NoError(t, err)
			if !assert.Equal(t, val, v, cell) {
				t.FailNow()
			}
		}
	})
}

func TestDuplicateRowWithLargeOffsetToEmptyRows(t *testing.T) {
	const sheet = "Sheet1"
	outFile := filepath.Join("test", "TestDuplicateRow.%s.xlsx")

	cells := map[string]string{
		"A1": "A1 Value",
		"A2": "A2 Value",
		"A3": "A3 Value",
		"B1": "B1 Value",
		"B2": "B2 Value",
		"B3": "B3 Value",
	}

	newFileWithDefaults := func() *File {
		f := NewFile()
		for cell, val := range cells {
			f.SetCellStr(sheet, cell, val)

		}
		return f
	}

	t.Run("WithLargeOffsetToEmptyRows", func(t *testing.T) {
		xlsx := newFileWithDefaults()

		assert.NoError(t, xlsx.DuplicateRowTo(sheet, 1, 7))

		if !assert.NoError(t, xlsx.SaveAs(fmt.Sprintf(outFile, "TestDuplicateRow.WithLargeOffsetToEmptyRows"))) {
			t.FailNow()
		}
		expect := map[string]string{
			"A1": cells["A1"], "B1": cells["B1"],
			"A2": cells["A2"], "B2": cells["B2"],
			"A3": cells["A3"], "B3": cells["B3"],
			"A7": cells["A1"], "B7": cells["B1"],
		}
		for cell, val := range expect {
			v, err := xlsx.GetCellValue(sheet, cell)
			assert.NoError(t, err)
			if !assert.Equal(t, val, v, cell) {
				t.FailNow()
			}
		}
	})
}

func TestDuplicateRowInsertBefore(t *testing.T) {
	const sheet = "Sheet1"
	outFile := filepath.Join("test", "TestDuplicateRow.%s.xlsx")

	cells := map[string]string{
		"A1": "A1 Value",
		"A2": "A2 Value",
		"A3": "A3 Value",
		"B1": "B1 Value",
		"B2": "B2 Value",
		"B3": "B3 Value",
	}

	newFileWithDefaults := func() *File {
		f := NewFile()
		for cell, val := range cells {
			f.SetCellStr(sheet, cell, val)

		}
		return f
	}

	t.Run("InsertBefore", func(t *testing.T) {
		xlsx := newFileWithDefaults()

		assert.NoError(t, xlsx.DuplicateRowTo(sheet, 2, 1))

		if !assert.NoError(t, xlsx.SaveAs(fmt.Sprintf(outFile, "TestDuplicateRow.InsertBefore"))) {
			t.FailNow()
		}

		expect := map[string]string{
			"A1": cells["A2"], "B1": cells["B2"],
			"A2": cells["A1"], "B2": cells["B1"],
			"A3": cells["A2"], "B3": cells["B2"],
			"A4": cells["A3"], "B4": cells["B3"],
		}
		for cell, val := range expect {
			v, err := xlsx.GetCellValue(sheet, cell)
			assert.NoError(t, err)
			if !assert.Equal(t, val, v, cell) {
				t.FailNow()
			}
		}
	})
}

func TestDuplicateRowInsertBeforeWithLargeOffset(t *testing.T) {
	const sheet = "Sheet1"
	outFile := filepath.Join("test", "TestDuplicateRow.%s.xlsx")

	cells := map[string]string{
		"A1": "A1 Value",
		"A2": "A2 Value",
		"A3": "A3 Value",
		"B1": "B1 Value",
		"B2": "B2 Value",
		"B3": "B3 Value",
	}

	newFileWithDefaults := func() *File {
		f := NewFile()
		for cell, val := range cells {
			f.SetCellStr(sheet, cell, val)

		}
		return f
	}

	t.Run("InsertBeforeWithLargeOffset", func(t *testing.T) {
		xlsx := newFileWithDefaults()

		assert.NoError(t, xlsx.DuplicateRowTo(sheet, 3, 1))

		if !assert.NoError(t, xlsx.SaveAs(fmt.Sprintf(outFile, "TestDuplicateRow.InsertBeforeWithLargeOffset"))) {
			t.FailNow()
		}

		expect := map[string]string{
			"A1": cells["A3"], "B1": cells["B3"],
			"A2": cells["A1"], "B2": cells["B1"],
			"A3": cells["A2"], "B3": cells["B2"],
			"A4": cells["A3"], "B4": cells["B3"],
		}
		for cell, val := range expect {
			v, err := xlsx.GetCellValue(sheet, cell)
			assert.NoError(t, err)
			if !assert.Equal(t, val, v) {
				t.FailNow()
			}
		}
	})
}

func TestDuplicateRowInvalidRownum(t *testing.T) {
	const sheet = "Sheet1"
	outFile := filepath.Join("test", "TestDuplicateRowInvalidRownum.%s.xlsx")

	cells := map[string]string{
		"A1": "A1 Value",
		"A2": "A2 Value",
		"A3": "A3 Value",
		"B1": "B1 Value",
		"B2": "B2 Value",
		"B3": "B3 Value",
	}

	invalidIndexes := []int{-100, -2, -1, 0}

	for _, row := range invalidIndexes {
		name := fmt.Sprintf("%d", row)
		t.Run(name, func(t *testing.T) {
			xlsx := NewFile()
			for col, val := range cells {
				xlsx.SetCellStr(sheet, col, val)
			}

			assert.EqualError(t, xlsx.DuplicateRow(sheet, row), fmt.Sprintf("invalid row number %d", row))

			for col, val := range cells {
				v, err := xlsx.GetCellValue(sheet, col)
				assert.NoError(t, err)
				if !assert.Equal(t, val, v) {
					t.FailNow()
				}
			}
			assert.NoError(t, xlsx.SaveAs(fmt.Sprintf(outFile, name)))
		})
	}

	for _, row1 := range invalidIndexes {
		for _, row2 := range invalidIndexes {
			name := fmt.Sprintf("[%d,%d]", row1, row2)
			t.Run(name, func(t *testing.T) {
				xlsx := NewFile()
				for col, val := range cells {
					xlsx.SetCellStr(sheet, col, val)
				}

				assert.EqualError(t, xlsx.DuplicateRowTo(sheet, row1, row2), fmt.Sprintf("invalid row number %d", row1))

				for col, val := range cells {
					v, err := xlsx.GetCellValue(sheet, col)
					assert.NoError(t, err)
					if !assert.Equal(t, val, v) {
						t.FailNow()
					}
				}
				assert.NoError(t, xlsx.SaveAs(fmt.Sprintf(outFile, name)))
			})
		}
	}
}

func trimSliceSpace(s []string) []string {
	for {
		if len(s) > 0 && s[len(s)-1] == "" {
			s = s[:len(s)-1]
		} else {
			break
		}
	}
	return s
}
