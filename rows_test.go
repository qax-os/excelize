package excelize

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func TestGetRows(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "A1"))
	// Test get rows with unsupported charset shared strings table
	f.SharedStrings = nil
	f.Pkg.Store(defaultXMLPathSharedStrings, MacintoshCyrillicCharset)
	_, err := f.GetRows("Sheet1")
	assert.NoError(t, err)
}

func TestRows(t *testing.T) {
	const sheet2 = "Sheet2"
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)

	// Test get rows with invalid sheet name
	_, err = f.Rows("Sheet:1")
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())

	rows, err := f.Rows(sheet2)
	assert.NoError(t, err)
	var collectedRows [][]string
	for rows.Next() {
		columns, err := rows.Columns()
		assert.NoError(t, err)
		collectedRows = append(collectedRows, trimSliceSpace(columns))
	}
	if !assert.NoError(t, rows.Error()) {
		t.FailNow()
	}
	assert.NoError(t, rows.Close())

	returnedRows, err := f.GetRows(sheet2)
	assert.NoError(t, err)
	for i := range returnedRows {
		returnedRows[i] = trimSliceSpace(returnedRows[i])
	}
	if !assert.Equal(t, collectedRows, returnedRows) {
		t.FailNow()
	}
	assert.NoError(t, f.Close())

	f.Pkg.Store("xl/worksheets/sheet1.xml", nil)
	_, err = f.Rows("Sheet1")
	assert.NoError(t, err)

	// Test reload the file to memory from system temporary directory
	f, err = OpenFile(filepath.Join("test", "Book1.xlsx"), Options{UnzipXMLSizeLimit: 128})
	assert.NoError(t, err)
	value, err := f.GetCellValue("Sheet1", "A19")
	assert.NoError(t, err)
	assert.Equal(t, "Total:", value)
	// Test load shared string table to memory
	err = f.SetCellValue("Sheet1", "A19", "A19")
	assert.NoError(t, err)
	value, err = f.GetCellValue("Sheet1", "A19")
	assert.NoError(t, err)
	assert.Equal(t, "A19", value)
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetRow.xlsx")))
	assert.NoError(t, f.Close())

	// Test rows iterator with unsupported charset shared strings table
	f.SharedStrings = nil
	f.Pkg.Store(defaultXMLPathSharedStrings, MacintoshCyrillicCharset)
	rows, err = f.Rows(sheet2)
	assert.NoError(t, err)
	_, err = rows.Columns()
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
}

func TestRowsIterator(t *testing.T) {
	sheetName, rowCount, expectedNumRow := "Sheet2", 0, 11
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	require.NoError(t, err)

	rows, err := f.Rows(sheetName)
	require.NoError(t, err)

	for rows.Next() {
		rowCount++
		require.True(t, rowCount <= expectedNumRow, "rowCount is greater than expected")
	}
	assert.Equal(t, expectedNumRow, rowCount)
	assert.NoError(t, rows.Close())
	assert.NoError(t, f.Close())

	// Valued cell sparse distribution test
	f, sheetName, rowCount, expectedNumRow = NewFile(), "Sheet1", 0, 3
	cells := []string{"C1", "E1", "A3", "B3", "C3", "D3", "E3"}
	for _, cell := range cells {
		assert.NoError(t, f.SetCellValue(sheetName, cell, 1))
	}
	rows, err = f.Rows(sheetName)
	require.NoError(t, err)
	for rows.Next() {
		rowCount++
		require.True(t, rowCount <= expectedNumRow, "rowCount is greater than expected")
	}
	assert.Equal(t, expectedNumRow, rowCount)
}

func TestRowsGetRowOpts(t *testing.T) {
	sheetName := "Sheet2"
	expectedRowStyleID1 := RowOpts{Height: 17.0, Hidden: false, StyleID: 1}
	expectedRowStyleID2 := RowOpts{Height: 17.0, Hidden: false, StyleID: 0}
	expectedRowStyleID3 := RowOpts{Height: 17.0, Hidden: false, StyleID: 2}
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	require.NoError(t, err)

	rows, err := f.Rows(sheetName)
	require.NoError(t, err)

	assert.Equal(t, true, rows.Next())
	_, err = rows.Columns()
	require.NoError(t, err)
	rowOpts := rows.GetRowOpts()
	assert.Equal(t, expectedRowStyleID1, rowOpts)
	assert.Equal(t, true, rows.Next())
	rowOpts = rows.GetRowOpts()
	assert.Equal(t, expectedRowStyleID2, rowOpts)
	assert.Equal(t, true, rows.Next())
	_, err = rows.Columns()
	require.NoError(t, err)
	rowOpts = rows.GetRowOpts()
	assert.Equal(t, expectedRowStyleID3, rowOpts)
}

func TestRowsError(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	_, err = f.Rows("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	assert.NoError(t, f.Close())
}

func TestRowHeight(t *testing.T) {
	f := NewFile()
	sheet1 := f.GetSheetName(0)

	assert.EqualError(t, f.SetRowHeight(sheet1, 0, defaultRowHeightPixels+1.0), newInvalidRowNumberError(0).Error())

	_, err := f.GetRowHeight("Sheet1", 0)
	assert.EqualError(t, err, newInvalidRowNumberError(0).Error())

	assert.NoError(t, f.SetRowHeight(sheet1, 1, 111.0))
	height, err := f.GetRowHeight(sheet1, 1)
	assert.NoError(t, err)
	assert.Equal(t, 111.0, height)

	// Test set row height overflow max row height limit
	assert.EqualError(t, f.SetRowHeight(sheet1, 4, MaxRowHeight+1), ErrMaxRowHeight.Error())

	// Test get row height that rows index over exists rows
	height, err = f.GetRowHeight(sheet1, 5)
	assert.NoError(t, err)
	assert.Equal(t, defaultRowHeight, height)

	// Test get row height that rows heights haven't changed
	height, err = f.GetRowHeight(sheet1, 3)
	assert.NoError(t, err)
	assert.Equal(t, defaultRowHeight, height)

	// Test set and get row height on not exists worksheet
	assert.EqualError(t, f.SetRowHeight("SheetN", 1, 111.0), "sheet SheetN does not exist")
	_, err = f.GetRowHeight("SheetN", 3)
	assert.EqualError(t, err, "sheet SheetN does not exist")

	// Test set row height with invalid sheet name
	assert.EqualError(t, f.SetRowHeight("Sheet:1", 1, 10.0), ErrSheetNameInvalid.Error())

	// Test get row height with invalid sheet name
	_, err = f.GetRowHeight("Sheet:1", 3)
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())

	// Test get row height with custom default row height
	assert.NoError(t, f.SetSheetProps(sheet1, &SheetPropsOptions{
		DefaultRowHeight: float64Ptr(30.0),
		CustomHeight:     boolPtr(true),
	}))
	height, err = f.GetRowHeight(sheet1, 100)
	assert.NoError(t, err)
	assert.Equal(t, 30.0, height)

	// Test set row height with custom default row height with prepare XML
	assert.NoError(t, f.SetCellValue(sheet1, "A10", "A10"))

	_, err = f.NewSheet("Sheet2")
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellValue("Sheet2", "A2", true))
	height, err = f.GetRowHeight("Sheet2", 1)
	assert.NoError(t, err)
	assert.Equal(t, 15.0, height)

	err = f.SaveAs(filepath.Join("test", "TestRowHeight.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.Equal(t, 0.0, convertColWidthToPixels(0))
}

func TestColumns(t *testing.T) {
	f := NewFile()
	rows, err := f.Rows("Sheet1")
	assert.NoError(t, err)

	rows.decoder = f.xmlNewDecoder(bytes.NewReader([]byte(`<worksheet><sheetData><row r="2"><c r="A1" t="s"><v>1</v></c></row></sheetData></worksheet>`)))
	_, err = rows.Columns()
	assert.NoError(t, err)
	rows.decoder = f.xmlNewDecoder(bytes.NewReader([]byte(`<worksheet><sheetData><row r="2"><c r="A1" t="s"><v>1</v></c></row></sheetData></worksheet>`)))
	rows.curRow = 1
	_, err = rows.Columns()
	assert.NoError(t, err)

	rows.decoder = f.xmlNewDecoder(bytes.NewReader([]byte(`<worksheet><sheetData><row r="A"><c r="A1" t="s"><v>1</v></c></row><row r="A"><c r="2" t="inlineStr"><is><t>B</t></is></c></row></sheetData></worksheet>`)))
	assert.True(t, rows.Next())
	_, err = rows.Columns()
	assert.EqualError(t, err, `strconv.Atoi: parsing "A": invalid syntax`)

	rows.decoder = f.xmlNewDecoder(bytes.NewReader([]byte(`<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>1</v></c></row><row r="A"><c r="2" t="inlineStr"><is><t>B</t></is></c></row></sheetData></worksheet>`)))
	_, err = rows.Columns()
	assert.NoError(t, err)

	rows.decoder = f.xmlNewDecoder(bytes.NewReader([]byte(`<worksheet><sheetData><row r="1"><c r="A" t="s"><v>1</v></c></row></sheetData></worksheet>`)))
	assert.True(t, rows.Next())
	_, err = rows.Columns()
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), err)

	// Test token is nil
	rows.decoder = f.xmlNewDecoder(bytes.NewReader(nil))
	_, err = rows.Columns()
	assert.NoError(t, err)
}

func TestSharedStringsReader(t *testing.T) {
	f := NewFile()
	// Test read shared string with unsupported charset
	f.Pkg.Store(defaultXMLPathSharedStrings, MacintoshCyrillicCharset)
	_, err := f.sharedStringsReader()
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	// Test read shared strings with unsupported charset content types
	f = NewFile()
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	_, err = f.sharedStringsReader()
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	// Test read shared strings with unsupported charset workbook relationships
	f = NewFile()
	f.Relationships.Delete(defaultXMLPathWorkbookRels)
	f.Pkg.Store(defaultXMLPathWorkbookRels, MacintoshCyrillicCharset)
	_, err = f.sharedStringsReader()
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
}

func TestRowVisibility(t *testing.T) {
	f, err := prepareTestBook1()
	assert.NoError(t, err)
	_, err = f.NewSheet("Sheet3")
	assert.NoError(t, err)
	assert.NoError(t, f.SetRowVisible("Sheet3", 2, false))
	assert.NoError(t, f.SetRowVisible("Sheet3", 2, true))
	visible, err := f.GetRowVisible("Sheet3", 2)
	assert.Equal(t, true, visible)
	assert.NoError(t, err)
	visible, err = f.GetRowVisible("Sheet3", 25)
	assert.Equal(t, false, visible)
	assert.NoError(t, err)
	assert.EqualError(t, f.SetRowVisible("Sheet3", 0, true), newInvalidRowNumberError(0).Error())
	assert.EqualError(t, f.SetRowVisible("SheetN", 2, false), "sheet SheetN does not exist")
	// Test set row visibility with invalid sheet name
	assert.EqualError(t, f.SetRowVisible("Sheet:1", 1, false), ErrSheetNameInvalid.Error())

	visible, err = f.GetRowVisible("Sheet3", 0)
	assert.Equal(t, false, visible)
	assert.EqualError(t, err, newInvalidRowNumberError(0).Error())
	_, err = f.GetRowVisible("SheetN", 1)
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Test get row visibility with invalid sheet name
	_, err = f.GetRowVisible("Sheet:1", 1)
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestRowVisibility.xlsx")))
}

func TestRemoveRow(t *testing.T) {
	f := NewFile()
	sheet1 := f.GetSheetName(0)
	r, err := f.workSheetReader(sheet1)
	assert.NoError(t, err)
	const (
		colCount = 10
		rowCount = 10
	)
	assert.NoError(t, fillCells(f, sheet1, colCount, rowCount))

	assert.NoError(t, f.SetCellHyperLink(sheet1, "A5", "https://github.com/xuri/excelize", "External"))

	assert.EqualError(t, f.RemoveRow(sheet1, -1), newInvalidRowNumberError(-1).Error())

	assert.EqualError(t, f.RemoveRow(sheet1, 0), newInvalidRowNumberError(0).Error())

	assert.NoError(t, f.RemoveRow(sheet1, 4))
	if !assert.Len(t, r.SheetData.Row, rowCount-1) {
		t.FailNow()
	}

	assert.NoError(t, f.MergeCell(sheet1, "B3", "B5"))

	assert.NoError(t, f.RemoveRow(sheet1, 2))
	if !assert.Len(t, r.SheetData.Row, rowCount-2) {
		t.FailNow()
	}

	assert.NoError(t, f.RemoveRow(sheet1, 4))
	if !assert.Len(t, r.SheetData.Row, rowCount-3) {
		t.FailNow()
	}

	err = f.AutoFilter(sheet1, "A2:A2", []AutoFilterOptions{{Column: "A", Expression: "x != blanks"}})
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.RemoveRow(sheet1, 1))
	if !assert.Len(t, r.SheetData.Row, rowCount-4) {
		t.FailNow()
	}

	assert.NoError(t, f.RemoveRow(sheet1, 2))
	if !assert.Len(t, r.SheetData.Row, rowCount-5) {
		t.FailNow()
	}

	assert.NoError(t, f.RemoveRow(sheet1, 1))
	if !assert.Len(t, r.SheetData.Row, rowCount-6) {
		t.FailNow()
	}

	assert.NoError(t, f.RemoveRow(sheet1, 10))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestRemoveRow.xlsx")))

	f = NewFile()
	assert.NoError(t, f.MergeCell("Sheet1", "A1", "C1"))
	assert.NoError(t, f.MergeCell("Sheet1", "A2", "C2"))
	assert.NoError(t, f.RemoveRow("Sheet1", 1))
	mergedCells, err := f.GetMergeCells("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, "A1", mergedCells[0].GetStartAxis())
	assert.Equal(t, "C1", mergedCells[0].GetEndAxis())

	// Test remove row on not exist worksheet
	assert.EqualError(t, f.RemoveRow("SheetN", 1), "sheet SheetN does not exist")
	// Test remove row with invalid sheet name
	assert.EqualError(t, f.RemoveRow("Sheet:1", 1), ErrSheetNameInvalid.Error())
}

func TestInsertRows(t *testing.T) {
	f := NewFile()
	sheet1 := f.GetSheetName(0)
	r, err := f.workSheetReader(sheet1)
	assert.NoError(t, err)
	const (
		colCount = 10
		rowCount = 10
	)
	assert.NoError(t, fillCells(f, sheet1, colCount, rowCount))

	assert.NoError(t, f.SetCellHyperLink(sheet1, "A5", "https://github.com/xuri/excelize", "External"))

	assert.NoError(t, f.InsertRows(sheet1, 1, 1))
	if !assert.Len(t, r.SheetData.Row, rowCount+1) {
		t.FailNow()
	}

	assert.NoError(t, f.InsertRows(sheet1, 4, 1))
	if !assert.Len(t, r.SheetData.Row, rowCount+2) {
		t.FailNow()
	}

	assert.NoError(t, f.InsertRows(sheet1, 4, 2))
	if !assert.Len(t, r.SheetData.Row, rowCount+4) {
		t.FailNow()
	}
	// Test insert rows with invalid sheet name
	assert.EqualError(t, f.InsertRows("Sheet:1", 1, 1), ErrSheetNameInvalid.Error())

	assert.EqualError(t, f.InsertRows(sheet1, -1, 1), newInvalidRowNumberError(-1).Error())
	assert.EqualError(t, f.InsertRows(sheet1, 0, 1), newInvalidRowNumberError(0).Error())
	assert.EqualError(t, f.InsertRows(sheet1, 4, 0), ErrParameterInvalid.Error())
	assert.EqualError(t, f.InsertRows(sheet1, 4, TotalRows), ErrMaxRows.Error())
	assert.EqualError(t, f.InsertRows(sheet1, 4, TotalRows-5), ErrMaxRows.Error())
	assert.EqualError(t, f.InsertRows(sheet1, TotalRows, 1), ErrMaxRows.Error())

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestInsertRows.xlsx")))
}

// Test internal structure state after insert operations. It is important
// for insert workflow to be constant to avoid side effect with functions
// related to internal structure.
func TestInsertRowsInEmptyFile(t *testing.T) {
	f := NewFile()
	sheet1 := f.GetSheetName(0)
	r, err := f.workSheetReader(sheet1)
	assert.NoError(t, err)
	assert.NoError(t, f.InsertRows(sheet1, 1, 1))
	assert.Len(t, r.SheetData.Row, 0)
	assert.NoError(t, f.InsertRows(sheet1, 2, 1))
	assert.Len(t, r.SheetData.Row, 0)
	assert.NoError(t, f.InsertRows(sheet1, 99, 1))
	assert.Len(t, r.SheetData.Row, 0)
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestInsertRowInEmptyFile.xlsx")))
}

func prepareTestBook2() (*File, error) {
	f := NewFile()
	for cell, val := range map[string]string{
		"A1": "A1 Value",
		"A2": "A2 Value",
		"A3": "A3 Value",
		"B1": "B1 Value",
		"B2": "B2 Value",
		"B3": "B3 Value",
	} {
		if err := f.SetCellStr("Sheet1", cell, val); err != nil {
			return f, err
		}
	}
	return f, nil
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
		f := NewFile()
		assert.NoError(t, f.SetCellStr(sheet, "A1", cells["A1"]))
		assert.NoError(t, f.SetCellStr(sheet, "B1", cells["B1"]))

		assert.NoError(t, f.DuplicateRow(sheet, 1))
		if !assert.NoError(t, f.SaveAs(fmt.Sprintf(outFile, "FromSingleRow_1"))) {
			t.FailNow()
		}
		expect := map[string]string{
			"A1": cells["A1"], "B1": cells["B1"],
			"A2": cells["A1"], "B2": cells["B1"],
		}
		for cell, val := range expect {
			v, err := f.GetCellValue(sheet, cell)
			assert.NoError(t, err)
			if !assert.Equal(t, val, v, cell) {
				t.FailNow()
			}
		}

		assert.NoError(t, f.DuplicateRow(sheet, 2))
		if !assert.NoError(t, f.SaveAs(fmt.Sprintf(outFile, "FromSingleRow_2"))) {
			t.FailNow()
		}
		expect = map[string]string{
			"A1": cells["A1"], "B1": cells["B1"],
			"A2": cells["A1"], "B2": cells["B1"],
			"A3": cells["A1"], "B3": cells["B1"],
		}
		for cell, val := range expect {
			v, err := f.GetCellValue(sheet, cell)
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
		f := NewFile()
		assert.NoError(t, f.SetCellStr(sheet, "A1", cells["A1"]))
		assert.NoError(t, f.SetCellStr(sheet, "B1", cells["B1"]))

		assert.NoError(t, f.DuplicateRow(sheet, 1))

		assert.NoError(t, f.SetCellStr(sheet, "A2", cells["A2"]))
		assert.NoError(t, f.SetCellStr(sheet, "B2", cells["B2"]))

		if !assert.NoError(t, f.SaveAs(fmt.Sprintf(outFile, "UpdateDuplicatedRows"))) {
			t.FailNow()
		}
		expect := map[string]string{
			"A1": cells["A1"], "B1": cells["B1"],
			"A2": cells["A2"], "B2": cells["B2"],
		}
		for cell, val := range expect {
			v, err := f.GetCellValue(sheet, cell)
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
	t.Run("FirstOfMultipleRows", func(t *testing.T) {
		f, err := prepareTestBook2()
		assert.NoError(t, err)
		assert.NoError(t, f.DuplicateRow(sheet, 1))

		if !assert.NoError(t, f.SaveAs(fmt.Sprintf(outFile, "FirstOfMultipleRows"))) {
			t.FailNow()
		}
		expect := map[string]string{
			"A1": cells["A1"], "B1": cells["B1"],
			"A2": cells["A1"], "B2": cells["B1"],
			"A3": cells["A2"], "B3": cells["B2"],
			"A4": cells["A3"], "B4": cells["B3"],
		}
		for cell, val := range expect {
			v, err := f.GetCellValue(sheet, cell)
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
		f := NewFile()

		assert.EqualError(t, f.DuplicateRow(sheet, 0), newInvalidRowNumberError(0).Error())

		if !assert.NoError(t, f.SaveAs(fmt.Sprintf(outFile, "ZeroWithNoRows"))) {
			t.FailNow()
		}

		val, err := f.GetCellValue(sheet, "A1")
		assert.NoError(t, err)
		assert.Equal(t, "", val)
		val, err = f.GetCellValue(sheet, "B1")
		assert.NoError(t, err)
		assert.Equal(t, "", val)
		val, err = f.GetCellValue(sheet, "A2")
		assert.NoError(t, err)
		assert.Equal(t, "", val)
		val, err = f.GetCellValue(sheet, "B2")
		assert.NoError(t, err)
		assert.Equal(t, "", val)

		assert.NoError(t, err)
		expect := map[string]string{
			"A1": "", "B1": "",
			"A2": "", "B2": "",
		}

		for cell, val := range expect {
			v, err := f.GetCellValue(sheet, cell)
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
		f := NewFile()

		assert.NoError(t, f.DuplicateRow(sheet, 99))

		if !assert.NoError(t, f.SaveAs(fmt.Sprintf(outFile, "MiddleRowOfEmptyFile"))) {
			t.FailNow()
		}
		expect := map[string]string{
			"A98":  "",
			"A99":  "",
			"A100": "",
		}
		for cell, val := range expect {
			v, err := f.GetCellValue(sheet, cell)
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
	t.Run("WithLargeOffsetToMiddleOfData", func(t *testing.T) {
		f, err := prepareTestBook2()
		assert.NoError(t, err)
		assert.NoError(t, f.DuplicateRowTo(sheet, 1, 3))

		if !assert.NoError(t, f.SaveAs(fmt.Sprintf(outFile, "WithLargeOffsetToMiddleOfData"))) {
			t.FailNow()
		}
		expect := map[string]string{
			"A1": cells["A1"], "B1": cells["B1"],
			"A2": cells["A2"], "B2": cells["B2"],
			"A3": cells["A1"], "B3": cells["B1"],
			"A4": cells["A3"], "B4": cells["B3"],
		}
		for cell, val := range expect {
			v, err := f.GetCellValue(sheet, cell)
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
	t.Run("WithLargeOffsetToEmptyRows", func(t *testing.T) {
		f, err := prepareTestBook2()
		assert.NoError(t, err)
		assert.NoError(t, f.DuplicateRowTo(sheet, 1, 7))

		if !assert.NoError(t, f.SaveAs(fmt.Sprintf(outFile, "WithLargeOffsetToEmptyRows"))) {
			t.FailNow()
		}
		expect := map[string]string{
			"A1": cells["A1"], "B1": cells["B1"],
			"A2": cells["A2"], "B2": cells["B2"],
			"A3": cells["A3"], "B3": cells["B3"],
			"A7": cells["A1"], "B7": cells["B1"],
		}
		for cell, val := range expect {
			v, err := f.GetCellValue(sheet, cell)
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
	t.Run("InsertBefore", func(t *testing.T) {
		f, err := prepareTestBook2()
		assert.NoError(t, err)
		assert.NoError(t, f.DuplicateRowTo(sheet, 2, 1))
		assert.NoError(t, f.DuplicateRowTo(sheet, 10, 4))

		if !assert.NoError(t, f.SaveAs(fmt.Sprintf(outFile, "InsertBefore"))) {
			t.FailNow()
		}

		expect := map[string]string{
			"A1": cells["A2"], "B1": cells["B2"],
			"A2": cells["A1"], "B2": cells["B1"],
			"A3": cells["A2"], "B3": cells["B2"],
			"A5": cells["A3"], "B5": cells["B3"],
		}
		for cell, val := range expect {
			v, err := f.GetCellValue(sheet, cell)
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
	t.Run("InsertBeforeWithLargeOffset", func(t *testing.T) {
		f, err := prepareTestBook2()
		assert.NoError(t, err)
		assert.NoError(t, f.DuplicateRowTo(sheet, 3, 1))

		if !assert.NoError(t, f.SaveAs(fmt.Sprintf(outFile, "InsertBeforeWithLargeOffset"))) {
			t.FailNow()
		}

		expect := map[string]string{
			"A1": cells["A3"], "B1": cells["B3"],
			"A2": cells["A1"], "B2": cells["B1"],
			"A3": cells["A2"], "B3": cells["B2"],
			"A4": cells["A3"], "B4": cells["B3"],
		}
		for cell, val := range expect {
			v, err := f.GetCellValue(sheet, cell)
			assert.NoError(t, err)
			if !assert.Equal(t, val, v) {
				t.FailNow()
			}
		}
	})
}

func TestDuplicateRowInsertBeforeWithMergeCells(t *testing.T) {
	const sheet = "Sheet1"
	outFile := filepath.Join("test", "TestDuplicateRow.%s.xlsx")
	t.Run("InsertBeforeWithLargeOffset", func(t *testing.T) {
		f, err := prepareTestBook2()
		assert.NoError(t, err)
		assert.NoError(t, f.MergeCell(sheet, "B2", "C2"))
		assert.NoError(t, f.MergeCell(sheet, "C6", "C8"))

		assert.NoError(t, f.DuplicateRowTo(sheet, 2, 1))
		assert.NoError(t, f.DuplicateRowTo(sheet, 1, 8))

		if !assert.NoError(t, f.SaveAs(fmt.Sprintf(outFile, "InsertBeforeWithMergeCells"))) {
			t.FailNow()
		}

		expect := []MergeCell{
			{"B3:C3", "B2 Value"},
			{"C7:C10", ""},
			{"B1:C1", "B2 Value"},
		}

		mergeCells, err := f.GetMergeCells(sheet)
		assert.NoError(t, err)
		for idx, val := range expect {
			if !assert.Equal(t, val, mergeCells[idx]) {
				t.FailNow()
			}
		}
	})
}

func TestDuplicateRowInvalidRowNum(t *testing.T) {
	const sheet = "Sheet1"
	outFile := filepath.Join("test", "TestDuplicateRow.InvalidRowNum.%s.xlsx")

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
			f := NewFile()
			for col, val := range cells {
				assert.NoError(t, f.SetCellStr(sheet, col, val))
			}

			assert.EqualError(t, f.DuplicateRow(sheet, row), newInvalidRowNumberError(row).Error())

			for col, val := range cells {
				v, err := f.GetCellValue(sheet, col)
				assert.NoError(t, err)
				if !assert.Equal(t, val, v) {
					t.FailNow()
				}
			}
			assert.NoError(t, f.SaveAs(fmt.Sprintf(outFile, name)))
		})
	}

	for _, row1 := range invalidIndexes {
		for _, row2 := range invalidIndexes {
			name := fmt.Sprintf("[%d,%d]", row1, row2)
			t.Run(name, func(t *testing.T) {
				f := NewFile()
				for col, val := range cells {
					assert.NoError(t, f.SetCellStr(sheet, col, val))
				}

				assert.EqualError(t, f.DuplicateRowTo(sheet, row1, row2), newInvalidRowNumberError(row1).Error())

				for col, val := range cells {
					v, err := f.GetCellValue(sheet, col)
					assert.NoError(t, err)
					if !assert.Equal(t, val, v) {
						t.FailNow()
					}
				}
				assert.NoError(t, f.SaveAs(fmt.Sprintf(outFile, name)))
			})
		}
	}
}

func TestDuplicateRow(t *testing.T) {
	f := NewFile()
	// Test duplicate row with invalid sheet name
	assert.EqualError(t, f.DuplicateRowTo("Sheet:1", 1, 2), ErrSheetNameInvalid.Error())

	f = NewFile()
	assert.NoError(t, f.SetDefinedName(&DefinedName{
		Name:     "Amount",
		RefersTo: "Sheet1!$B$1",
	}))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "Amount+C1"))
	assert.NoError(t, f.SetCellValue("Sheet1", "A10", "A10"))

	format, err := f.NewConditionalStyle(&Style{Font: &Font{Color: "9A0511"}, Fill: Fill{Type: "pattern", Color: []string{"FEC7CE"}, Pattern: 1}})
	assert.NoError(t, err)

	expected := []ConditionalFormatOptions{
		{Type: "cell", Criteria: "greater than", Format: &format, Value: "0"},
	}
	assert.NoError(t, f.SetConditionalFormat("Sheet1", "A1", expected))

	dv := NewDataValidation(true)
	dv.Sqref = "A1"
	assert.NoError(t, dv.SetDropList([]string{"1", "2", "3"}))
	assert.NoError(t, f.AddDataValidation("Sheet1", dv))
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).DataValidations.DataValidation[0].Sqref = "A1"

	assert.NoError(t, f.DuplicateRowTo("Sheet1", 1, 10))
	formula, err := f.GetCellFormula("Sheet1", "A10")
	assert.NoError(t, err)
	assert.Equal(t, "Amount+C10", formula)
	value, err := f.GetCellValue("Sheet1", "A11")
	assert.NoError(t, err)
	assert.Equal(t, "A10", value)

	cfs, err := f.GetConditionalFormats("Sheet1")
	assert.NoError(t, err)
	assert.Len(t, cfs, 2)
	assert.Equal(t, expected, cfs["A10:A10"])

	dvs, err := f.GetDataValidations("Sheet1")
	assert.NoError(t, err)
	assert.Len(t, dvs, 2)
	assert.Equal(t, "A10:A10", dvs[1].Sqref)

	// Test duplicate data validation with row number exceeds maximum limit
	assert.Equal(t, ErrMaxRows, f.duplicateDataValidations(ws.(*xlsxWorksheet), "Sheet1", 1, TotalRows+1))
	// Test duplicate data validation with invalid range reference
	ws.(*xlsxWorksheet).DataValidations.DataValidation[0].Sqref = "A"
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), f.duplicateDataValidations(ws.(*xlsxWorksheet), "Sheet1", 1, 10))

	// Test duplicate conditional formatting with row number exceeds maximum limit
	assert.Equal(t, ErrMaxRows, f.duplicateConditionalFormat(ws.(*xlsxWorksheet), "Sheet1", 1, TotalRows+1))
	// Test duplicate conditional formatting with invalid range reference
	ws.(*xlsxWorksheet).ConditionalFormatting[0].SQRef = "A"
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), f.duplicateConditionalFormat(ws.(*xlsxWorksheet), "Sheet1", 1, 10))
}

func TestDuplicateRowTo(t *testing.T) {
	f, sheetName := NewFile(), "Sheet1"
	// Test duplicate row with invalid target row number
	assert.Equal(t, nil, f.DuplicateRowTo(sheetName, 1, 0))
	// Test duplicate row with equal source and target row number
	assert.Equal(t, nil, f.DuplicateRowTo(sheetName, 1, 1))
	// Test duplicate row on the blank worksheet
	assert.Equal(t, nil, f.DuplicateRowTo(sheetName, 1, 2))
	// Test duplicate row on the worksheet with illegal cell reference
	f.Sheet.Store("xl/worksheets/sheet1.xml", &xlsxWorksheet{
		MergeCells: &xlsxMergeCells{Cells: []*xlsxMergeCell{{Ref: "A:B1"}}},
	})
	assert.EqualError(t, f.DuplicateRowTo(sheetName, 1, 2), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	// Test duplicate row on not exists worksheet
	assert.EqualError(t, f.DuplicateRowTo("SheetN", 1, 2), "sheet SheetN does not exist")
	// Test duplicate row with invalid sheet name
	assert.EqualError(t, f.DuplicateRowTo("Sheet:1", 1, 2), ErrSheetNameInvalid.Error())
}

func TestDuplicateMergeCells(t *testing.T) {
	f := File{}
	ws := &xlsxWorksheet{MergeCells: &xlsxMergeCells{
		Cells: []*xlsxMergeCell{{Ref: "A1:-"}},
	}}
	assert.EqualError(t, f.duplicateMergeCells(ws, "Sheet1", 0, 0), `cannot convert cell "-" to coordinates: invalid cell name "-"`)
	ws.MergeCells.Cells[0].Ref = "A1:B1"
	assert.EqualError(t, f.duplicateMergeCells(ws, "SheetN", 1, 2), "sheet SheetN does not exist")
}

func TestGetValueFromInlineStr(t *testing.T) {
	c := &xlsxC{T: "inlineStr"}
	f := NewFile()
	d := &xlsxSST{}
	val, err := c.getValueFrom(f, d, false)
	assert.NoError(t, err)
	assert.Equal(t, "", val)
}

func TestGetValueFromNumber(t *testing.T) {
	c := &xlsxC{T: "n"}
	f := NewFile()
	d := &xlsxSST{}
	for input, expected := range map[string]string{
		"2.2.":                     "2.2.",
		"1.1000000000000001":       "1.1",
		"2.2200000000000002":       "2.22",
		"28.552":                   "28.552",
		"27.399000000000001":       "27.399",
		"26.245999999999999":       "26.246",
		"2422.3000000000002":       "2422.3",
		"2.220000ddsf0000000002-r": "2.220000ddsf0000000002-r",
	} {
		c.V = input
		val, err := c.getValueFrom(f, d, false)
		assert.NoError(t, err)
		assert.Equal(t, expected, val)
	}
}

func TestErrSheetNotExistError(t *testing.T) {
	assert.Equal(t, "sheet Sheet1 does not exist", ErrSheetNotExist{"Sheet1"}.Error())
}

func TestCheckRow(t *testing.T) {
	f := NewFile()
	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(xml.Header+`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ><sheetData><row r="2"><c><v>1</v></c><c r="F2"><v>2</v></c><c><v>3</v></c><c><v>4</v></c><c r="M2"><v>5</v></c></row></sheetData></worksheet>`))
	_, err := f.GetRows("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", false))
	f = NewFile()
	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(xml.Header+`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ><sheetData><row r="2"><c><v>1</v></c><c r="-"><v>2</v></c><c><v>3</v></c><c><v>4</v></c><c r="M2"><v>5</v></c></row></sheetData></worksheet>`))
	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.checked.Delete("xl/worksheets/sheet1.xml")
	assert.EqualError(t, f.SetCellValue("Sheet1", "A1", false), newCellNameToCoordinatesError("-", newInvalidCellNameError("-")).Error())
}

func TestSetRowStyle(t *testing.T) {
	f := NewFile()
	style1, err := f.NewStyle(&Style{Fill: Fill{Type: "pattern", Color: []string{"63BE7B"}, Pattern: 1}})
	assert.NoError(t, err)
	style2, err := f.NewStyle(&Style{Fill: Fill{Type: "pattern", Color: []string{"E0EBF5"}, Pattern: 1}})
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "B2", "B2", style1))
	assert.EqualError(t, f.SetRowStyle("Sheet1", 5, -1, style2), newInvalidRowNumberError(-1).Error())
	assert.EqualError(t, f.SetRowStyle("Sheet1", 1, TotalRows+1, style2), ErrMaxRows.Error())
	// Test set row style with invalid style ID
	assert.EqualError(t, f.SetRowStyle("Sheet1", 1, 1, -1), newInvalidStyleID(-1).Error())
	// Test set row style with not exists style ID
	assert.EqualError(t, f.SetRowStyle("Sheet1", 1, 1, 10), newInvalidStyleID(10).Error())
	assert.EqualError(t, f.SetRowStyle("SheetN", 1, 1, style2), "sheet SheetN does not exist")
	// Test set row style with invalid sheet name
	assert.EqualError(t, f.SetRowStyle("Sheet:1", 1, 1, 0), ErrSheetNameInvalid.Error())
	assert.NoError(t, f.SetRowStyle("Sheet1", 5, 1, style2))
	cellStyleID, err := f.GetCellStyle("Sheet1", "B2")
	assert.NoError(t, err)
	assert.Equal(t, style2, cellStyleID)
	// Test cell inheritance rows style
	assert.NoError(t, f.SetCellValue("Sheet1", "C1", nil))
	cellStyleID, err = f.GetCellStyle("Sheet1", "C1")
	assert.NoError(t, err)
	assert.Equal(t, style2, cellStyleID)
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetRowStyle.xlsx")))
	// Test set row style with unsupported charset style sheet
	f.Styles = nil
	f.Pkg.Store(defaultXMLPathStyles, MacintoshCyrillicCharset)
	assert.EqualError(t, f.SetRowStyle("Sheet1", 1, 1, cellStyleID), "XML syntax error on line 1: invalid UTF-8")
}

func TestSetRowHeight(t *testing.T) {
	f := NewFile()
	// Test hidden row by set row height to 0
	assert.NoError(t, f.SetRowHeight("Sheet1", 2, 0))
	ht, err := f.GetRowHeight("Sheet1", 2)
	assert.NoError(t, err)
	assert.Empty(t, ht)
	// Test unset custom row height
	assert.NoError(t, f.SetRowHeight("Sheet1", 2, -1))
	ht, err = f.GetRowHeight("Sheet1", 2)
	assert.NoError(t, err)
	assert.Equal(t, defaultRowHeight, ht)
	// Test set row height with invalid height value
	assert.Equal(t, ErrParameterInvalid, f.SetRowHeight("Sheet1", 2, -2))
}

func TestNumberFormats(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	cells := make([][]string, 0)
	cols, err := f.Cols("Sheet2")
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	for cols.Next() {
		col, err := cols.Rows()
		assert.NoError(t, err)
		if err != nil {
			break
		}
		cells = append(cells, col)
	}
	assert.Equal(t, []string{"", "200", "450", "200", "510", "315", "127", "89", "348", "53", "37"}, cells[3])
	assert.NoError(t, f.Close())

	f = NewFile()
	numFmt1, err := f.NewStyle(&Style{NumFmt: 1})
	assert.NoError(t, err)
	numFmt2, err := f.NewStyle(&Style{NumFmt: 2})
	assert.NoError(t, err)
	numFmt3, err := f.NewStyle(&Style{NumFmt: 3})
	assert.NoError(t, err)
	numFmt9, err := f.NewStyle(&Style{NumFmt: 9})
	assert.NoError(t, err)
	numFmt10, err := f.NewStyle(&Style{NumFmt: 10})
	assert.NoError(t, err)
	numFmt21, err := f.NewStyle(&Style{NumFmt: 21})
	assert.NoError(t, err)
	numFmt37, err := f.NewStyle(&Style{NumFmt: 37})
	assert.NoError(t, err)
	numFmt38, err := f.NewStyle(&Style{NumFmt: 38})
	assert.NoError(t, err)
	numFmt39, err := f.NewStyle(&Style{NumFmt: 39})
	assert.NoError(t, err)
	numFmt40, err := f.NewStyle(&Style{NumFmt: 40})
	assert.NoError(t, err)
	for _, cases := range [][]interface{}{
		{"A1", numFmt1, 8.8888666665555493e+19, "88888666665555500000"},
		{"A2", numFmt1, 8.8888666665555487, "9"},
		{"A3", numFmt2, 8.8888666665555493e+19, "88888666665555500000.00"},
		{"A4", numFmt2, 8.8888666665555487, "8.89"},
		{"A5", numFmt3, 8.8888666665555493e+19, "88,888,666,665,555,500,000"},
		{"A6", numFmt3, 8.8888666665555487, "9"},
		{"A7", numFmt3, 123, "123"},
		{"A8", numFmt3, -1234, "-1,234"},
		{"A9", numFmt9, 8.8888666665555493e+19, "8888866666555550000000%"},
		{"A10", numFmt9, -8.8888666665555493e+19, "-8888866666555550000000%"},
		{"A11", numFmt9, 8.8888666665555487, "889%"},
		{"A12", numFmt9, -8.8888666665555487, "-889%"},
		{"A13", numFmt10, 8.8888666665555493e+19, "8888866666555550000000.00%"},
		{"A14", numFmt10, -8.8888666665555493e+19, "-8888866666555550000000.00%"},
		{"A15", numFmt10, 8.8888666665555487, "888.89%"},
		{"A16", numFmt10, -8.8888666665555487, "-888.89%"},
		{"A17", numFmt37, 8.8888666665555493e+19, "88,888,666,665,555,500,000 "},
		{"A18", numFmt37, -8.8888666665555493e+19, "(88,888,666,665,555,500,000)"},
		{"A19", numFmt37, 8.8888666665555487, "9 "},
		{"A20", numFmt37, -8.8888666665555487, "(9)"},
		{"A21", numFmt38, 8.8888666665555493e+19, "88,888,666,665,555,500,000 "},
		{"A22", numFmt38, -8.8888666665555493e+19, "(88,888,666,665,555,500,000)"},
		{"A23", numFmt38, 8.8888666665555487, "9 "},
		{"A24", numFmt38, -8.8888666665555487, "(9)"},
		{"A25", numFmt39, 8.8888666665555493e+19, "88,888,666,665,555,500,000.00 "},
		{"A26", numFmt39, -8.8888666665555493e+19, "(88,888,666,665,555,500,000.00)"},
		{"A27", numFmt39, 8.8888666665555487, "8.89 "},
		{"A28", numFmt39, -8.8888666665555487, "(8.89)"},
		{"A29", numFmt40, 8.8888666665555493e+19, "88,888,666,665,555,500,000.00 "},
		{"A30", numFmt40, -8.8888666665555493e+19, "(88,888,666,665,555,500,000.00)"},
		{"A31", numFmt40, 8.8888666665555487, "8.89 "},
		{"A32", numFmt40, -8.8888666665555487, "(8.89)"},
		{"A33", numFmt21, 44729.999988368058, "23:59:59"},
		{"A34", numFmt21, 44944.375005787035, "09:00:00"},
		{"A35", numFmt21, 44944.375005798611, "09:00:01"},
	} {
		cell, styleID, value, expected := cases[0].(string), cases[1].(int), cases[2], cases[3].(string)
		assert.NoError(t, f.SetCellStyle("Sheet1", cell, cell, styleID))
		assert.NoError(t, f.SetCellValue("Sheet1", cell, value))
		result, err := f.GetCellValue("Sheet1", cell)
		assert.NoError(t, err)
		assert.Equal(t, expected, result, cell)
	}
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestNumberFormats.xlsx")))

	f = NewFile(Options{ShortDatePattern: "yyyy/m/d"})
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 43543.503206018519))
	numFmt14, err := f.NewStyle(&Style{NumFmt: 14})
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "A1", "A1", numFmt14))
	result, err := f.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "2019/3/19", result, "A1")
}

func BenchmarkRows(b *testing.B) {
	f, _ := OpenFile(filepath.Join("test", "Book1.xlsx"))
	for i := 0; i < b.N; i++ {
		rows, _ := f.Rows("Sheet2")
		for rows.Next() {
			row, _ := rows.Columns()
			for i := range row {
				if i >= 0 {
					continue
				}
			}
		}
		if err := rows.Close(); err != nil {
			b.Error(err)
		}
	}
	if err := f.Close(); err != nil {
		b.Error(err)
	}
}

// trimSliceSpace trim continually blank element in the tail of slice.
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
