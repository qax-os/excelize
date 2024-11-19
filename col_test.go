package excelize

import (
	"fmt"
	"path/filepath"
	"sync"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

func TestCols(t *testing.T) {
	const sheet2 = "Sheet2"

	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	cols, err := f.Cols(sheet2)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	var collectedRows [][]string
	for cols.Next() {
		rows, err := cols.Rows()
		assert.NoError(t, err)
		collectedRows = append(collectedRows, trimSliceSpace(rows))
	}
	if !assert.NoError(t, cols.Error()) {
		t.FailNow()
	}

	returnedColumns, err := f.GetCols(sheet2)
	assert.NoError(t, err)
	for i := range returnedColumns {
		returnedColumns[i] = trimSliceSpace(returnedColumns[i])
	}
	if !assert.Equal(t, collectedRows, returnedColumns) {
		t.FailNow()
	}
	assert.NoError(t, f.Close())

	f = NewFile()
	cells := []string{"C2", "C3", "C4"}
	for _, cell := range cells {
		assert.NoError(t, f.SetCellValue("Sheet1", cell, 1))
	}
	_, err = f.Rows("Sheet1")
	assert.NoError(t, err)

	f.Sheet.Store("xl/worksheets/sheet1.xml", &xlsxWorksheet{
		Dimension: &xlsxDimension{
			Ref: "C2:C4",
		},
	})
	_, err = f.Rows("Sheet1")
	assert.NoError(t, err)

	// Test columns iterator with invalid sheet name
	_, err = f.Cols("Sheet:1")
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())
	// Test get columns cells with invalid sheet name
	_, err = f.GetCols("Sheet:1")
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())
	// Test columns iterator with unsupported charset shared strings table
	f.SharedStrings = nil
	f.Pkg.Store(defaultXMLPathSharedStrings, MacintoshCyrillicCharset)
	cols, err = f.Cols("Sheet1")
	assert.NoError(t, err)
	cols.Next()
	_, err = cols.Rows()
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")

	f = NewFile()
	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(`<worksheet><sheetData><row r="A"><c r="2" t="inlineStr"><is><t>B</t></is></c></row></sheetData></worksheet>`))
	f.checked = sync.Map{}
	_, err = f.Cols("Sheet1")
	assert.EqualError(t, err, `strconv.Atoi: parsing "A": invalid syntax`)

	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(`<worksheet><sheetData><row r="2"><c r="A" t="inlineStr"><is><t>B</t></is></c></row></sheetData></worksheet>`))
	_, err = f.Cols("Sheet1")
	assert.EqualError(t, err, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
}

func TestColumnsIterator(t *testing.T) {
	sheetName, colCount, expectedNumCol := "Sheet2", 0, 9
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	require.NoError(t, err)

	cols, err := f.Cols(sheetName)
	require.NoError(t, err)

	for cols.Next() {
		colCount++
		require.True(t, colCount <= expectedNumCol, "colCount is greater than expected")
	}
	assert.Equal(t, expectedNumCol, colCount)
	assert.NoError(t, f.Close())

	f, sheetName, colCount, expectedNumCol = NewFile(), "Sheet1", 0, 4
	cells := []string{"C2", "C3", "C4", "D2", "D3", "D4"}
	for _, cell := range cells {
		assert.NoError(t, f.SetCellValue(sheetName, cell, 1))
	}
	cols, err = f.Cols(sheetName)
	require.NoError(t, err)

	for cols.Next() {
		colCount++
		require.True(t, colCount <= 4, "colCount is greater than expected")
	}
	assert.Equal(t, expectedNumCol, colCount)
}

func TestColsError(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	_, err = f.Cols("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	assert.NoError(t, f.Close())
}

func TestGetColsError(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	_, err = f.GetCols("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	assert.NoError(t, f.Close())

	f = NewFile()
	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(fmt.Sprintf(`<worksheet xmlns="%s"><sheetData><row r="A"><c r="2" t="inlineStr"><is><t>B</t></is></c></row></sheetData></worksheet>`, NameSpaceSpreadSheet.Value)))
	f.checked = sync.Map{}
	_, err = f.GetCols("Sheet1")
	assert.EqualError(t, err, `strconv.Atoi: parsing "A": invalid syntax`)

	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(fmt.Sprintf(`<worksheet xmlns="%s"><sheetData><row r="2"><c r="A" t="inlineStr"><is><t>B</t></is></c></row></sheetData></worksheet>`, NameSpaceSpreadSheet.Value)))
	_, err = f.GetCols("Sheet1")
	assert.EqualError(t, err, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())

	f = NewFile()
	cols, err := f.Cols("Sheet1")
	assert.NoError(t, err)
	cols.totalRows = 2
	cols.totalCols = 2
	cols.curCol = 1
	cols.sheetXML = []byte(fmt.Sprintf(`<worksheet xmlns="%s"><sheetData><row r="1"><c r="A" t="inlineStr"><is><t>A</t></is></c></row></sheetData></worksheet>`, NameSpaceSpreadSheet.Value))
	_, err = cols.Rows()
	assert.EqualError(t, err, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())

	f.Pkg.Store("xl/worksheets/sheet1.xml", nil)
	f.Sheet.Store("xl/worksheets/sheet1.xml", nil)
	_, err = f.Cols("Sheet1")
	assert.NoError(t, err)
}

func TestColsRows(t *testing.T) {
	f := NewFile()

	_, err := f.Cols("Sheet1")
	assert.NoError(t, err)

	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 1))
	f.Sheet.Store("xl/worksheets/sheet1.xml", &xlsxWorksheet{
		Dimension: &xlsxDimension{
			Ref: "A1:A1",
		},
	})

	f = NewFile()
	f.Pkg.Store("xl/worksheets/sheet1.xml", nil)
	_, err = f.Cols("Sheet1")
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	f = NewFile()
	cols, err := f.Cols("Sheet1")
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	_, err = cols.Rows()
	assert.NoError(t, err)
	cols.stashCol, cols.curCol = 0, 1
	// Test if token is nil
	cols.sheetXML = nil
	_, err = cols.Rows()
	assert.NoError(t, err)
}

func TestColumnVisibility(t *testing.T) {
	t.Run("TestBook1", func(t *testing.T) {
		f, err := prepareTestBook1()
		assert.NoError(t, err)

		// Hide/display a column with SetColVisible
		assert.NoError(t, f.SetColVisible("Sheet1", "F", false))
		assert.NoError(t, f.SetColVisible("Sheet1", "F", true))
		visible, err := f.GetColVisible("Sheet1", "F")
		assert.Equal(t, true, visible)
		assert.NoError(t, err)

		// Test hiding a few columns SetColVisible(...false)...
		assert.NoError(t, f.SetColVisible("Sheet1", "F:V", false))
		visible, err = f.GetColVisible("Sheet1", "F")
		assert.Equal(t, false, visible)
		assert.NoError(t, err)
		visible, err = f.GetColVisible("Sheet1", "U")
		assert.Equal(t, false, visible)
		assert.NoError(t, err)
		visible, err = f.GetColVisible("Sheet1", "V")
		assert.Equal(t, false, visible)
		assert.NoError(t, err)
		// ...and displaying them back SetColVisible(...true)
		assert.NoError(t, f.SetColVisible("Sheet1", "V:F", true))
		visible, err = f.GetColVisible("Sheet1", "F")
		assert.Equal(t, true, visible)
		assert.NoError(t, err)
		visible, err = f.GetColVisible("Sheet1", "U")
		assert.Equal(t, true, visible)
		assert.NoError(t, err)
		visible, err = f.GetColVisible("Sheet1", "G")
		assert.Equal(t, true, visible)
		assert.NoError(t, err)

		// Test get column visible on not exists worksheet
		_, err = f.GetColVisible("SheetN", "F")
		assert.EqualError(t, err, "sheet SheetN does not exist")
		// Test get column visible with invalid sheet name
		_, err = f.GetColVisible("Sheet:1", "F")
		assert.EqualError(t, err, ErrSheetNameInvalid.Error())
		// Test get column visible with illegal cell reference
		_, err = f.GetColVisible("Sheet1", "*")
		assert.EqualError(t, err, newInvalidColumnNameError("*").Error())
		assert.EqualError(t, f.SetColVisible("Sheet1", "*", false), newInvalidColumnNameError("*").Error())
		// Test set column visible with invalid sheet name
		assert.EqualError(t, f.SetColVisible("Sheet:1", "A", false), ErrSheetNameInvalid.Error())

		_, err = f.NewSheet("Sheet3")
		assert.NoError(t, err)
		assert.NoError(t, f.SetColVisible("Sheet3", "E", false))
		assert.EqualError(t, f.SetColVisible("Sheet1", "A:-1", true), newInvalidColumnNameError("-1").Error())
		assert.EqualError(t, f.SetColVisible("SheetN", "E", false), "sheet SheetN does not exist")
		assert.NoError(t, f.SaveAs(filepath.Join("test", "TestColumnVisibility.xlsx")))
	})

	t.Run("TestBook3", func(t *testing.T) {
		f, err := prepareTestBook3()
		assert.NoError(t, err)
		visible, err := f.GetColVisible("Sheet1", "B")
		assert.Equal(t, true, visible)
		assert.NoError(t, err)
	})
}

func TestOutlineLevel(t *testing.T) {
	f := NewFile()
	level, err := f.GetColOutlineLevel("Sheet1", "D")
	assert.Equal(t, uint8(0), level)
	assert.NoError(t, err)

	_, err = f.NewSheet("Sheet2")
	assert.NoError(t, err)
	assert.NoError(t, f.SetColOutlineLevel("Sheet1", "D", 4))

	level, err = f.GetColOutlineLevel("Sheet1", "D")
	assert.Equal(t, uint8(4), level)
	assert.NoError(t, err)

	level, err = f.GetColOutlineLevel("SheetN", "A")
	assert.Equal(t, uint8(0), level)
	assert.EqualError(t, err, "sheet SheetN does not exist")

	// Test column outline level with invalid sheet name
	_, err = f.GetColOutlineLevel("Sheet:1", "A")
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())

	assert.NoError(t, f.SetColWidth("Sheet2", "A", "D", 13))
	assert.EqualError(t, f.SetColWidth("Sheet2", "A", "D", MaxColumnWidth+1), ErrColumnWidth.Error())
	// Test set column width with invalid sheet name
	assert.EqualError(t, f.SetColWidth("Sheet:1", "A", "D", 13), ErrSheetNameInvalid.Error())

	assert.NoError(t, f.SetColOutlineLevel("Sheet2", "B", 2))
	assert.NoError(t, f.SetRowOutlineLevel("Sheet1", 2, 7))
	assert.EqualError(t, f.SetColOutlineLevel("Sheet1", "D", 8), ErrOutlineLevel.Error())
	assert.EqualError(t, f.SetRowOutlineLevel("Sheet1", 2, 8), ErrOutlineLevel.Error())
	// Test set row outline level on not exists worksheet
	assert.EqualError(t, f.SetRowOutlineLevel("SheetN", 1, 4), "sheet SheetN does not exist")
	// Test set row outline level with invalid sheet name
	assert.EqualError(t, f.SetRowOutlineLevel("Sheet:1", 1, 4), ErrSheetNameInvalid.Error())
	// Test get row outline level on not exists worksheet
	_, err = f.GetRowOutlineLevel("SheetN", 1)
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Test get row outline level with invalid sheet name
	_, err = f.GetRowOutlineLevel("Sheet:1", 1)
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())
	// Test set and get column outline level with illegal cell reference
	assert.EqualError(t, f.SetColOutlineLevel("Sheet1", "*", 1), newInvalidColumnNameError("*").Error())
	_, err = f.GetColOutlineLevel("Sheet1", "*")
	assert.EqualError(t, err, newInvalidColumnNameError("*").Error())

	// Test set column outline level on not exists worksheet
	assert.EqualError(t, f.SetColOutlineLevel("SheetN", "E", 2), "sheet SheetN does not exist")

	assert.EqualError(t, f.SetRowOutlineLevel("Sheet1", 0, 1), newInvalidRowNumberError(0).Error())
	level, err = f.GetRowOutlineLevel("Sheet1", 2)
	assert.NoError(t, err)
	assert.Equal(t, uint8(7), level)

	_, err = f.GetRowOutlineLevel("Sheet1", 0)
	assert.EqualError(t, err, newInvalidRowNumberError(0).Error())

	level, err = f.GetRowOutlineLevel("Sheet1", 10)
	assert.NoError(t, err)
	assert.Equal(t, uint8(0), level)

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestOutlineLevel.xlsx")))

	f, err = OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)
	assert.NoError(t, f.SetColOutlineLevel("Sheet2", "B", 2))
	assert.NoError(t, f.Close())
}

func TestSetColStyle(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellValue("Sheet1", "B2", "Hello"))

	styleID, err := f.NewStyle(&Style{Fill: Fill{Type: "pattern", Color: []string{"94D3A2"}, Pattern: 1}})
	assert.NoError(t, err)
	// Test set column style on not exists worksheet
	assert.EqualError(t, f.SetColStyle("SheetN", "E", styleID), "sheet SheetN does not exist")
	// Test set column style with illegal column name
	assert.EqualError(t, f.SetColStyle("Sheet1", "*", styleID), newInvalidColumnNameError("*").Error())
	assert.EqualError(t, f.SetColStyle("Sheet1", "A:*", styleID), newInvalidColumnNameError("*").Error())
	// Test set column style with invalid style ID
	assert.EqualError(t, f.SetColStyle("Sheet1", "B", -1), newInvalidStyleID(-1).Error())
	// Test set column style with not exists style ID
	assert.EqualError(t, f.SetColStyle("Sheet1", "B", 10), newInvalidStyleID(10).Error())
	// Test set column style with invalid sheet name
	assert.EqualError(t, f.SetColStyle("Sheet:1", "A", 0), ErrSheetNameInvalid.Error())

	assert.NoError(t, f.SetColStyle("Sheet1", "B", styleID))
	style, err := f.GetColStyle("Sheet1", "B")
	assert.NoError(t, err)
	assert.Equal(t, styleID, style)

	// Test set column style with already exists column with style
	assert.NoError(t, f.SetColStyle("Sheet1", "B", styleID))
	assert.NoError(t, f.SetColStyle("Sheet1", "D:C", styleID))
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).SheetData.Row[1].C[2].S = 0
	cellStyleID, err := f.GetCellStyle("Sheet1", "C2")
	assert.NoError(t, err)
	assert.Equal(t, styleID, cellStyleID)
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetColStyle.xlsx")))
	// Test set column style with unsupported charset style sheet
	f.Styles = nil
	f.Pkg.Store(defaultXMLPathStyles, MacintoshCyrillicCharset)
	assert.EqualError(t, f.SetColStyle("Sheet1", "C:F", styleID), "XML syntax error on line 1: invalid UTF-8")

	// Test set column style with worksheet properties columns default width settings
	f = NewFile()
	assert.NoError(t, f.SetSheetProps("Sheet1", &SheetPropsOptions{DefaultColWidth: float64Ptr(20)}))
	style, err = f.NewStyle(&Style{Alignment: &Alignment{Vertical: "center"}})
	assert.NoError(t, err)
	assert.NoError(t, f.SetColStyle("Sheet1", "A:Z", style))
	width, err := f.GetColWidth("Sheet1", "B")
	assert.NoError(t, err)
	assert.Equal(t, 20.0, width)
}

func TestColWidth(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetColWidth("Sheet1", "B", "A", 12))
	assert.NoError(t, f.SetColWidth("Sheet1", "A", "B", 12))
	width, err := f.GetColWidth("Sheet1", "A")
	assert.Equal(t, float64(12), width)
	assert.NoError(t, err)
	width, err = f.GetColWidth("Sheet1", "C")
	assert.Equal(t, defaultColWidth, width)
	assert.NoError(t, err)

	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).SheetFormatPr = &xlsxSheetFormatPr{DefaultColWidth: 10}
	ws.(*xlsxWorksheet).Cols = nil
	width, err = f.GetColWidth("Sheet1", "A")
	assert.NoError(t, err)
	assert.Equal(t, 10.0, width)
	assert.Equal(t, 76, f.getColWidth("Sheet1", 1))

	// Test set and get column width with illegal cell reference
	width, err = f.GetColWidth("Sheet1", "*")
	assert.Equal(t, defaultColWidth, width)
	assert.EqualError(t, err, newInvalidColumnNameError("*").Error())
	assert.EqualError(t, f.SetColWidth("Sheet1", "*", "B", 1), newInvalidColumnNameError("*").Error())
	assert.EqualError(t, f.SetColWidth("Sheet1", "A", "*", 1), newInvalidColumnNameError("*").Error())

	// Test set column width on not exists worksheet
	assert.EqualError(t, f.SetColWidth("SheetN", "B", "A", 12), "sheet SheetN does not exist")
	// Test get column width on not exists worksheet
	_, err = f.GetColWidth("SheetN", "A")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Test get column width invalid sheet name
	_, err = f.GetColWidth("Sheet:1", "A")
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestColWidth.xlsx")))
	convertRowHeightToPixels(0)
}

func TestGetColStyle(t *testing.T) {
	f := NewFile()
	styleID, err := f.GetColStyle("Sheet1", "A")
	assert.NoError(t, err)
	assert.Equal(t, styleID, 0)

	// Test get column style on not exists worksheet
	_, err = f.GetColStyle("SheetN", "A")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Test get column style with illegal column name
	_, err = f.GetColStyle("Sheet1", "*")
	assert.EqualError(t, err, newInvalidColumnNameError("*").Error())
	// Test get column style with invalid sheet name
	_, err = f.GetColStyle("Sheet:1", "A")
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())
}

func TestInsertCols(t *testing.T) {
	f := NewFile()
	sheet1 := f.GetSheetName(0)

	assert.NoError(t, fillCells(f, sheet1, 10, 10))

	assert.NoError(t, f.SetCellHyperLink(sheet1, "A5", "https://github.com/xuri/excelize", "External"))
	assert.NoError(t, f.MergeCell(sheet1, "A1", "C3"))

	assert.NoError(t, f.AutoFilter(sheet1, "A2:B2", []AutoFilterOptions{{Column: "B", Expression: "x != blanks"}}))
	assert.NoError(t, f.InsertCols(sheet1, "A", 1))

	// Test insert column with illegal cell reference
	assert.EqualError(t, f.InsertCols(sheet1, "*", 1), newInvalidColumnNameError("*").Error())
	// Test insert column with invalid sheet name
	assert.EqualError(t, f.InsertCols("Sheet:1", "A", 1), ErrSheetNameInvalid.Error())
	assert.EqualError(t, f.InsertCols(sheet1, "A", 0), ErrColumnNumber.Error())
	assert.EqualError(t, f.InsertCols(sheet1, "A", MaxColumns), ErrColumnNumber.Error())
	assert.EqualError(t, f.InsertCols(sheet1, "A", MaxColumns-10), ErrColumnNumber.Error())

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestInsertCols.xlsx")))
}

func TestRemoveCol(t *testing.T) {
	f := NewFile()
	sheet1 := f.GetSheetName(0)

	assert.NoError(t, fillCells(f, sheet1, 10, 15))

	assert.NoError(t, f.SetCellHyperLink(sheet1, "A5", "https://github.com/xuri/excelize", "External"))
	assert.NoError(t, f.SetCellHyperLink(sheet1, "C5", "https://github.com", "External"))

	assert.NoError(t, f.MergeCell(sheet1, "A1", "B1"))
	assert.NoError(t, f.MergeCell(sheet1, "A2", "B2"))

	assert.NoError(t, f.RemoveCol(sheet1, "A"))
	assert.NoError(t, f.RemoveCol(sheet1, "A"))

	// Test remove column with illegal cell reference
	assert.EqualError(t, f.RemoveCol("Sheet1", "*"), newInvalidColumnNameError("*").Error())
	// Test remove column on not exists worksheet
	assert.EqualError(t, f.RemoveCol("SheetN", "B"), "sheet SheetN does not exist")
	// Test remove column  with invalid sheet name
	assert.EqualError(t, f.RemoveCol("Sheet:1", "A"), ErrSheetNameInvalid.Error())

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestRemoveCol.xlsx")))
}

func TestConvertColWidthToPixels(t *testing.T) {
	assert.Equal(t, -11.0, convertColWidthToPixels(-1))
}
