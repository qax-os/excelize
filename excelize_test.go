package excelize

import (
	"fmt"
	"image/color"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"io/ioutil"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
)

func TestOpenFile(t *testing.T) {
	// Test update a XLSX file.
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Test get all the rows in a not exists worksheet.
	f.GetRows("Sheet4")
	// Test get all the rows in a worksheet.
	rows, err := f.GetRows("Sheet2")
	assert.NoError(t, err)
	for _, row := range rows {
		for _, cell := range row {
			t.Log(cell, "\t")
		}
		t.Log("\r\n")
	}
	f.UpdateLinkedValue()

	f.SetCellDefault("Sheet2", "A1", strconv.FormatFloat(float64(100.1588), 'f', -1, 32))
	f.SetCellDefault("Sheet2", "A1", strconv.FormatFloat(float64(-100.1588), 'f', -1, 64))

	// Test set cell value with illegal row number.
	assert.EqualError(t, f.SetCellDefault("Sheet2", "A", strconv.FormatFloat(float64(-100.1588), 'f', -1, 64)),
		`cannot convert cell "A" to coordinates: invalid cell name "A"`)

	f.SetCellInt("Sheet2", "A1", 100)

	// Test set cell integer value with illegal row number.
	assert.EqualError(t, f.SetCellInt("Sheet2", "A", 100), `cannot convert cell "A" to coordinates: invalid cell name "A"`)

	f.SetCellStr("Sheet2", "C11", "Knowns")
	// Test max characters in a cell.
	f.SetCellStr("Sheet2", "D11", strings.Repeat("c", 32769))
	f.NewSheet(":\\/?*[]Maximum 31 characters allowed in sheet title.")
	// Test set worksheet name with illegal name.
	f.SetSheetName("Maximum 31 characters allowed i", "[Rename]:\\/?* Maximum 31 characters allowed in sheet title.")
	f.SetCellInt("Sheet3", "A23", 10)
	f.SetCellStr("Sheet3", "b230", "10")
	assert.EqualError(t, f.SetCellStr("Sheet10", "b230", "10"), "sheet Sheet10 is not exist")

	// Test set cell string value with illegal row number.
	assert.EqualError(t, f.SetCellStr("Sheet1", "A", "10"), `cannot convert cell "A" to coordinates: invalid cell name "A"`)

	f.SetActiveSheet(2)
	// Test get cell formula with given rows number.
	_, err = f.GetCellFormula("Sheet1", "B19")
	assert.NoError(t, err)
	// Test get cell formula with illegal worksheet name.
	_, err = f.GetCellFormula("Sheet2", "B20")
	assert.NoError(t, err)
	_, err = f.GetCellFormula("Sheet1", "B20")
	assert.NoError(t, err)

	// Test get cell formula with illegal rows number.
	_, err = f.GetCellFormula("Sheet1", "B")
	assert.EqualError(t, err, `cannot convert cell "B" to coordinates: invalid cell name "B"`)
	// Test get shared cell formula
	f.GetCellFormula("Sheet2", "H11")
	f.GetCellFormula("Sheet2", "I11")
	getSharedForumula(&xlsxWorksheet{}, "")

	// Test read cell value with given illegal rows number.
	_, err = f.GetCellValue("Sheet2", "a-1")
	assert.EqualError(t, err, `cannot convert cell "A-1" to coordinates: invalid cell name "A-1"`)
	_, err = f.GetCellValue("Sheet2", "A")
	assert.EqualError(t, err, `cannot convert cell "A" to coordinates: invalid cell name "A"`)

	// Test read cell value with given lowercase column number.
	f.GetCellValue("Sheet2", "a5")
	f.GetCellValue("Sheet2", "C11")
	f.GetCellValue("Sheet2", "D11")
	f.GetCellValue("Sheet2", "D12")
	// Test SetCellValue function.
	assert.NoError(t, f.SetCellValue("Sheet2", "F1", " Hello"))
	assert.NoError(t, f.SetCellValue("Sheet2", "G1", []byte("World")))
	assert.NoError(t, f.SetCellValue("Sheet2", "F2", 42))
	assert.NoError(t, f.SetCellValue("Sheet2", "F3", int8(1<<8/2-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F4", int16(1<<16/2-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F5", int32(1<<32/2-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F6", int64(1<<32/2-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F7", float32(42.65418)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F8", float64(-42.65418)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F9", float32(42)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F10", float64(42)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F11", uint(1<<32-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F12", uint8(1<<8-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F13", uint16(1<<16-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F14", uint32(1<<32-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F15", uint64(1<<32-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F16", true))
	assert.NoError(t, f.SetCellValue("Sheet2", "F17", complex64(5+10i)))

	// Test on not exists worksheet.
	assert.EqualError(t, f.SetCellDefault("SheetN", "A1", ""), "sheet SheetN is not exist")
	assert.EqualError(t, f.SetCellFloat("SheetN", "A1", 42.65418, 2, 32), "sheet SheetN is not exist")
	assert.EqualError(t, f.SetCellBool("SheetN", "A1", true), "sheet SheetN is not exist")
	assert.EqualError(t, f.SetCellFormula("SheetN", "A1", ""), "sheet SheetN is not exist")
	assert.EqualError(t, f.SetCellHyperLink("SheetN", "A1", "Sheet1!A40", "Location"), "sheet SheetN is not exist")

	// Test boolean write
	booltest := []struct {
		value    bool
		expected string
	}{
		{false, "0"},
		{true, "1"},
	}
	for _, test := range booltest {
		f.SetCellValue("Sheet2", "F16", test.value)
		val, err := f.GetCellValue("Sheet2", "F16")
		assert.NoError(t, err)
		assert.Equal(t, test.expected, val)
	}

	f.SetCellValue("Sheet2", "G2", nil)

	assert.EqualError(t, f.SetCellValue("Sheet2", "G4", time.Now()), "only UTC time expected")

	f.SetCellValue("Sheet2", "G4", time.Now().UTC())
	// 02:46:40
	f.SetCellValue("Sheet2", "G5", time.Duration(1e13))
	// Test completion column.
	f.SetCellValue("Sheet2", "M2", nil)
	// Test read cell value with given axis large than exists row.
	f.GetCellValue("Sheet2", "E231")
	// Test get active worksheet of XLSX and get worksheet name of XLSX by given worksheet index.
	f.GetSheetName(f.GetActiveSheetIndex())
	// Test get worksheet index of XLSX by given worksheet name.
	f.GetSheetIndex("Sheet1")
	// Test get worksheet name of XLSX by given invalid worksheet index.
	f.GetSheetName(4)
	// Test get worksheet map of f.
	f.GetSheetMap()
	for i := 1; i <= 300; i++ {
		f.SetCellStr("Sheet3", "c"+strconv.Itoa(i), strconv.Itoa(i))
	}
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestOpenFile.xlsx")))
}

func TestSaveFile(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSaveFile.xlsx")))
	f, err = OpenFile(filepath.Join("test", "TestSaveFile.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.NoError(t, f.Save())
}

func TestSaveAsWrongPath(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if assert.NoError(t, err) {
		// Test write file to not exist directory.
		err = f.SaveAs("")
		if assert.Error(t, err) {
			assert.True(t, os.IsNotExist(err), "Error: %v: Expected os.IsNotExists(err) == true", err)
		}
	}
}

func TestBrokenFile(t *testing.T) {
	// Test write file with broken file struct.
	f := File{}

	t.Run("SaveWithoutName", func(t *testing.T) {
		assert.EqualError(t, f.Save(), "no path defined for file, consider File.WriteTo or File.Write")
	})

	t.Run("SaveAsEmptyStruct", func(t *testing.T) {
		// Test write file with broken file struct with given path.
		assert.NoError(t, f.SaveAs(filepath.Join("test", "BrokenFile.SaveAsEmptyStruct.xlsx")))
	})

	t.Run("OpenBadWorkbook", func(t *testing.T) {
		// Test set active sheet without BookViews and Sheets maps in xl/workbook.xml.
		f3, err := OpenFile(filepath.Join("test", "BadWorkbook.xlsx"))
		f3.GetActiveSheetIndex()
		f3.SetActiveSheet(2)
		assert.NoError(t, err)
	})

	t.Run("OpenNotExistsFile", func(t *testing.T) {
		// Test open a XLSX file with given illegal path.
		_, err := OpenFile(filepath.Join("test", "NotExistsFile.xlsx"))
		if assert.Error(t, err) {
			assert.True(t, os.IsNotExist(err), "Expected os.IsNotExists(err) == true")
		}
	})
}

func TestNewFile(t *testing.T) {
	// Test create a XLSX file.
	f := NewFile()
	f.NewSheet("Sheet1")
	f.NewSheet("XLSXSheet2")
	f.NewSheet("XLSXSheet3")
	f.SetCellInt("XLSXSheet2", "A23", 56)
	f.SetCellStr("Sheet1", "B20", "42")
	f.SetActiveSheet(0)

	// Test add picture to sheet with scaling and positioning.
	err := f.AddPicture("Sheet1", "H2", filepath.Join("test", "images", "excel.gif"),
		`{"x_scale": 0.5, "y_scale": 0.5, "positioning": "absolute"}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Test add picture to worksheet without formatset.
	err = f.AddPicture("Sheet1", "C2", filepath.Join("test", "images", "excel.png"), "")
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Test add picture to worksheet with invalid formatset.
	err = f.AddPicture("Sheet1", "C2", filepath.Join("test", "images", "excel.png"), `{`)
	if !assert.Error(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestNewFile.xlsx")))
}

func TestAddDrawingVML(t *testing.T) {
	// Test addDrawingVML with illegal cell coordinates.
	f := NewFile()
	assert.EqualError(t, f.addDrawingVML(0, "", "*", 0, 0), `cannot convert cell "*" to coordinates: invalid cell name "*"`)
}

func TestSetCellHyperLink(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if err != nil {
		t.Log(err)
	}
	// Test set cell hyperlink in a work sheet already have hyperlinks.
	assert.NoError(t, f.SetCellHyperLink("Sheet1", "B19", "https://github.com/360EntSecGroup-Skylar/excelize", "External"))
	// Test add first hyperlink in a work sheet.
	assert.NoError(t, f.SetCellHyperLink("Sheet2", "C1", "https://github.com/360EntSecGroup-Skylar/excelize", "External"))
	// Test add Location hyperlink in a work sheet.
	assert.NoError(t, f.SetCellHyperLink("Sheet2", "D6", "Sheet1!D8", "Location"))

	assert.EqualError(t, f.SetCellHyperLink("Sheet2", "C3", "Sheet1!D8", ""), `invalid link type ""`)

	assert.EqualError(t, f.SetCellHyperLink("Sheet2", "", "Sheet1!D60", "Location"), `invalid cell name ""`)

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellHyperLink.xlsx")))

	file := NewFile()
	for row := 1; row <= 65530; row++ {
		cell, err := CoordinatesToCellName(1, row)
		assert.NoError(t, err)
		assert.NoError(t, file.SetCellHyperLink("Sheet1", cell, "https://github.com/360EntSecGroup-Skylar/excelize", "External"))
	}
	assert.EqualError(t, file.SetCellHyperLink("Sheet1", "A65531", "https://github.com/360EntSecGroup-Skylar/excelize", "External"), "over maximum limit hyperlinks in a worksheet")
}

func TestGetCellHyperLink(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	_, _, err = f.GetCellHyperLink("Sheet1", "")
	assert.EqualError(t, err, `invalid cell name ""`)

	link, target, err := f.GetCellHyperLink("Sheet1", "A22")
	assert.NoError(t, err)
	t.Log(link, target)
	link, target, err = f.GetCellHyperLink("Sheet2", "D6")
	assert.NoError(t, err)
	t.Log(link, target)
	link, target, err = f.GetCellHyperLink("Sheet3", "H3")
	assert.EqualError(t, err, "sheet Sheet3 is not exist")
	t.Log(link, target)
}

func TestSetCellFormula(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	f.SetCellFormula("Sheet1", "B19", "SUM(Sheet2!D2,Sheet2!D11)")
	f.SetCellFormula("Sheet1", "C19", "SUM(Sheet2!D2,Sheet2!D9)")

	// Test set cell formula with illegal rows number.
	assert.EqualError(t, f.SetCellFormula("Sheet1", "C", "SUM(Sheet2!D2,Sheet2!D9)"), `cannot convert cell "C" to coordinates: invalid cell name "C"`)

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellFormula1.xlsx")))

	f, err = OpenFile(filepath.Join("test", "CalcChain.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	// Test remove cell formula.
	f.SetCellFormula("Sheet1", "A1", "")
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellFormula2.xlsx")))
	// Test remove all cell formula.
	f.SetCellFormula("Sheet1", "B1", "")
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellFormula3.xlsx")))
}

func TestSetSheetBackground(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	err = f.SetSheetBackground("Sheet2", filepath.Join("test", "images", "background.jpg"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	err = f.SetSheetBackground("Sheet2", filepath.Join("test", "images", "background.jpg"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetSheetBackground.xlsx")))
}

func TestSetSheetBackgroundErrors(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	err = f.SetSheetBackground("Sheet2", filepath.Join("test", "not_exists", "not_exists.png"))
	if assert.Error(t, err) {
		assert.True(t, os.IsNotExist(err), "Expected os.IsNotExists(err) == true")
	}

	err = f.SetSheetBackground("Sheet2", filepath.Join("test", "Book1.xlsx"))
	assert.EqualError(t, err, "unsupported image extension")
}

func TestMergeCell(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

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

func TestSetCellStyleAlignment(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	var style int
	style, err = f.NewStyle(`{"alignment":{"horizontal":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":45,"vertical":"top","wrap_text":true}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SetCellStyle("Sheet1", "A22", "A22", style))

	// Test set cell style with given illegal rows number.
	assert.EqualError(t, f.SetCellStyle("Sheet1", "A", "A22", style), `cannot convert cell "A" to coordinates: invalid cell name "A"`)
	assert.EqualError(t, f.SetCellStyle("Sheet1", "A22", "A", style), `cannot convert cell "A" to coordinates: invalid cell name "A"`)

	// Test get cell style with given illegal rows number.
	index, err := f.GetCellStyle("Sheet1", "A")
	assert.Equal(t, 0, index)
	assert.EqualError(t, err, `cannot convert cell "A" to coordinates: invalid cell name "A"`)

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleAlignment.xlsx")))
}

func TestSetCellStyleBorder(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	var style int

	// Test set border on overlapping area with vertical variants shading styles gradient fill.
	style, err = f.NewStyle(`{"border":[{"type":"left","color":"0000FF","style":2},{"type":"top","color":"00FF00","style":12},{"type":"bottom","color":"FFFF00","style":5},{"type":"right","color":"FF0000","style":6},{"type":"diagonalDown","color":"A020F0","style":9},{"type":"diagonalUp","color":"A020F0","style":8}]}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.NoError(t, f.SetCellStyle("Sheet1", "J21", "L25", style))

	style, err = f.NewStyle(`{"border":[{"type":"left","color":"0000FF","style":2},{"type":"top","color":"00FF00","style":3},{"type":"bottom","color":"FFFF00","style":4},{"type":"right","color":"FF0000","style":5},{"type":"diagonalDown","color":"A020F0","style":6},{"type":"diagonalUp","color":"A020F0","style":7}],"fill":{"type":"gradient","color":["#FFFFFF","#E0EBF5"],"shading":1}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.NoError(t, f.SetCellStyle("Sheet1", "M28", "K24", style))

	style, err = f.NewStyle(`{"border":[{"type":"left","color":"0000FF","style":2},{"type":"top","color":"00FF00","style":3},{"type":"bottom","color":"FFFF00","style":4},{"type":"right","color":"FF0000","style":5},{"type":"diagonalDown","color":"A020F0","style":6},{"type":"diagonalUp","color":"A020F0","style":7}],"fill":{"type":"gradient","color":["#FFFFFF","#E0EBF5"],"shading":4}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.NoError(t, f.SetCellStyle("Sheet1", "M28", "K24", style))

	// Test set border and solid style pattern fill for a single cell.
	style, err = f.NewStyle(`{"border":[{"type":"left","color":"0000FF","style":8},{"type":"top","color":"00FF00","style":9},{"type":"bottom","color":"FFFF00","style":10},{"type":"right","color":"FF0000","style":11},{"type":"diagonalDown","color":"A020F0","style":12},{"type":"diagonalUp","color":"A020F0","style":13}],"fill":{"type":"pattern","color":["#E0EBF5"],"pattern":1}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SetCellStyle("Sheet1", "O22", "O22", style))

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleBorder.xlsx")))
}

func TestSetCellStyleBorderErrors(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Set border with invalid style parameter.
	_, err = f.NewStyle("")
	if !assert.EqualError(t, err, "unexpected end of JSON input") {
		t.FailNow()
	}

	// Set border with invalid style index number.
	_, err = f.NewStyle(`{"border":[{"type":"left","color":"0000FF","style":-1},{"type":"top","color":"00FF00","style":14},{"type":"bottom","color":"FFFF00","style":5},{"type":"right","color":"FF0000","style":6},{"type":"diagonalDown","color":"A020F0","style":9},{"type":"diagonalUp","color":"A020F0","style":8}]}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}
}

func TestSetCellStyleNumberFormat(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Test only set fill and number format for a cell.
	col := []string{"L", "M", "N", "O", "P"}
	data := []int{0, 1, 2, 3, 4, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49}
	value := []string{"37947.7500001", "-37947.7500001", "0.007", "2.1", "String"}
	for i, v := range value {
		for k, d := range data {
			c := col[i] + strconv.Itoa(k+1)
			var val float64
			val, err = strconv.ParseFloat(v, 64)
			if err != nil {
				f.SetCellValue("Sheet2", c, v)
			} else {
				f.SetCellValue("Sheet2", c, val)
			}
			style, err := f.NewStyle(`{"fill":{"type":"gradient","color":["#FFFFFF","#E0EBF5"],"shading":5},"number_format": ` + strconv.Itoa(d) + `}`)
			if !assert.NoError(t, err) {
				t.FailNow()
			}
			assert.NoError(t, f.SetCellStyle("Sheet2", c, c, style))
			t.Log(f.GetCellValue("Sheet2", c))
		}
	}
	var style int
	style, err = f.NewStyle(`{"number_format":-1}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.NoError(t, f.SetCellStyle("Sheet2", "L33", "L33", style))

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleNumberFormat.xlsx")))
}

func TestSetCellStyleCurrencyNumberFormat(t *testing.T) {
	t.Run("TestBook3", func(t *testing.T) {
		f, err := prepareTestBook3()
		if !assert.NoError(t, err) {
			t.FailNow()
		}

		f.SetCellValue("Sheet1", "A1", 56)
		f.SetCellValue("Sheet1", "A2", -32.3)
		var style int
		style, err = f.NewStyle(`{"number_format": 188, "decimal_places": -1}`)
		if !assert.NoError(t, err) {
			t.FailNow()
		}

		assert.NoError(t, f.SetCellStyle("Sheet1", "A1", "A1", style))
		style, err = f.NewStyle(`{"number_format": 188, "decimal_places": 31, "negred": true}`)
		if !assert.NoError(t, err) {
			t.FailNow()
		}

		assert.NoError(t, f.SetCellStyle("Sheet1", "A2", "A2", style))

		assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleCurrencyNumberFormat.TestBook3.xlsx")))
	})

	t.Run("TestBook4", func(t *testing.T) {
		f, err := prepareTestBook4()
		if !assert.NoError(t, err) {
			t.FailNow()
		}
		f.SetCellValue("Sheet1", "A1", 42920.5)
		f.SetCellValue("Sheet1", "A2", 42920.5)

		_, err = f.NewStyle(`{"number_format": 26, "lang": "zh-tw"}`)
		if !assert.NoError(t, err) {
			t.FailNow()
		}

		style, err := f.NewStyle(`{"number_format": 27}`)
		if !assert.NoError(t, err) {
			t.FailNow()
		}

		assert.NoError(t, f.SetCellStyle("Sheet1", "A1", "A1", style))
		style, err = f.NewStyle(`{"number_format": 31, "lang": "ko-kr"}`)
		if !assert.NoError(t, err) {
			t.FailNow()
		}

		assert.NoError(t, f.SetCellStyle("Sheet1", "A2", "A2", style))

		style, err = f.NewStyle(`{"number_format": 71, "lang": "th-th"}`)
		if !assert.NoError(t, err) {
			t.FailNow()
		}
		assert.NoError(t, f.SetCellStyle("Sheet1", "A2", "A2", style))

		assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleCurrencyNumberFormat.TestBook4.xlsx")))
	})
}

func TestSetCellStyleCustomNumberFormat(t *testing.T) {
	f := NewFile()
	f.SetCellValue("Sheet1", "A1", 42920.5)
	f.SetCellValue("Sheet1", "A2", 42920.5)
	style, err := f.NewStyle(`{"custom_number_format": "[$-380A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@"}`)
	if err != nil {
		t.Log(err)
	}
	assert.NoError(t, f.SetCellStyle("Sheet1", "A1", "A1", style))
	style, err = f.NewStyle(`{"custom_number_format": "[$-380A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@"}`)
	if err != nil {
		t.Log(err)
	}
	assert.NoError(t, f.SetCellStyle("Sheet1", "A2", "A2", style))

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleCustomNumberFormat.xlsx")))
}

func TestSetCellStyleFill(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	var style int
	// Test set fill for cell with invalid parameter.
	style, err = f.NewStyle(`{"fill":{"type":"gradient","color":["#FFFFFF","#E0EBF5"],"shading":6}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.NoError(t, f.SetCellStyle("Sheet1", "O23", "O23", style))

	style, err = f.NewStyle(`{"fill":{"type":"gradient","color":["#FFFFFF"],"shading":1}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.NoError(t, f.SetCellStyle("Sheet1", "O23", "O23", style))

	style, err = f.NewStyle(`{"fill":{"type":"pattern","color":[],"pattern":1}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.NoError(t, f.SetCellStyle("Sheet1", "O23", "O23", style))

	style, err = f.NewStyle(`{"fill":{"type":"pattern","color":["#E0EBF5"],"pattern":19}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.NoError(t, f.SetCellStyle("Sheet1", "O23", "O23", style))

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleFill.xlsx")))
}

func TestSetCellStyleFont(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	var style int
	style, err = f.NewStyle(`{"font":{"bold":true,"italic":true,"family":"Berlin Sans FB Demi","size":36,"color":"#777777","underline":"single"}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SetCellStyle("Sheet2", "A1", "A1", style))

	style, err = f.NewStyle(`{"font":{"italic":true,"underline":"double"}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SetCellStyle("Sheet2", "A2", "A2", style))

	style, err = f.NewStyle(`{"font":{"bold":true}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SetCellStyle("Sheet2", "A3", "A3", style))

	style, err = f.NewStyle(`{"font":{"bold":true,"family":"","size":0,"color":"","underline":""}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SetCellStyle("Sheet2", "A4", "A4", style))

	style, err = f.NewStyle(`{"font":{"color":"#777777"}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SetCellStyle("Sheet2", "A5", "A5", style))

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleFont.xlsx")))
}

func TestSetCellStyleProtection(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	var style int
	style, err = f.NewStyle(`{"protection":{"hidden":true, "locked":true}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SetCellStyle("Sheet2", "A6", "A6", style))
	err = f.SaveAs(filepath.Join("test", "TestSetCellStyleProtection.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
}

func TestSetDeleteSheet(t *testing.T) {
	t.Run("TestBook3", func(t *testing.T) {
		f, err := prepareTestBook3()
		if !assert.NoError(t, err) {
			t.FailNow()
		}

		f.DeleteSheet("XLSXSheet3")
		assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetDeleteSheet.TestBook3.xlsx")))
	})

	t.Run("TestBook4", func(t *testing.T) {
		f, err := prepareTestBook4()
		if !assert.NoError(t, err) {
			t.FailNow()
		}
		f.DeleteSheet("Sheet1")
		f.AddComment("Sheet1", "A1", "")
		f.AddComment("Sheet1", "A1", `{"author":"Excelize: ","text":"This is a comment."}`)
		assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetDeleteSheet.TestBook4.xlsx")))
	})
}

func TestSheetVisibility(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	f.SetSheetVisible("Sheet2", false)
	f.SetSheetVisible("Sheet1", false)
	f.SetSheetVisible("Sheet1", true)
	f.GetSheetVisible("Sheet1")

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSheetVisibility.xlsx")))
}

func TestCopySheet(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	idx := f.NewSheet("CopySheet")
	assert.NoError(t, f.CopySheet(1, idx))

	f.SetCellValue("CopySheet", "F1", "Hello")
	val, err := f.GetCellValue("Sheet1", "F1")
	assert.NoError(t, err)
	assert.NotEqual(t, "Hello", val)

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestCopySheet.xlsx")))
}

func TestCopySheetError(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.EqualError(t, f.copySheet(0, -1), "sheet  is not exist")
	if !assert.EqualError(t, f.CopySheet(0, -1), "invalid worksheet index") {
		t.FailNow()
	}

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestCopySheetError.xlsx")))
}

func TestAddTable(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	err = f.AddTable("Sheet1", "B26", "A21", `{}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	err = f.AddTable("Sheet2", "A2", "B5", `{"table_name":"table","table_style":"TableStyleMedium2", "show_first_column":true,"show_last_column":true,"show_row_stripes":false,"show_column_stripes":true}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	err = f.AddTable("Sheet2", "F1", "F1", `{"table_style":"TableStyleMedium8"}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Test add table with illegal formatset.
	assert.EqualError(t, f.AddTable("Sheet1", "B26", "A21", `{x}`), "invalid character 'x' looking for beginning of object key string")
	// Test add table with illegal cell coordinates.
	assert.EqualError(t, f.AddTable("Sheet1", "A", "B1", `{}`), `cannot convert cell "A" to coordinates: invalid cell name "A"`)
	assert.EqualError(t, f.AddTable("Sheet1", "A1", "B", `{}`), `cannot convert cell "B" to coordinates: invalid cell name "B"`)

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddTable.xlsx")))

	// Test addTable with illegal cell coordinates.
	f = NewFile()
	assert.EqualError(t, f.addTable("sheet1", "", 0, 0, 0, 0, 0, nil), "invalid cell coordinates [0, 0]")
	assert.EqualError(t, f.addTable("sheet1", "", 1, 1, 0, 0, 0, nil), "invalid cell coordinates [0, 0]")
}

func TestAddShape(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	f.AddShape("Sheet1", "A30", `{"type":"rect","paragraph":[{"text":"Rectangle","font":{"color":"CD5C5C"}},{"text":"Shape","font":{"bold":true,"color":"2980B9"}}]}`)
	f.AddShape("Sheet1", "B30", `{"type":"rect","paragraph":[{"text":"Rectangle"},{}]}`)
	f.AddShape("Sheet1", "C30", `{"type":"rect","paragraph":[]}`)
	f.AddShape("Sheet3", "H1", `{"type":"ellipseRibbon", "color":{"line":"#4286f4","fill":"#8eb9ff"}, "paragraph":[{"font":{"bold":true,"italic":true,"family":"Berlin Sans FB Demi","size":36,"color":"#777777","underline":"single"}}], "height": 90}`)
	f.AddShape("Sheet3", "H1", "")

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddShape.xlsx")))
}

func TestAddComments(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	s := strings.Repeat("c", 32768)
	assert.NoError(t, f.AddComment("Sheet1", "A30", `{"author":"`+s+`","text":"`+s+`"}`))
	assert.NoError(t, f.AddComment("Sheet2", "B7", `{"author":"Excelize: ","text":"This is a comment."}`))

	// Test add comment on not exists worksheet.
	assert.EqualError(t, f.AddComment("SheetN", "B7", `{"author":"Excelize: ","text":"This is a comment."}`), "sheet SheetN is not exist")

	if assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddComments.xlsx"))) {
		assert.Len(t, f.GetComments(), 2)
	}
}

func TestGetSheetComments(t *testing.T) {
	f := NewFile()
	assert.Equal(t, "", f.getSheetComments(0))
}

func TestAutoFilter(t *testing.T) {
	outFile := filepath.Join("test", "TestAutoFilter%d.xlsx")

	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	formats := []string{
		``,
		`{"column":"B","expression":"x != blanks"}`,
		`{"column":"B","expression":"x == blanks"}`,
		`{"column":"B","expression":"x != nonblanks"}`,
		`{"column":"B","expression":"x == nonblanks"}`,
		`{"column":"B","expression":"x <= 1 and x >= 2"}`,
		`{"column":"B","expression":"x == 1 or x == 2"}`,
		`{"column":"B","expression":"x == 1 or x == 2*"}`,
	}

	for i, format := range formats {
		t.Run(fmt.Sprintf("Expression%d", i+1), func(t *testing.T) {
			err = f.AutoFilter("Sheet1", "D4", "B1", format)
			assert.NoError(t, err)
			assert.NoError(t, f.SaveAs(fmt.Sprintf(outFile, i+1)))
		})
	}

	// testing AutoFilter with illegal cell coordinates.
	assert.EqualError(t, f.AutoFilter("Sheet1", "A", "B1", ""), `cannot convert cell "A" to coordinates: invalid cell name "A"`)
	assert.EqualError(t, f.AutoFilter("Sheet1", "A1", "B", ""), `cannot convert cell "B" to coordinates: invalid cell name "B"`)
}

func TestAutoFilterError(t *testing.T) {
	outFile := filepath.Join("test", "TestAutoFilterError%d.xlsx")

	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	formats := []string{
		`{"column":"B","expression":"x <= 1 and x >= blanks"}`,
		`{"column":"B","expression":"x -- y or x == *2*"}`,
		`{"column":"B","expression":"x != y or x ? *2"}`,
		`{"column":"B","expression":"x -- y o r x == *2"}`,
		`{"column":"B","expression":"x -- y"}`,
		`{"column":"A","expression":"x -- y"}`,
	}
	for i, format := range formats {
		t.Run(fmt.Sprintf("Expression%d", i+1), func(t *testing.T) {
			err = f.AutoFilter("Sheet3", "D4", "B1", format)
			if assert.Error(t, err) {
				assert.NoError(t, f.SaveAs(fmt.Sprintf(outFile, i+1)))
			}
		})
	}
}

func TestSetPane(t *testing.T) {
	f := NewFile()
	f.SetPanes("Sheet1", `{"freeze":false,"split":false}`)
	f.NewSheet("Panes 2")
	f.SetPanes("Panes 2", `{"freeze":true,"split":false,"x_split":1,"y_split":0,"top_left_cell":"B1","active_pane":"topRight","panes":[{"sqref":"K16","active_cell":"K16","pane":"topRight"}]}`)
	f.NewSheet("Panes 3")
	f.SetPanes("Panes 3", `{"freeze":false,"split":true,"x_split":3270,"y_split":1800,"top_left_cell":"N57","active_pane":"bottomLeft","panes":[{"sqref":"I36","active_cell":"I36"},{"sqref":"G33","active_cell":"G33","pane":"topRight"},{"sqref":"J60","active_cell":"J60","pane":"bottomLeft"},{"sqref":"O60","active_cell":"O60","pane":"bottomRight"}]}`)
	f.NewSheet("Panes 4")
	f.SetPanes("Panes 4", `{"freeze":true,"split":false,"x_split":0,"y_split":9,"top_left_cell":"A34","active_pane":"bottomLeft","panes":[{"sqref":"A11:XFD11","active_cell":"A11","pane":"bottomLeft"}]}`)
	f.SetPanes("Panes 4", "")

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetPane.xlsx")))
}

func TestConditionalFormat(t *testing.T) {
	f := NewFile()
	sheet1 := f.GetSheetName(1)

	fillCells(f, sheet1, 10, 15)

	var format1, format2, format3 int
	var err error
	// Rose format for bad conditional.
	format1, err = f.NewConditionalStyle(`{"font":{"color":"#9A0511"},"fill":{"type":"pattern","color":["#FEC7CE"],"pattern":1}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Light yellow format for neutral conditional.
	format2, err = f.NewConditionalStyle(`{"fill":{"type":"pattern","color":["#FEEAA0"],"pattern":1}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Light green format for good conditional.
	format3, err = f.NewConditionalStyle(`{"font":{"color":"#09600B"},"fill":{"type":"pattern","color":["#C7EECF"],"pattern":1}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Color scales: 2 color.
	f.SetConditionalFormat(sheet1, "A1:A10", `[{"type":"2_color_scale","criteria":"=","min_type":"min","max_type":"max","min_color":"#F8696B","max_color":"#63BE7B"}]`)
	// Color scales: 3 color.
	f.SetConditionalFormat(sheet1, "B1:B10", `[{"type":"3_color_scale","criteria":"=","min_type":"min","mid_type":"percentile","max_type":"max","min_color":"#F8696B","mid_color":"#FFEB84","max_color":"#63BE7B"}]`)
	// Hightlight cells rules: between...
	f.SetConditionalFormat(sheet1, "C1:C10", fmt.Sprintf(`[{"type":"cell","criteria":"between","format":%d,"minimum":"6","maximum":"8"}]`, format1))
	// Hightlight cells rules: Greater Than...
	f.SetConditionalFormat(sheet1, "D1:D10", fmt.Sprintf(`[{"type":"cell","criteria":">","format":%d,"value":"6"}]`, format3))
	// Hightlight cells rules: Equal To...
	f.SetConditionalFormat(sheet1, "E1:E10", fmt.Sprintf(`[{"type":"top","criteria":"=","format":%d}]`, format3))
	// Hightlight cells rules: Not Equal To...
	f.SetConditionalFormat(sheet1, "F1:F10", fmt.Sprintf(`[{"type":"unique","criteria":"=","format":%d}]`, format2))
	// Hightlight cells rules: Duplicate Values...
	f.SetConditionalFormat(sheet1, "G1:G10", fmt.Sprintf(`[{"type":"duplicate","criteria":"=","format":%d}]`, format2))
	// Top/Bottom rules: Top 10%.
	f.SetConditionalFormat(sheet1, "H1:H10", fmt.Sprintf(`[{"type":"top","criteria":"=","format":%d,"value":"6","percent":true}]`, format1))
	// Top/Bottom rules: Above Average...
	f.SetConditionalFormat(sheet1, "I1:I10", fmt.Sprintf(`[{"type":"average","criteria":"=","format":%d, "above_average": true}]`, format3))
	// Top/Bottom rules: Below Average...
	f.SetConditionalFormat(sheet1, "J1:J10", fmt.Sprintf(`[{"type":"average","criteria":"=","format":%d, "above_average": false}]`, format1))
	// Data Bars: Gradient Fill.
	f.SetConditionalFormat(sheet1, "K1:K10", `[{"type":"data_bar", "criteria":"=", "min_type":"min","max_type":"max","bar_color":"#638EC6"}]`)
	// Use a formula to determine which cells to format.
	f.SetConditionalFormat(sheet1, "L1:L10", fmt.Sprintf(`[{"type":"formula", "criteria":"L2<3", "format":%d}]`, format1))
	// Test set invalid format set in conditional format
	f.SetConditionalFormat(sheet1, "L1:L10", "")

	err = f.SaveAs(filepath.Join("test", "TestConditionalFormat.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Set conditional format with illegal valid type.
	f.SetConditionalFormat(sheet1, "K1:K10", `[{"type":"", "criteria":"=", "min_type":"min","max_type":"max","bar_color":"#638EC6"}]`)
	// Set conditional format with illegal criteria type.
	f.SetConditionalFormat(sheet1, "K1:K10", `[{"type":"data_bar", "criteria":"", "min_type":"min","max_type":"max","bar_color":"#638EC6"}]`)

	// Set conditional format with file without dxfs element shold not return error.
	f, err = OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	_, err = f.NewConditionalStyle(`{"font":{"color":"#9A0511"},"fill":{"type":"pattern","color":["#FEC7CE"],"pattern":1}}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}
}

func TestConditionalFormatError(t *testing.T) {
	f := NewFile()
	sheet1 := f.GetSheetName(1)

	fillCells(f, sheet1, 10, 15)

	// Set conditional format with illegal JSON string should return error
	_, err := f.NewConditionalStyle("")
	if !assert.EqualError(t, err, "unexpected end of JSON input") {
		t.FailNow()
	}
}

func TestSharedStrings(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "SharedStrings.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	f.GetRows("Sheet1")
}

func TestSetSheetRow(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	f.SetSheetRow("Sheet1", "B27", &[]interface{}{"cell", nil, int32(42), float64(42), time.Now().UTC()})

	assert.EqualError(t, f.SetSheetRow("Sheet1", "", &[]interface{}{"cell", nil, 2}),
		`cannot convert cell "" to coordinates: invalid cell name ""`)

	assert.EqualError(t, f.SetSheetRow("Sheet1", "B27", []interface{}{}), `pointer to slice expected`)
	assert.EqualError(t, f.SetSheetRow("Sheet1", "B27", &f), `pointer to slice expected`)
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetSheetRow.xlsx")))
}

func TestThemeColor(t *testing.T) {
	t.Log(ThemeColor("000000", -0.1))
	t.Log(ThemeColor("000000", 0))
	t.Log(ThemeColor("000000", 1))
}

func TestHSL(t *testing.T) {
	var hsl HSL
	t.Log(hsl.RGBA())
	t.Log(hslModel(hsl))
	t.Log(hslModel(color.Gray16{Y: uint16(1)}))
	t.Log(HSLToRGB(0, 1, 0.4))
	t.Log(HSLToRGB(0, 1, 0.6))
	t.Log(hueToRGB(0, 0, -1))
	t.Log(hueToRGB(0, 0, 2))
	t.Log(hueToRGB(0, 0, 1.0/7))
	t.Log(hueToRGB(0, 0, 0.4))
	t.Log(hueToRGB(0, 0, 2.0/4))
	t.Log(RGBToHSL(255, 255, 0))
	t.Log(RGBToHSL(0, 255, 255))
	t.Log(RGBToHSL(250, 100, 50))
	t.Log(RGBToHSL(50, 100, 250))
	t.Log(RGBToHSL(250, 50, 100))
}

func TestSearchSheet(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "SharedStrings.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Test search in a not exists worksheet.
	t.Log(f.SearchSheet("Sheet4", ""))
	// Test search a not exists value.
	t.Log(f.SearchSheet("Sheet1", "X"))
	t.Log(f.SearchSheet("Sheet1", "A"))
	// Test search the coordinates where the numerical value in the range of
	// "0-9" of Sheet1 is described by regular expression:
	t.Log(f.SearchSheet("Sheet1", "[0-9]", true))
}

func TestProtectSheet(t *testing.T) {
	f := NewFile()
	f.ProtectSheet("Sheet1", nil)
	f.ProtectSheet("Sheet1", &FormatSheetProtection{
		Password:      "password",
		EditScenarios: false,
	})

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestProtectSheet.xlsx")))
	// Test protect not exists worksheet.
	assert.EqualError(t, f.ProtectSheet("SheetN", nil), "sheet SheetN is not exist")
}

func TestUnprotectSheet(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	// Test unprotect not exists worksheet.
	assert.EqualError(t, f.UnprotectSheet("SheetN"), "sheet SheetN is not exist")

	f.UnprotectSheet("Sheet1")
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestUnprotectSheet.xlsx")))
}

func TestSetDefaultTimeStyle(t *testing.T) {
	f := NewFile()
	// Test set default time style on not exists worksheet.
	assert.EqualError(t, f.setDefaultTimeStyle("SheetN", "", 0), "sheet SheetN is not exist")
}

func prepareTestBook1() (*File, error) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if err != nil {
		return nil, err
	}

	err = f.AddPicture("Sheet2", "I9", filepath.Join("test", "images", "excel.jpg"),
		`{"x_offset": 140, "y_offset": 120, "hyperlink": "#Sheet2!D8", "hyperlink_type": "Location"}`)
	if err != nil {
		return nil, err
	}

	// Test add picture to worksheet with offset, external hyperlink and positioning.
	err = f.AddPicture("Sheet1", "F21", filepath.Join("test", "images", "excel.png"),
		`{"x_offset": 10, "y_offset": 10, "hyperlink": "https://github.com/360EntSecGroup-Skylar/excelize", "hyperlink_type": "External", "positioning": "oneCell"}`)
	if err != nil {
		return nil, err
	}

	file, err := ioutil.ReadFile(filepath.Join("test", "images", "excel.jpg"))
	if err != nil {
		return nil, err
	}

	err = f.AddPictureFromBytes("Sheet1", "Q1", "", "Excel Logo", ".jpg", file)
	if err != nil {
		return nil, err
	}

	return f, nil
}

func prepareTestBook3() (*File, error) {
	f := NewFile()
	f.NewSheet("Sheet1")
	f.NewSheet("XLSXSheet2")
	f.NewSheet("XLSXSheet3")
	f.SetCellInt("XLSXSheet2", "A23", 56)
	f.SetCellStr("Sheet1", "B20", "42")
	f.SetActiveSheet(0)

	err := f.AddPicture("Sheet1", "H2", filepath.Join("test", "images", "excel.gif"),
		`{"x_scale": 0.5, "y_scale": 0.5, "positioning": "absolute"}`)
	if err != nil {
		return nil, err
	}

	err = f.AddPicture("Sheet1", "C2", filepath.Join("test", "images", "excel.png"), "")
	if err != nil {
		return nil, err
	}

	return f, nil
}

func prepareTestBook4() (*File, error) {
	f := NewFile()
	f.SetColWidth("Sheet1", "B", "A", 12)
	f.SetColWidth("Sheet1", "A", "B", 12)
	f.GetColWidth("Sheet1", "A")
	f.GetColWidth("Sheet1", "C")

	return f, nil
}

func fillCells(f *File, sheet string, colCount, rowCount int) {
	for col := 1; col <= colCount; col++ {
		for row := 1; row <= rowCount; row++ {
			cell, _ := CoordinatesToCellName(col, row)
			f.SetCellStr(sheet, cell, cell)
		}
	}
}
