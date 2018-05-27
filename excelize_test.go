package excelize

import (
	"fmt"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"io/ioutil"
	"reflect"
	"strconv"
	"strings"
	"testing"
	"time"
)

func TestOpenFile(t *testing.T) {
	// Test update a XLSX file.
	xlsx, err := OpenFile("./test/Book1.xlsx")
	if err != nil {
		t.Error(err)
	}
	// Test get all the rows in a not exists worksheet.
	rows := xlsx.GetRows("Sheet4")
	// Test get all the rows in a worksheet.
	rows = xlsx.GetRows("Sheet2")
	for _, row := range rows {
		for _, cell := range row {
			t.Log(cell, "\t")
		}
		t.Log("\r\n")
	}
	xlsx.UpdateLinkedValue()
	xlsx.SetCellDefault("Sheet2", "A1", strconv.FormatFloat(float64(100.1588), 'f', -1, 32))
	xlsx.SetCellDefault("Sheet2", "A1", strconv.FormatFloat(float64(-100.1588), 'f', -1, 64))
	// Test set cell value with illegal row number.
	xlsx.SetCellDefault("Sheet2", "A", strconv.FormatFloat(float64(-100.1588), 'f', -1, 64))
	xlsx.SetCellInt("Sheet2", "A1", 100)
	// Test set cell integer value with illegal row number.
	xlsx.SetCellInt("Sheet2", "A", 100)
	xlsx.SetCellStr("Sheet2", "C11", "Knowns")
	// Test max characters in a cell.
	xlsx.SetCellStr("Sheet2", "D11", strings.Repeat("c", 32769))
	xlsx.NewSheet(":\\/?*[]Maximum 31 characters allowed in sheet title.")
	// Test set worksheet name with illegal name.
	xlsx.SetSheetName("Maximum 31 characters allowed i", "[Rename]:\\/?* Maximum 31 characters allowed in sheet title.")
	xlsx.SetCellInt("Sheet3", "A23", 10)
	xlsx.SetCellStr("Sheet3", "b230", "10")
	xlsx.SetCellStr("Sheet10", "b230", "10")
	// Test set cell string value with illegal row number.
	xlsx.SetCellStr("Sheet10", "A", "10")
	xlsx.SetActiveSheet(2)
	// Test get cell formula with given rows number.
	xlsx.GetCellFormula("Sheet1", "B19")
	// Test get cell formula with illegal worksheet name.
	xlsx.GetCellFormula("Sheet2", "B20")
	// Test get cell formula with illegal rows number.
	xlsx.GetCellFormula("Sheet1", "B20")
	xlsx.GetCellFormula("Sheet1", "B")
	// Test get shared cell formula
	xlsx.GetCellFormula("Sheet2", "H11")
	xlsx.GetCellFormula("Sheet2", "I11")
	getSharedForumula(&xlsxWorksheet{}, "")
	// Test read cell value with given illegal rows number.
	xlsx.GetCellValue("Sheet2", "a-1")
	xlsx.GetCellValue("Sheet2", "A")
	// Test read cell value with given lowercase column number.
	xlsx.GetCellValue("Sheet2", "a5")
	xlsx.GetCellValue("Sheet2", "C11")
	xlsx.GetCellValue("Sheet2", "D11")
	xlsx.GetCellValue("Sheet2", "D12")
	// Test SetCellValue function.
	xlsx.SetCellValue("Sheet2", "F1", " Hello")
	xlsx.SetCellValue("Sheet2", "G1", []byte("World"))
	xlsx.SetCellValue("Sheet2", "F2", 42)
	xlsx.SetCellValue("Sheet2", "F3", int8(1<<8/2-1))
	xlsx.SetCellValue("Sheet2", "F4", int16(1<<16/2-1))
	xlsx.SetCellValue("Sheet2", "F5", int32(1<<32/2-1))
	xlsx.SetCellValue("Sheet2", "F6", int64(1<<32/2-1))
	xlsx.SetCellValue("Sheet2", "F7", float32(42.65418))
	xlsx.SetCellValue("Sheet2", "F8", float64(-42.65418))
	xlsx.SetCellValue("Sheet2", "F9", float32(42))
	xlsx.SetCellValue("Sheet2", "F10", float64(42))
	xlsx.SetCellValue("Sheet2", "F11", uint(1<<32-1))
	xlsx.SetCellValue("Sheet2", "F12", uint8(1<<8-1))
	xlsx.SetCellValue("Sheet2", "F13", uint16(1<<16-1))
	xlsx.SetCellValue("Sheet2", "F14", uint32(1<<32-1))
	xlsx.SetCellValue("Sheet2", "F15", uint64(1<<32-1))
	xlsx.SetCellValue("Sheet2", "F16", true)
	xlsx.SetCellValue("Sheet2", "F17", complex64(5+10i))
	t.Log(letterOnlyMapF('x'))
	t.Log(deepCopy(nil, nil))
	shiftJulianToNoon(1, -0.6)
	timeFromExcelTime(61, true)
	timeFromExcelTime(62, true)
	// Test boolean write
	booltest := []struct {
		value    bool
		expected string
	}{
		{false, "0"},
		{true, "1"},
	}
	for _, test := range booltest {
		xlsx.SetCellValue("Sheet2", "F16", test.value)
		value := xlsx.GetCellValue("Sheet2", "F16")
		if value != test.expected {
			t.Errorf(`Expecting result of xlsx.SetCellValue("Sheet2", "F16", %v) to be %v (false), got: %s `, test.value, test.expected, value)
		}
	}
	xlsx.SetCellValue("Sheet2", "G2", nil)
	xlsx.SetCellValue("Sheet2", "G4", time.Now())
	// 02:46:40
	xlsx.SetCellValue("Sheet2", "G5", time.Duration(1e13))
	// Test completion column.
	xlsx.SetCellValue("Sheet2", "M2", nil)
	// Test read cell value with given axis large than exists row.
	xlsx.GetCellValue("Sheet2", "E231")
	// Test get active worksheet of XLSX and get worksheet name of XLSX by given worksheet index.
	xlsx.GetSheetName(xlsx.GetActiveSheetIndex())
	// Test get worksheet index of XLSX by given worksheet name.
	xlsx.GetSheetIndex("Sheet1")
	// Test get worksheet name of XLSX by given invalid worksheet index.
	xlsx.GetSheetName(4)
	// Test get worksheet map of XLSX.
	xlsx.GetSheetMap()
	for i := 1; i <= 300; i++ {
		xlsx.SetCellStr("Sheet3", "c"+strconv.Itoa(i), strconv.Itoa(i))
	}
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
	// Test write file to not exist directory.
	err = xlsx.SaveAs("")
	if err != nil {
		t.Log(err)
	}
}

func TestAddPicture(t *testing.T) {
	xlsx, err := OpenFile("./test/Book1.xlsx")
	if err != nil {
		t.Error(err)
	}
	// Test add picture to worksheet with offset and location hyperlink.
	err = xlsx.AddPicture("Sheet2", "I9", "./test/images/excel.jpg", `{"x_offset": 140, "y_offset": 120, "hyperlink": "#Sheet2!D8", "hyperlink_type": "Location"}`)
	if err != nil {
		t.Error(err)
	}
	// Test add picture to worksheet with offset, external hyperlink and positioning.
	err = xlsx.AddPicture("Sheet1", "F21", "./test/images/excel.png", `{"x_offset": 10, "y_offset": 10, "hyperlink": "https://github.com/360EntSecGroup-Skylar/excelize", "hyperlink_type": "External", "positioning": "oneCell"}`)
	if err != nil {
		t.Error(err)
	}
	// Test add picture to worksheet with invalid file path.
	err = xlsx.AddPicture("Sheet1", "G21", "./test/images/excel.icon", "")
	if err != nil {
		t.Log(err)
	}
	// Test add picture to worksheet with unsupport file type.
	err = xlsx.AddPicture("Sheet1", "G21", "./test/Book1.xlsx", "")
	if err != nil {
		t.Log(err)
	}
	// Test write file to given path.
	err = xlsx.SaveAs("./test/Book2.xlsx")
	if err != nil {
		t.Error(err)
	}
}

func TestBrokenFile(t *testing.T) {
	// Test write file with broken file struct.
	xlsx := File{}
	err := xlsx.Save()
	if err != nil {
		t.Log(err)
	}
	// Test write file with broken file struct with given path.
	err = xlsx.SaveAs("./test/Book3.xlsx")
	if err != nil {
		t.Log(err)
	}

	// Test set active sheet without BookViews and Sheets maps in xl/workbook.xml.
	f3, err := OpenFile("./test/badWorkbook.xlsx")
	f3.GetActiveSheetIndex()
	f3.SetActiveSheet(2)
	if err != nil {
		t.Log(err)
	}

	// Test open a XLSX file with given illegal path.
	_, err = OpenFile("./test/Book.xlsx")
	if err != nil {
		t.Log(err)
	}
}

func TestNewFile(t *testing.T) {
	// Test create a XLSX file.
	xlsx := NewFile()
	xlsx.NewSheet("Sheet1")
	xlsx.NewSheet("XLSXSheet2")
	xlsx.NewSheet("XLSXSheet3")
	xlsx.SetCellInt("XLSXSheet2", "A23", 56)
	xlsx.SetCellStr("Sheet1", "B20", "42")
	xlsx.SetActiveSheet(0)
	// Test add picture to sheet with scaling and positioning.
	err := xlsx.AddPicture("Sheet1", "H2", "./test/images/excel.gif", `{"x_scale": 0.5, "y_scale": 0.5, "positioning": "absolute"}`)
	if err != nil {
		t.Error(err)
	}
	// Test add picture to worksheet with invalid formatset
	err = xlsx.AddPicture("Sheet1", "C2", "./test/images/excel.png", "")
	if err != nil {
		t.Log(err)
	}
	err = xlsx.SaveAs("./test/Book3.xlsx")
	if err != nil {
		t.Error(err)
	}
}

func TestColWidth(t *testing.T) {
	xlsx := NewFile()
	xlsx.SetColWidth("Sheet1", "B", "A", 12)
	xlsx.SetColWidth("Sheet1", "A", "B", 12)
	xlsx.GetColWidth("Sheet1", "A")
	xlsx.GetColWidth("Sheet1", "C")
	err := xlsx.SaveAs("./test/Book4.xlsx")
	if err != nil {
		t.Error(err)
	}
	convertRowHeightToPixels(0)
}

func TestRowHeight(t *testing.T) {
	xlsx := NewFile()
	xlsx.SetRowHeight("Sheet1", 1, 50)
	xlsx.SetRowHeight("Sheet1", 4, 90)
	t.Log(xlsx.GetRowHeight("Sheet1", 1))
	t.Log(xlsx.GetRowHeight("Sheet1", 0))
	err := xlsx.SaveAs("./test/Book5.xlsx")
	if err != nil {
		t.Error(err)
	}
	convertColWidthToPixels(0)
}

func TestSetCellHyperLink(t *testing.T) {
	xlsx, err := OpenFile("./test/Book1.xlsx")
	if err != nil {
		t.Log(err)
	}
	// Test set cell hyperlink in a work sheet already have hyperlinks.
	xlsx.SetCellHyperLink("Sheet1", "B19", "https://github.com/360EntSecGroup-Skylar/excelize", "External")
	// Test add first hyperlink in a work sheet.
	xlsx.SetCellHyperLink("Sheet2", "C1", "https://github.com/360EntSecGroup-Skylar/excelize", "External")
	// Test add Location hyperlink in a work sheet.
	xlsx.SetCellHyperLink("Sheet2", "D6", "Sheet1!D8", "Location")
	xlsx.SetCellHyperLink("Sheet2", "C3", "Sheet1!D8", "")
	xlsx.SetCellHyperLink("Sheet2", "", "Sheet1!D60", "Location")
	err = xlsx.Save()
	if err != nil {
		t.Error(err)
	}
}

func TestGetCellHyperLink(t *testing.T) {
	xlsx, err := OpenFile("./test/Book1.xlsx")
	if err != nil {
		t.Error(err)
	}
	link, target := xlsx.GetCellHyperLink("Sheet1", "")
	t.Log(link, target)
	link, target = xlsx.GetCellHyperLink("Sheet1", "B19")
	t.Log(link, target)
	link, target = xlsx.GetCellHyperLink("Sheet2", "D6")
	t.Log(link, target)
	link, target = xlsx.GetCellHyperLink("Sheet3", "H3")
	t.Log(link, target)
}

func TestSetCellFormula(t *testing.T) {
	xlsx, err := OpenFile("./test/Book1.xlsx")
	if err != nil {
		t.Error(err)
	}
	xlsx.SetCellFormula("Sheet1", "B19", "SUM(Sheet2!D2,Sheet2!D11)")
	xlsx.SetCellFormula("Sheet1", "C19", "SUM(Sheet2!D2,Sheet2!D9)")
	// Test set cell formula with illegal rows number.
	xlsx.SetCellFormula("Sheet1", "C", "SUM(Sheet2!D2,Sheet2!D9)")
	err = xlsx.Save()
	if err != nil {
		t.Error(err)
	}
}

func TestSetSheetBackground(t *testing.T) {
	xlsx, err := OpenFile("./test/Book1.xlsx")
	if err != nil {
		t.Error(err)
	}
	err = xlsx.SetSheetBackground("Sheet2", "./test/images/background.png")
	if err != nil {
		t.Log(err)
	}
	err = xlsx.SetSheetBackground("Sheet2", "./test/Book1.xlsx")
	if err != nil {
		t.Log(err)
	}
	err = xlsx.SetSheetBackground("Sheet2", "./test/images/background.jpg")
	if err != nil {
		t.Error(err)
	}
	err = xlsx.SetSheetBackground("Sheet2", "./test/images/background.jpg")
	if err != nil {
		t.Error(err)
	}
	err = xlsx.Save()
	if err != nil {
		t.Error(err)
	}
}

func TestMergeCell(t *testing.T) {
	xlsx, err := OpenFile("./test/Book1.xlsx")
	if err != nil {
		t.Error(err)
	}
	xlsx.MergeCell("Sheet1", "D9", "D9")
	xlsx.MergeCell("Sheet1", "D9", "E9")
	xlsx.MergeCell("Sheet1", "H14", "G13")
	xlsx.MergeCell("Sheet1", "C9", "D8")
	xlsx.MergeCell("Sheet1", "F11", "G13")
	xlsx.MergeCell("Sheet1", "H7", "B15")
	xlsx.MergeCell("Sheet1", "D11", "F13")
	xlsx.MergeCell("Sheet1", "G10", "K12")
	xlsx.SetCellValue("Sheet1", "G11", "set value in merged cell")
	xlsx.SetCellInt("Sheet1", "H11", 100)
	xlsx.SetCellValue("Sheet1", "I11", float64(0.5))
	xlsx.SetCellHyperLink("Sheet1", "J11", "https://github.com/360EntSecGroup-Skylar/excelize", "External")
	xlsx.SetCellFormula("Sheet1", "G12", "SUM(Sheet1!B19,Sheet1!C19)")
	xlsx.GetCellValue("Sheet1", "H11")
	xlsx.GetCellValue("Sheet2", "A6") // Merged cell ref is single coordinate.
	xlsx.GetCellFormula("Sheet1", "G12")
	err = xlsx.Save()
	if err != nil {
		t.Error(err)
	}
}

func TestSetCellStyleAlignment(t *testing.T) {
	xlsx, err := OpenFile("./test/Book2.xlsx")
	if err != nil {
		t.Error(err)
	}
	var style int
	style, err = xlsx.NewStyle(`{"alignment":{"horizontal":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":45,"vertical":"top","wrap_text":true}}`)
	if err != nil {
		t.Error(err)
	}
	xlsx.SetCellStyle("Sheet1", "A22", "A22", style)
	// Test set cell style with given illegal rows number.
	xlsx.SetCellStyle("Sheet1", "A", "A22", style)
	xlsx.SetCellStyle("Sheet1", "A22", "A", style)
	// Test get cell style with given illegal rows number.
	xlsx.GetCellStyle("Sheet1", "A")
	err = xlsx.Save()
	if err != nil {
		t.Error(err)
	}
}

func TestSetCellStyleBorder(t *testing.T) {
	xlsx, err := OpenFile("./test/Book2.xlsx")
	if err != nil {
		t.Error(err)
	}
	var style int
	// Test set border with invalid style parameter.
	style, err = xlsx.NewStyle("")
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet1", "J21", "L25", style)

	// Test set border with invalid style index number.
	style, err = xlsx.NewStyle(`{"border":[{"type":"left","color":"0000FF","style":-1},{"type":"top","color":"00FF00","style":14},{"type":"bottom","color":"FFFF00","style":5},{"type":"right","color":"FF0000","style":6},{"type":"diagonalDown","color":"A020F0","style":9},{"type":"diagonalUp","color":"A020F0","style":8}]}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet1", "J21", "L25", style)

	// Test set border on overlapping area with vertical variants shading styles gradient fill.
	style, err = xlsx.NewStyle(`{"border":[{"type":"left","color":"0000FF","style":2},{"type":"top","color":"00FF00","style":12},{"type":"bottom","color":"FFFF00","style":5},{"type":"right","color":"FF0000","style":6},{"type":"diagonalDown","color":"A020F0","style":9},{"type":"diagonalUp","color":"A020F0","style":8}]}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet1", "J21", "L25", style)

	style, err = xlsx.NewStyle(`{"border":[{"type":"left","color":"0000FF","style":2},{"type":"top","color":"00FF00","style":3},{"type":"bottom","color":"FFFF00","style":4},{"type":"right","color":"FF0000","style":5},{"type":"diagonalDown","color":"A020F0","style":6},{"type":"diagonalUp","color":"A020F0","style":7}],"fill":{"type":"gradient","color":["#FFFFFF","#E0EBF5"],"shading":1}}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet1", "M28", "K24", style)

	style, err = xlsx.NewStyle(`{"border":[{"type":"left","color":"0000FF","style":2},{"type":"top","color":"00FF00","style":3},{"type":"bottom","color":"FFFF00","style":4},{"type":"right","color":"FF0000","style":5},{"type":"diagonalDown","color":"A020F0","style":6},{"type":"diagonalUp","color":"A020F0","style":7}],"fill":{"type":"gradient","color":["#FFFFFF","#E0EBF5"],"shading":4}}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet1", "M28", "K24", style)

	// Test set border and solid style pattern fill for a single cell.
	style, err = xlsx.NewStyle(`{"border":[{"type":"left","color":"0000FF","style":8},{"type":"top","color":"00FF00","style":9},{"type":"bottom","color":"FFFF00","style":10},{"type":"right","color":"FF0000","style":11},{"type":"diagonalDown","color":"A020F0","style":12},{"type":"diagonalUp","color":"A020F0","style":13}],"fill":{"type":"pattern","color":["#E0EBF5"],"pattern":1}}`)
	if err != nil {
		t.Log(err)
	}

	xlsx.SetCellStyle("Sheet1", "O22", "O22", style)
	err = xlsx.Save()
	if err != nil {
		t.Error(err)
	}
}

func TestSetCellStyleNumberFormat(t *testing.T) {
	xlsx, err := OpenFile("./test/Book2.xlsx")
	if err != nil {
		t.Error(err)
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
				xlsx.SetCellValue("Sheet2", c, v)
			} else {
				xlsx.SetCellValue("Sheet2", c, val)
			}
			style, err := xlsx.NewStyle(`{"fill":{"type":"gradient","color":["#FFFFFF","#E0EBF5"],"shading":5},"number_format": ` + strconv.Itoa(d) + `}`)
			if err != nil {
				t.Log(err)
			}
			xlsx.SetCellStyle("Sheet2", c, c, style)
			t.Log(xlsx.GetCellValue("Sheet2", c))
		}
	}
	var style int
	style, err = xlsx.NewStyle(`{"number_format":-1}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet2", "L33", "L33", style)
	err = xlsx.Save()
	if err != nil {
		t.Error(err)
	}
}

func TestSetCellStyleCurrencyNumberFormat(t *testing.T) {
	xlsx, err := OpenFile("./test/Book3.xlsx")
	if err != nil {
		t.Error(err)
	}
	xlsx.SetCellValue("Sheet1", "A1", 56)
	xlsx.SetCellValue("Sheet1", "A2", -32.3)
	var style int
	style, err = xlsx.NewStyle(`{"number_format": 188, "decimal_places": -1}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet1", "A1", "A1", style)
	style, err = xlsx.NewStyle(`{"number_format": 188, "decimal_places": 31, "negred": true}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet1", "A2", "A2", style)

	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}

	xlsx, err = OpenFile("./test/Book4.xlsx")
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellValue("Sheet1", "A1", 42920.5)
	xlsx.SetCellValue("Sheet1", "A2", 42920.5)

	style, err = xlsx.NewStyle(`{"number_format": 26, "lang": "zh-tw"}`)
	if err != nil {
		t.Log(err)
	}
	style, err = xlsx.NewStyle(`{"number_format": 27}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet1", "A1", "A1", style)
	style, err = xlsx.NewStyle(`{"number_format": 31, "lang": "ko-kr"}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet1", "A2", "A2", style)

	style, err = xlsx.NewStyle(`{"number_format": 71, "lang": "th-th"}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet1", "A2", "A2", style)

	err = xlsx.Save()
	if err != nil {
		t.Error(err)
	}
}

func TestSetCellStyleCustomNumberFormat(t *testing.T) {
	xlsx := NewFile()
	xlsx.SetCellValue("Sheet1", "A1", 42920.5)
	xlsx.SetCellValue("Sheet1", "A2", 42920.5)
	style, err := xlsx.NewStyle(`{"custom_number_format": "[$-380A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@"}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet1", "A1", "A1", style)
	style, err = xlsx.NewStyle(`{"custom_number_format": "[$-380A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@"}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet1", "A2", "A2", style)
	err = xlsx.SaveAs("./test/Book_custom_number_format.xlsx")
	if err != nil {
		t.Error(err)
	}
}

func TestSetCellStyleFill(t *testing.T) {
	xlsx, err := OpenFile("./test/Book2.xlsx")
	if err != nil {
		t.Error(err)
	}
	var style int
	// Test set fill for cell with invalid parameter.
	style, err = xlsx.NewStyle(`{"fill":{"type":"gradient","color":["#FFFFFF","#E0EBF5"],"shading":6}}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet1", "O23", "O23", style)

	style, err = xlsx.NewStyle(`{"fill":{"type":"gradient","color":["#FFFFFF"],"shading":1}}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet1", "O23", "O23", style)

	style, err = xlsx.NewStyle(`{"fill":{"type":"pattern","color":[],"pattern":1}}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet1", "O23", "O23", style)

	style, err = xlsx.NewStyle(`{"fill":{"type":"pattern","color":["#E0EBF5"],"pattern":19}}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet1", "O23", "O23", style)

	err = xlsx.Save()
	if err != nil {
		t.Error(err)
	}
}

func TestSetCellStyleFont(t *testing.T) {
	xlsx, err := OpenFile("./test/Book2.xlsx")
	if err != nil {
		t.Error(err)
	}
	var style int
	style, err = xlsx.NewStyle(`{"font":{"bold":true,"italic":true,"family":"Berlin Sans FB Demi","size":36,"color":"#777777","underline":"single"}}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet2", "A1", "A1", style)

	style, err = xlsx.NewStyle(`{"font":{"italic":true,"underline":"double"}}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet2", "A2", "A2", style)

	style, err = xlsx.NewStyle(`{"font":{"bold":true}}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet2", "A3", "A3", style)

	style, err = xlsx.NewStyle(`{"font":{"bold":true,"family":"","size":0,"color":"","underline":""}}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet2", "A4", "A4", style)

	style, err = xlsx.NewStyle(`{"font":{"color":"#777777"}}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet2", "A5", "A5", style)
	err = xlsx.Save()
	if err != nil {
		t.Error(err)
	}
}

func TestSetCellStyleProtection(t *testing.T) {
	xlsx, err := OpenFile("./test/Book2.xlsx")
	if err != nil {
		t.Error(err)
	}
	var style int
	style, err = xlsx.NewStyle(`{"protection":{"hidden":true, "locked":true}}`)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellStyle("Sheet2", "A6", "A6", style)
	err = xlsx.Save()
	if err != nil {
		t.Error(err)
	}
}

func TestSetDeleteSheet(t *testing.T) {
	xlsx, err := OpenFile("./test/Book3.xlsx")
	if err != nil {
		t.Error(err)
	}
	xlsx.DeleteSheet("XLSXSheet3")
	err = xlsx.Save()
	if err != nil {
		t.Error(err)
	}
	xlsx, err = OpenFile("./test/Book4.xlsx")
	if err != nil {
		t.Error(err)
	}
	xlsx.DeleteSheet("Sheet1")
	xlsx.AddComment("Sheet1", "A1", "")
	xlsx.AddComment("Sheet1", "A1", `{"author":"Excelize: ","text":"This is a comment."}`)
	err = xlsx.SaveAs("./test/Book_delete_sheet.xlsx")
	if err != nil {
		t.Error(err)
	}
}

func TestGetPicture(t *testing.T) {
	xlsx, err := OpenFile("./test/Book2.xlsx")
	if err != nil {
		t.Log(err)
	}
	file, raw := xlsx.GetPicture("Sheet1", "F21")
	if file == "" {
		err = ioutil.WriteFile(file, raw, 0644)
		if err != nil {
			t.Log(err)
		}
	}
	// Try to get picture from a worksheet that doesn't contain any images.
	file, raw = xlsx.GetPicture("Sheet3", "I9")
	if file != "" {
		err = ioutil.WriteFile(file, raw, 0644)
		if err != nil {
			t.Log(err)
		}
	}
	// Try to get picture from a cell that doesn't contain an image.
	file, raw = xlsx.GetPicture("Sheet2", "A2")
	t.Log(file, len(raw))
	xlsx.getDrawingRelationships("xl/worksheets/_rels/sheet1.xml.rels", "rId8")
	xlsx.getDrawingRelationships("", "")
	xlsx.getSheetRelationshipsTargetByID("", "")
	xlsx.deleteSheetRelationships("", "")
}

func TestSheetVisibility(t *testing.T) {
	xlsx, err := OpenFile("./test/Book2.xlsx")
	if err != nil {
		t.Error(err)
	}
	xlsx.SetSheetVisible("Sheet2", false)
	xlsx.SetSheetVisible("Sheet1", false)
	xlsx.SetSheetVisible("Sheet1", true)
	xlsx.GetSheetVisible("Sheet1")
	err = xlsx.Save()
	if err != nil {
		t.Error(err)
	}
}

func TestRowVisibility(t *testing.T) {
	xlsx, err := OpenFile("./test/Book2.xlsx")
	if err != nil {
		t.Error(err)
	}
	xlsx.SetRowVisible("Sheet3", 2, false)
	xlsx.SetRowVisible("Sheet3", 2, true)
	xlsx.GetRowVisible("Sheet3", 2)
	err = xlsx.Save()
	if err != nil {
		t.Error(err)
	}
}

func TestColumnVisibility(t *testing.T) {
	xlsx, err := OpenFile("./test/Book2.xlsx")
	if err != nil {
		t.Error(err)
	}
	xlsx.SetColVisible("Sheet1", "F", false)
	xlsx.SetColVisible("Sheet1", "F", true)
	xlsx.GetColVisible("Sheet1", "F")
	xlsx.SetColVisible("Sheet3", "E", false)
	err = xlsx.Save()
	if err != nil {
		t.Error(err)
	}
	xlsx, err = OpenFile("./test/Book3.xlsx")
	if err != nil {
		t.Error(err)
	}
	xlsx.GetColVisible("Sheet1", "B")
}

func TestCopySheet(t *testing.T) {
	xlsx, err := OpenFile("./test/Book2.xlsx")
	if err != nil {
		t.Error(err)
	}
	err = xlsx.CopySheet(0, -1)
	if err != nil {
		t.Log(err)
	}
	idx := xlsx.NewSheet("CopySheet")
	err = xlsx.CopySheet(1, idx)
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellValue("Sheet4", "F1", "Hello")
	if xlsx.GetCellValue("Sheet1", "F1") == "Hello" {
		t.Error("Invalid value \"Hello\" in Sheet1")
	}
	err = xlsx.Save()
	if err != nil {
		t.Error(err)
	}
}

func TestAddTable(t *testing.T) {
	xlsx, err := OpenFile("./test/Book2.xlsx")
	if err != nil {
		t.Error(err)
	}
	err = xlsx.AddTable("Sheet1", "B26", "A21", `{}`)
	if err != nil {
		t.Error(err)
	}
	err = xlsx.AddTable("Sheet2", "A2", "B5", ``)
	if err != nil {
		t.Log(err)
	}
	err = xlsx.AddTable("Sheet2", "A2", "B5", `{"table_name":"table","table_style":"TableStyleMedium2", "show_first_column":true,"show_last_column":true,"show_row_stripes":false,"show_column_stripes":true}`)
	if err != nil {
		t.Error(err)
	}
	err = xlsx.AddTable("Sheet2", "F1", "F1", `{"table_style":"TableStyleMedium8"}`)
	if err != nil {
		t.Error(err)
	}
	err = xlsx.Save()
	if err != nil {
		t.Error(err)
	}
}

func TestAddShape(t *testing.T) {
	xlsx, err := OpenFile("./test/Book2.xlsx")
	if err != nil {
		t.Error(err)
	}
	xlsx.AddShape("Sheet1", "A30", `{"type":"rect","paragraph":[{"text":"Rectangle","font":{"color":"CD5C5C"}},{"text":"Shape","font":{"bold":true,"color":"2980B9"}}]}`)
	xlsx.AddShape("Sheet1", "B30", `{"type":"rect","paragraph":[{"text":"Rectangle"},{}]}`)
	xlsx.AddShape("Sheet1", "C30", `{"type":"rect","paragraph":[]}`)
	xlsx.AddShape("Sheet3", "H1", `{"type":"ellipseRibbon", "color":{"line":"#4286f4","fill":"#8eb9ff"}, "paragraph":[{"font":{"bold":true,"italic":true,"family":"Berlin Sans FB Demi","size":36,"color":"#777777","underline":"single"}}], "height": 90}`)
	xlsx.AddShape("Sheet3", "H1", "")
	err = xlsx.Save()
	if err != nil {
		t.Error(err)
	}
}

func TestAddComments(t *testing.T) {
	xlsx, err := OpenFile("./test/Book2.xlsx")
	if err != nil {
		t.Error(err)
	}
	s := strings.Repeat("c", 32768)
	xlsx.AddComment("Sheet1", "A30", `{"author":"`+s+`","text":"`+s+`"}`)
	xlsx.AddComment("Sheet2", "B7", `{"author":"Excelize: ","text":"This is a comment."}`)
	err = xlsx.Save()
	if err != nil {
		t.Error(err)
	}
}

func TestAutoFilter(t *testing.T) {
	xlsx, err := OpenFile("./test/Book2.xlsx")
	if err != nil {
		t.Error(err)
	}
	formats := []string{``,
		`{"column":"B","expression":"x != blanks"}`,
		`{"column":"B","expression":"x == blanks"}`,
		`{"column":"B","expression":"x != nonblanks"}`,
		`{"column":"B","expression":"x == nonblanks"}`,
		`{"column":"B","expression":"x <= 1 and x >= 2"}`,
		`{"column":"B","expression":"x == 1 or x == 2"}`,
		`{"column":"B","expression":"x == 1 or x == 2*"}`,
		`{"column":"B","expression":"x <= 1 and x >= blanks"}`,
		`{"column":"B","expression":"x -- y or x == *2*"}`,
		`{"column":"B","expression":"x != y or x ? *2"}`,
		`{"column":"B","expression":"x -- y o r x == *2"}`,
		`{"column":"B","expression":"x -- y"}`,
		`{"column":"A","expression":"x -- y"}`,
	}
	for _, format := range formats {
		err = xlsx.AutoFilter("Sheet3", "D4", "B1", format)
		t.Log(err)
	}
	err = xlsx.Save()
	if err != nil {
		t.Error(err)
	}
}

func TestAddChart(t *testing.T) {
	xlsx, err := OpenFile("./test/Book1.xlsx")
	if err != nil {
		t.Error(err)
	}
	categories := map[string]string{"A30": "Small", "A31": "Normal", "A32": "Large", "B29": "Apple", "C29": "Orange", "D29": "Pear"}
	values := map[string]int{"B30": 2, "C30": 3, "D30": 3, "B31": 5, "C31": 2, "D31": 4, "B32": 6, "C32": 7, "D32": 8}
	for k, v := range categories {
		xlsx.SetCellValue("Sheet1", k, v)
	}
	for k, v := range values {
		xlsx.SetCellValue("Sheet1", k, v)
	}
	xlsx.AddChart("Sheet1", "P1", "")
	xlsx.AddChart("Sheet1", "P1", `{"type":"col","series":[{"name":"Sheet1!$A$30","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$30:$D$30"},{"name":"Sheet1!$A$31","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$31:$D$31"},{"name":"Sheet1!$A$32","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"left","show_legend_key":false},"title":{"name":"Fruit 2D Column Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"zero"}`)
	xlsx.AddChart("Sheet1", "X1", `{"type":"colStacked","series":[{"name":"Sheet1!$A$30","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$30:$D$30"},{"name":"Sheet1!$A$31","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$31:$D$31"},{"name":"Sheet1!$A$32","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"left","show_legend_key":false},"title":{"name":"Fruit 2D Stacked Column Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"zero"}`)
	xlsx.AddChart("Sheet1", "P16", `{"type":"colPercentStacked","series":[{"name":"Sheet1!$A$30","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$30:$D$30"},{"name":"Sheet1!$A$31","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$31:$D$31"},{"name":"Sheet1!$A$32","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"left","show_legend_key":false},"title":{"name":"Fruit 100% Stacked Column Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"zero"}`)
	xlsx.AddChart("Sheet1", "X16", `{"type":"col3DClustered","series":[{"name":"Sheet1!$A$30","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$30:$D$30"},{"name":"Sheet1!$A$31","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$31:$D$31"},{"name":"Sheet1!$A$32","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"bottom","show_legend_key":false},"title":{"name":"Fruit 3D Clustered Column Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"zero"}`)
	xlsx.AddChart("Sheet1", "P30", `{"type":"col3DStacked","series":[{"name":"Sheet1!$A$30","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$30:$D$30"},{"name":"Sheet1!$A$31","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$31:$D$31"},{"name":"Sheet1!$A$32","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"left","show_legend_key":false},"title":{"name":"Fruit 3D 100% Stacked Bar Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"zero"}`)
	xlsx.AddChart("Sheet1", "X30", `{"type":"col3DPercentStacked","series":[{"name":"Sheet1!$A$30","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$30:$D$30"},{"name":"Sheet1!$A$31","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$31:$D$31"},{"name":"Sheet1!$A$32","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"left","show_legend_key":false},"title":{"name":"Fruit 3D 100% Stacked Column Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"zero"}`)
	xlsx.AddChart("Sheet1", "P45", `{"type":"col3D","series":[{"name":"Sheet1!$A$30","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$30:$D$30"},{"name":"Sheet1!$A$31","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$31:$D$31"},{"name":"Sheet1!$A$32","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"left","show_legend_key":false},"title":{"name":"Fruit 3D Column Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"zero"}`)
	xlsx.AddChart("Sheet2", "P1", `{"type":"radar","series":[{"name":"Sheet1!$A$30","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$30:$D$30"},{"name":"Sheet1!$A$31","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$31:$D$31"},{"name":"Sheet1!$A$32","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"top_right","show_legend_key":false},"title":{"name":"Fruit Radar Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"span"}`)
	xlsx.AddChart("Sheet2", "X1", `{"type":"scatter","series":[{"name":"Sheet1!$A$30","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$30:$D$30"},{"name":"Sheet1!$A$31","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$31:$D$31"},{"name":"Sheet1!$A$32","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"bottom","show_legend_key":false},"title":{"name":"Fruit Scatter Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"zero"}`)
	xlsx.AddChart("Sheet2", "P16", `{"type":"doughnut","series":[{"name":"Sheet1!$A$30","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$30:$D$30"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"right","show_legend_key":false},"title":{"name":"Fruit Doughnut Chart"},"plotarea":{"show_bubble_size":false,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":false,"show_val":false},"show_blanks_as":"zero"}`)
	xlsx.AddChart("Sheet2", "X16", `{"type":"line","series":[{"name":"Sheet1!$A$30","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$30:$D$30"},{"name":"Sheet1!$A$31","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$31:$D$31"},{"name":"Sheet1!$A$32","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"top","show_legend_key":false},"title":{"name":"Fruit Line Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"zero"}`)
	xlsx.AddChart("Sheet2", "P32", `{"type":"pie3D","series":[{"name":"Sheet1!$A$30","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$30:$D$30"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"bottom","show_legend_key":false},"title":{"name":"Fruit 3D Pie Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":false,"show_val":false},"show_blanks_as":"zero"}`)
	xlsx.AddChart("Sheet2", "X32", `{"type":"pie","series":[{"name":"Sheet1!$A$30","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$30:$D$30"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"bottom","show_legend_key":false},"title":{"name":"Fruit Pie Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":false,"show_val":false},"show_blanks_as":"gap"}`)
	xlsx.AddChart("Sheet2", "P48", `{"type":"bar","series":[{"name":"Sheet1!$A$30","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$30:$D$30"},{"name":"Sheet1!$A$31","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$31:$D$31"},{"name":"Sheet1!$A$32","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"left","show_legend_key":false},"title":{"name":"Fruit 2D Clustered Bar Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"zero"}`)
	xlsx.AddChart("Sheet2", "X48", `{"type":"barStacked","series":[{"name":"Sheet1!$A$30","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$30:$D$30"},{"name":"Sheet1!$A$31","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$31:$D$31"},{"name":"Sheet1!$A$32","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"left","show_legend_key":false},"title":{"name":"Fruit 2D Stacked Bar Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"zero"}`)
	xlsx.AddChart("Sheet2", "P64", `{"type":"barPercentStacked","series":[{"name":"Sheet1!$A$30","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$30:$D$30"},{"name":"Sheet1!$A$31","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$31:$D$31"},{"name":"Sheet1!$A$32","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"left","show_legend_key":false},"title":{"name":"Fruit 2D Stacked 100% Bar Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"zero"}`)
	xlsx.AddChart("Sheet2", "X64", `{"type":"bar3DClustered","series":[{"name":"Sheet1!$A$30","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$30:$D$30"},{"name":"Sheet1!$A$31","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$31:$D$31"},{"name":"Sheet1!$A$32","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"left","show_legend_key":false},"title":{"name":"Fruit 3D Clustered Bar Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"zero"}`)
	xlsx.AddChart("Sheet2", "P80", `{"type":"bar3DStacked","series":[{"name":"Sheet1!$A$30","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$30:$D$30"},{"name":"Sheet1!$A$31","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$31:$D$31"},{"name":"Sheet1!$A$32","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"left","show_legend_key":false},"title":{"name":"Fruit 3D Stacked Bar Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"zero","y_axis":{"maximum":7.5,"minimum":0.5}}`)
	xlsx.AddChart("Sheet2", "X80", `{"type":"bar3DPercentStacked","series":[{"name":"Sheet1!$A$30","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$30:$D$30"},{"name":"Sheet1!$A$31","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$31:$D$31"},{"name":"Sheet1!$A$32","categories":"Sheet1!$B$29:$D$29","values":"Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"left","show_legend_key":false},"title":{"name":"Fruit 3D 100% Stacked Bar Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"zero","x_axis":{"reverse_order":true,"maximum":0,"minimum":0},"y_axis":{"reverse_order":true,"maximum":0,"minimum":0}}`)
	// Save xlsx file by the given path.
	err = xlsx.SaveAs("./test/Book_addchart.xlsx")
	if err != nil {
		t.Error(err)
	}
}

func TestInsertCol(t *testing.T) {
	xlsx := NewFile()
	for j := 1; j <= 10; j++ {
		for i := 0; i <= 10; i++ {
			axis := ToAlphaString(i) + strconv.Itoa(j)
			xlsx.SetCellStr("Sheet1", axis, axis)
		}
	}
	xlsx.SetCellHyperLink("Sheet1", "A5", "https://github.com/360EntSecGroup-Skylar/excelize", "External")
	xlsx.MergeCell("Sheet1", "A1", "C3")
	err := xlsx.AutoFilter("Sheet1", "A2", "B2", `{"column":"B","expression":"x != blanks"}`)
	t.Log(err)
	xlsx.InsertCol("Sheet1", "A")
	err = xlsx.SaveAs("./test/Book_insertcol.xlsx")
	if err != nil {
		t.Error(err)
	}
}

func TestRemoveCol(t *testing.T) {
	xlsx := NewFile()
	for j := 1; j <= 10; j++ {
		for i := 0; i <= 10; i++ {
			axis := ToAlphaString(i) + strconv.Itoa(j)
			xlsx.SetCellStr("Sheet1", axis, axis)
		}
	}
	xlsx.SetCellHyperLink("Sheet1", "A5", "https://github.com/360EntSecGroup-Skylar/excelize", "External")
	xlsx.SetCellHyperLink("Sheet1", "C5", "https://github.com", "External")
	xlsx.MergeCell("Sheet1", "A1", "B1")
	xlsx.MergeCell("Sheet1", "A2", "B2")
	xlsx.RemoveCol("Sheet1", "A")
	xlsx.RemoveCol("Sheet1", "A")
	err := xlsx.SaveAs("./test/Book_removecol.xlsx")
	if err != nil {
		t.Error(err)
	}
}

func TestInsertRow(t *testing.T) {
	xlsx := NewFile()
	for j := 1; j <= 10; j++ {
		for i := 0; i <= 10; i++ {
			axis := ToAlphaString(i) + strconv.Itoa(j)
			xlsx.SetCellStr("Sheet1", axis, axis)
		}
	}
	xlsx.SetCellHyperLink("Sheet1", "A5", "https://github.com/360EntSecGroup-Skylar/excelize", "External")
	xlsx.InsertRow("Sheet1", -1)
	xlsx.InsertRow("Sheet1", 4)
	err := xlsx.SaveAs("./test/Book_insertrow.xlsx")
	if err != nil {
		t.Error(err)
	}
}

func TestSetPane(t *testing.T) {
	xlsx := NewFile()
	xlsx.SetPanes("Sheet1", `{"freeze":false,"split":false}`)
	xlsx.NewSheet("Panes 2")
	xlsx.SetPanes("Panes 2", `{"freeze":true,"split":false,"x_split":1,"y_split":0,"top_left_cell":"B1","active_pane":"topRight","panes":[{"sqref":"K16","active_cell":"K16","pane":"topRight"}]}`)
	xlsx.NewSheet("Panes 3")
	xlsx.SetPanes("Panes 3", `{"freeze":false,"split":true,"x_split":3270,"y_split":1800,"top_left_cell":"N57","active_pane":"bottomLeft","panes":[{"sqref":"I36","active_cell":"I36"},{"sqref":"G33","active_cell":"G33","pane":"topRight"},{"sqref":"J60","active_cell":"J60","pane":"bottomLeft"},{"sqref":"O60","active_cell":"O60","pane":"bottomRight"}]}`)
	xlsx.NewSheet("Panes 4")
	xlsx.SetPanes("Panes 4", `{"freeze":true,"split":false,"x_split":0,"y_split":9,"top_left_cell":"A34","active_pane":"bottomLeft","panes":[{"sqref":"A11:XFD11","active_cell":"A11","pane":"bottomLeft"}]}`)
	xlsx.SetPanes("Panes 4", "")
	err := xlsx.SaveAs("./test/Book_set_panes.xlsx")
	if err != nil {
		t.Error(err)
	}
}

func TestRemoveRow(t *testing.T) {
	xlsx := NewFile()
	for j := 1; j <= 10; j++ {
		for i := 0; i <= 10; i++ {
			axis := ToAlphaString(i) + strconv.Itoa(j)
			xlsx.SetCellStr("Sheet1", axis, axis)
		}
	}
	xlsx.SetCellHyperLink("Sheet1", "A5", "https://github.com/360EntSecGroup-Skylar/excelize", "External")
	xlsx.RemoveRow("Sheet1", -1)
	xlsx.RemoveRow("Sheet1", 4)
	xlsx.MergeCell("Sheet1", "B3", "B5")
	xlsx.RemoveRow("Sheet1", 2)
	xlsx.RemoveRow("Sheet1", 4)
	err := xlsx.AutoFilter("Sheet1", "A2", "A2", `{"column":"A","expression":"x != blanks"}`)
	t.Log(err)
	xlsx.RemoveRow("Sheet1", 0)
	xlsx.RemoveRow("Sheet1", 1)
	xlsx.RemoveRow("Sheet1", 0)
	err = xlsx.SaveAs("./test/Book_removerow.xlsx")
	if err != nil {
		t.Error(err)
	}
}

func TestConditionalFormat(t *testing.T) {
	xlsx := NewFile()
	for j := 1; j <= 10; j++ {
		for i := 0; i <= 15; i++ {
			xlsx.SetCellInt("Sheet1", ToAlphaString(i)+strconv.Itoa(j), j)
		}
	}
	var format1, format2, format3 int
	var err error
	// Rose format for bad conditional.
	format1, err = xlsx.NewConditionalStyle(`{"font":{"color":"#9A0511"},"fill":{"type":"pattern","color":["#FEC7CE"],"pattern":1}}`)
	t.Log(err)
	// Light yellow format for neutral conditional.
	format2, err = xlsx.NewConditionalStyle(`{"fill":{"type":"pattern","color":["#FEEAA0"],"pattern":1}}`)
	t.Log(err)
	// Light green format for good conditional.
	format3, err = xlsx.NewConditionalStyle(`{"font":{"color":"#09600B"},"fill":{"type":"pattern","color":["#C7EECF"],"pattern":1}}`)
	t.Log(err)
	// Color scales: 2 color.
	xlsx.SetConditionalFormat("Sheet1", "A1:A10", `[{"type":"2_color_scale","criteria":"=","min_type":"min","max_type":"max","min_color":"#F8696B","max_color":"#63BE7B"}]`)
	// Color scales: 3 color.
	xlsx.SetConditionalFormat("Sheet1", "B1:B10", `[{"type":"3_color_scale","criteria":"=","min_type":"min","mid_type":"percentile","max_type":"max","min_color":"#F8696B","mid_color":"#FFEB84","max_color":"#63BE7B"}]`)
	// Hightlight cells rules: between...
	xlsx.SetConditionalFormat("Sheet1", "C1:C10", fmt.Sprintf(`[{"type":"cell","criteria":"between","format":%d,"minimum":"6","maximum":"8"}]`, format1))
	// Hightlight cells rules: Greater Than...
	xlsx.SetConditionalFormat("Sheet1", "D1:D10", fmt.Sprintf(`[{"type":"cell","criteria":">","format":%d,"value":"6"}]`, format3))
	// Hightlight cells rules: Equal To...
	xlsx.SetConditionalFormat("Sheet1", "E1:E10", fmt.Sprintf(`[{"type":"top","criteria":"=","format":%d}]`, format3))
	// Hightlight cells rules: Not Equal To...
	xlsx.SetConditionalFormat("Sheet1", "F1:F10", fmt.Sprintf(`[{"type":"unique","criteria":"=","format":%d}]`, format2))
	// Hightlight cells rules: Duplicate Values...
	xlsx.SetConditionalFormat("Sheet1", "G1:G10", fmt.Sprintf(`[{"type":"duplicate","criteria":"=","format":%d}]`, format2))
	// Top/Bottom rules: Top 10%.
	xlsx.SetConditionalFormat("Sheet1", "H1:H10", fmt.Sprintf(`[{"type":"top","criteria":"=","format":%d,"value":"6","percent":true}]`, format1))
	// Top/Bottom rules: Above Average...
	xlsx.SetConditionalFormat("Sheet1", "I1:I10", fmt.Sprintf(`[{"type":"average","criteria":"=","format":%d, "above_average": true}]`, format3))
	// Top/Bottom rules: Below Average...
	xlsx.SetConditionalFormat("Sheet1", "J1:J10", fmt.Sprintf(`[{"type":"average","criteria":"=","format":%d, "above_average": false}]`, format1))
	// Data Bars: Gradient Fill.
	xlsx.SetConditionalFormat("Sheet1", "K1:K10", `[{"type":"data_bar", "criteria":"=", "min_type":"min","max_type":"max","bar_color":"#638EC6"}]`)
	// Use a formula to determine which cells to format.
	xlsx.SetConditionalFormat("Sheet1", "L1:L10", fmt.Sprintf(`[{"type":"formula", "criteria":"L2<3", "format":%d}]`, format1))
	// Test set invalid format set in conditional format
	xlsx.SetConditionalFormat("Sheet1", "L1:L10", "")
	err = xlsx.SaveAs("./test/Book_conditional_format.xlsx")
	if err != nil {
		t.Log(err)
	}

	// Set conditional format with illegal JSON string.
	_, err = xlsx.NewConditionalStyle("")
	t.Log(err)
	// Set conditional format with illegal valid type.
	xlsx.SetConditionalFormat("Sheet1", "K1:K10", `[{"type":"", "criteria":"=", "min_type":"min","max_type":"max","bar_color":"#638EC6"}]`)
	// Set conditional format with illegal criteria type.
	xlsx.SetConditionalFormat("Sheet1", "K1:K10", `[{"type":"data_bar", "criteria":"", "min_type":"min","max_type":"max","bar_color":"#638EC6"}]`)
	// Set conditional format with file without dxfs element.
	xlsx, err = OpenFile("./test/Book1.xlsx")
	t.Log(err)
	_, err = xlsx.NewConditionalStyle(`{"font":{"color":"#9A0511"},"fill":{"type":"pattern","color":["#FEC7CE"],"pattern":1}}`)
	t.Log(err)
}

func TestTitleToNumber(t *testing.T) {
	if TitleToNumber("AK") != 36 {
		t.Error("Conver title to number failed")
	}
	if TitleToNumber("ak") != 36 {
		t.Error("Conver title to number failed")
	}
}

func TestSharedStrings(t *testing.T) {
	xlsx, err := OpenFile("./test/SharedStrings.xlsx")
	if err != nil {
		t.Error(err)
		return
	}
	xlsx.GetRows("Sheet1")
}

func TestSetSheetRow(t *testing.T) {
	xlsx, err := OpenFile("./test/Book1.xlsx")
	if err != nil {
		t.Error(err)
		return
	}
	xlsx.SetSheetRow("Sheet1", "B27", &[]interface{}{"cell", nil, int32(42), float64(42), time.Now()})
	xlsx.SetSheetRow("Sheet1", "", &[]interface{}{"cell", nil, 2})
	xlsx.SetSheetRow("Sheet1", "B27", []interface{}{})
	xlsx.SetSheetRow("Sheet1", "B27", &xlsx)
	err = xlsx.Save()
	if err != nil {
		t.Error(err)
		return
	}
}

func TestRows(t *testing.T) {
	xlsx, err := OpenFile("./test/Book1.xlsx")
	if err != nil {
		t.Error(err)
		return
	}
	rows, err := xlsx.Rows("Sheet2")
	if err != nil {
		t.Error(err)
		return
	}
	rowStrs := make([][]string, 0)
	var i = 0
	for rows.Next() {
		i++
		columns := rows.Columns()
		rowStrs = append(rowStrs, columns)
	}
	if rows.Error() != nil {
		t.Error(rows.Error())
		return
	}
	dstRows := xlsx.GetRows("Sheet2")
	if len(dstRows) != len(rowStrs) {
		t.Error("values not equal")
		return
	}
	for i := 0; i < len(rowStrs); i++ {
		if !reflect.DeepEqual(trimSliceSpace(dstRows[i]), trimSliceSpace(rowStrs[i])) {
			t.Error("values not equal")
			return
		}
	}
	rows, err = xlsx.Rows("SheetN")
	if err != nil {
		t.Log(err)
	}
}

func TestOutlineLevel(t *testing.T) {
	xlsx := NewFile()
	xlsx.NewSheet("Sheet2")
	xlsx.SetColOutlineLevel("Sheet1", "D", 4)
	xlsx.GetColOutlineLevel("Sheet1", "D")
	xlsx.GetColOutlineLevel("Shee2", "A")
	xlsx.SetColWidth("Sheet2", "A", "D", 13)
	xlsx.SetColOutlineLevel("Sheet2", "B", 2)
	xlsx.SetRowOutlineLevel("Sheet1", 2, 1)
	xlsx.GetRowOutlineLevel("Sheet1", 2)
	err := xlsx.SaveAs("./test/Book_outline_level.xlsx")
	if err != nil {
		t.Error(err)
		return
	}
	xlsx, err = OpenFile("./test/Book1.xlsx")
	if err != nil {
		t.Error(err)
		return
	}
	xlsx.SetColOutlineLevel("Sheet2", "B", 2)
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
