package excelize

import (
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"io/ioutil"
	"strconv"
	"testing"
	"time"
)

func TestOpenFile(t *testing.T) {
	// Test update a XLSX file.
	xlsx, err := OpenFile("./test/Workbook1.xlsx")
	if err != nil {
		t.Log(err)
	}
	// Test get all the rows in a not exists sheet.
	rows := xlsx.GetRows("Sheet4")
	// Test get all the rows in a sheet.
	rows = xlsx.GetRows("Sheet2")
	for _, row := range rows {
		for _, cell := range row {
			t.Log(cell, "\t")
		}
		t.Log("\r\n")
	}
	xlsx.UpdateLinkedValue()
	xlsx.SetCellDefault("SHEET2", "A1", strconv.FormatFloat(float64(100.1588), 'f', -1, 32))
	xlsx.SetCellDefault("SHEET2", "A1", strconv.FormatFloat(float64(-100.1588), 'f', -1, 64))
	xlsx.SetCellInt("SHEET2", "A1", 100)
	xlsx.SetCellStr("SHEET2", "C11", "Knowns")
	// Test max characters in a cell.
	var s = "c"
	for i := 0; i < 32768; i++ {
		s += "c"
	}
	xlsx.SetCellStr("SHEET2", "D11", s)
	xlsx.NewSheet(3, ":\\/?*[]Maximum 31 characters allowed in sheet title.")
	// Test set sheet name with illegal name.
	xlsx.SetSheetName("Maximum 31 characters allowed i", "[Rename]:\\/?* Maximum 31 characters allowed in sheet title.")
	xlsx.SetCellInt("Sheet3", "A23", 10)
	xlsx.SetCellStr("SHEET3", "b230", "10")
	xlsx.SetCellStr("SHEET10", "b230", "10")
	xlsx.SetActiveSheet(2)
	xlsx.GetCellFormula("Sheet1", "B19") // Test get cell formula with given rows number.
	xlsx.GetCellFormula("Sheet2", "B20") // Test get cell formula with illegal sheet index.
	xlsx.GetCellFormula("Sheet1", "B20") // Test get cell formula with illegal rows number.
	// Test read cell value with given illegal rows number.
	xlsx.GetCellValue("Sheet2", "a-1")
	// Test read cell value with given lowercase column number.
	xlsx.GetCellValue("Sheet2", "a5")
	xlsx.GetCellValue("Sheet2", "C11")
	xlsx.GetCellValue("Sheet2", "D11")
	xlsx.GetCellValue("Sheet2", "D12")
	// Test SetCellValue function.
	xlsx.SetCellValue("Sheet2", "F1", " Hello")
	xlsx.SetCellValue("Sheet2", "G1", []byte("World"))
	xlsx.SetCellValue("Sheet2", "F2", 42)
	xlsx.SetCellValue("Sheet2", "F2", int8(42))
	xlsx.SetCellValue("Sheet2", "F2", int16(42))
	xlsx.SetCellValue("Sheet2", "F2", int32(42))
	xlsx.SetCellValue("Sheet2", "F2", int64(42))
	xlsx.SetCellValue("Sheet2", "F2", float32(42.65418))
	xlsx.SetCellValue("Sheet2", "F2", float64(-42.65418))
	xlsx.SetCellValue("Sheet2", "F2", float32(42))
	xlsx.SetCellValue("Sheet2", "F2", float64(42))
	xlsx.SetCellValue("Sheet2", "G2", nil)
	xlsx.SetCellValue("Sheet2", "G3", uint8(8))
	xlsx.SetCellValue("Sheet2", "G4", time.Now())
	// Test completion column.
	xlsx.SetCellValue("Sheet2", "M2", nil)
	// Test read cell value with given axis large than exists row.
	xlsx.GetCellValue("Sheet2", "E231")
	// Test get active sheet of XLSX and get sheet name of XLSX by given sheet index.
	xlsx.GetSheetName(xlsx.GetActiveSheetIndex())
	// Test get sheet index of XLSX by given worksheet name.
	xlsx.GetSheetIndex("Sheet1")
	// Test get sheet name of XLSX by given invalid sheet index.
	xlsx.GetSheetName(4)
	// Test get sheet map of XLSX.
	xlsx.GetSheetMap()
	for i := 1; i <= 300; i++ {
		xlsx.SetCellStr("SHEET3", "c"+strconv.Itoa(i), strconv.Itoa(i))
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
	xlsx, err := OpenFile("./test/Workbook1.xlsx")
	if err != nil {
		t.Log(err)
	}
	// Test add picture to sheet.
	err = xlsx.AddPicture("Sheet2", "I9", "./test/images/excel.jpg", `{"x_offset": 140, "y_offset": 120}`)
	if err != nil {
		t.Log(err)
	}
	// Test add picture to sheet with offset.
	err = xlsx.AddPicture("Sheet1", "F21", "./test/images/excel.png", `{"x_offset": 10, "y_offset": 10}`)
	if err != nil {
		t.Log(err)
	}
	// Test add picture to sheet with invalid file path.
	err = xlsx.AddPicture("Sheet1", "G21", "./test/images/excel.icon", "")
	if err != nil {
		t.Log(err)
	}
	// Test add picture to sheet with unsupport file type.
	err = xlsx.AddPicture("Sheet1", "G21", "./test/Workbook1.xlsx", "")
	if err != nil {
		t.Log(err)
	}
	// Test write file to given path.
	err = xlsx.SaveAs("./test/Workbook_2.xlsx")
	if err != nil {
		t.Log(err)
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
	err = xlsx.SaveAs("./test/Workbook_3.xlsx")
	if err != nil {
		t.Log(err)
	}

	// Test set active sheet without BookViews and Sheets maps in xl/workbook.xml.
	f3, err := OpenFile("./test/badWorkbook.xlsx")
	f3.SetActiveSheet(2)
	if err != nil {
		t.Log(err)
	}

	// Test open a XLSX file with given illegal path.
	_, err = OpenFile("./test/Workbook.xlsx")
	if err != nil {
		t.Log(err)
	}
}

func TestNewFile(t *testing.T) {
	// Test create a XLSX file.
	xlsx := NewFile()
	xlsx.NewSheet(2, "XLSXSheet2")
	xlsx.NewSheet(3, "XLSXSheet3")
	xlsx.SetCellInt("Sheet2", "A23", 56)
	xlsx.SetCellStr("SHEET1", "B20", "42")
	xlsx.SetActiveSheet(0)
	// Test add picture to sheet with scaling.
	err := xlsx.AddPicture("Sheet1", "H2", "./test/images/excel.gif", `{"x_scale": 0.5, "y_scale": 0.5}`)
	if err != nil {
		t.Log(err)
	}
	err = xlsx.AddPicture("Sheet1", "C2", "./test/images/excel.png", "")
	if err != nil {
		t.Log(err)
	}
	err = xlsx.SaveAs("./test/Workbook_3.xlsx")
	if err != nil {
		t.Log(err)
	}
}

func TestColWidth(t *testing.T) {
	xlsx := NewFile()
	xlsx.SetColWidth("sheet1", "B", "A", 12)
	xlsx.SetColWidth("sheet1", "A", "B", 12)
	xlsx.GetColWidth("sheet1", "A")
	xlsx.GetColWidth("sheet1", "C")
	err := xlsx.SaveAs("./test/Workbook_4.xlsx")
	if err != nil {
		t.Log(err)
	}
	convertRowHeightToPixels(0)
}

func TestRowHeight(t *testing.T) {
	xlsx := NewFile()
	xlsx.SetRowHeight("Sheet1", 0, 50)
	xlsx.SetRowHeight("Sheet1", 3, 90)
	t.Log(xlsx.GetRowHeight("Sheet1", 1))
	t.Log(xlsx.GetRowHeight("Sheet1", 3))
	err := xlsx.SaveAs("./test/Workbook_5.xlsx")
	if err != nil {
		t.Log(err)
	}
	convertColWidthToPixels(0)
}

func TestSetCellHyperLink(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook1.xlsx")
	if err != nil {
		t.Log(err)
	}
	// Test set cell hyperlink in a work sheet already have hyperlinks.
	xlsx.SetCellHyperLink("sheet1", "B19", "https://github.com/Luxurioust/excelize")
	// Test add first hyperlink in a work sheet.
	xlsx.SetCellHyperLink("sheet2", "C1", "https://github.com/Luxurioust/excelize")
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
}

func TestSetCellFormula(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook1.xlsx")
	if err != nil {
		t.Log(err)
	}
	xlsx.SetCellFormula("sheet1", "B19", "SUM(Sheet2!D2,Sheet2!D11)")
	xlsx.SetCellFormula("sheet1", "C19", "SUM(Sheet2!D2,Sheet2!D9)")
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
}

func TestSetSheetBackground(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook1.xlsx")
	if err != nil {
		t.Log(err)
	}
	err = xlsx.SetSheetBackground("sheet2", "./test/images/background.png")
	if err != nil {
		t.Log(err)
	}
	err = xlsx.SetSheetBackground("sheet2", "./test/Workbook1.xlsx")
	if err != nil {
		t.Log(err)
	}
	err = xlsx.SetSheetBackground("sheet2", "./test/images/background.jpg")
	if err != nil {
		t.Log(err)
	}
	err = xlsx.SetSheetBackground("sheet2", "./test/images/background.jpg")
	if err != nil {
		t.Log(err)
	}
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
}

func TestMergeCell(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook1.xlsx")
	if err != nil {
		t.Log(err)
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
	xlsx.SetCellHyperLink("Sheet1", "J11", "https://github.com/Luxurioust/excelize")
	xlsx.SetCellFormula("Sheet1", "G12", "SUM(Sheet1!B19,Sheet1!C19)")
	xlsx.GetCellValue("Sheet1", "H11")
	xlsx.GetCellFormula("Sheet1", "G12")
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
}

func TestSetCellStyleAlignment(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook_2.xlsx")
	if err != nil {
		t.Log(err)
	}
	err = xlsx.SetCellStyle("Sheet1", "A22", "A22", `{"alignment":{"horizontal":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":45,"vertical":"top","wrap_text":true}}`)
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
}

func TestSetCellStyleBorder(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook_2.xlsx")
	if err != nil {
		t.Log(err)
	}
	// Test set border with invalid style parameter.
	err = xlsx.SetCellStyle("Sheet1", "J21", "L25", "")
	if err != nil {
		t.Log(err)
	}
	// Test set border with invalid style index number.
	err = xlsx.SetCellStyle("Sheet1", "J21", "L25", `{"border":[{"type":"left","color":"0000FF","style":-1},{"type":"top","color":"00FF00","style":14},{"type":"bottom","color":"FFFF00","style":5},{"type":"right","color":"FF0000","style":6},{"type":"diagonalDown","color":"A020F0","style":9},{"type":"diagonalUp","color":"A020F0","style":8}]}`)
	if err != nil {
		t.Log(err)
	}
	// Test set border on overlapping area with vertical variants shading styles gradient fill.
	err = xlsx.SetCellStyle("Sheet1", "J21", "L25", `{"border":[{"type":"left","color":"0000FF","style":2},{"type":"top","color":"00FF00","style":12},{"type":"bottom","color":"FFFF00","style":5},{"type":"right","color":"FF0000","style":6},{"type":"diagonalDown","color":"A020F0","style":9},{"type":"diagonalUp","color":"A020F0","style":8}]}`)
	if err != nil {
		t.Log(err)
	}
	err = xlsx.SetCellStyle("Sheet1", "M28", "K24", `{"border":[{"type":"left","color":"0000FF","style":2},{"type":"top","color":"00FF00","style":3},{"type":"bottom","color":"FFFF00","style":4},{"type":"right","color":"FF0000","style":5},{"type":"diagonalDown","color":"A020F0","style":6},{"type":"diagonalUp","color":"A020F0","style":7}],"fill":{"type":"gradient","color":["#FFFFFF","#E0EBF5"],"shading":1}}`)
	if err != nil {
		t.Log(err)
	}
	// Test set border and solid style pattern fill for a single cell.
	err = xlsx.SetCellStyle("Sheet1", "O22", "O22", `{"border":[{"type":"left","color":"0000FF","style":8},{"type":"top","color":"00FF00","style":9},{"type":"bottom","color":"FFFF00","style":10},{"type":"right","color":"FF0000","style":11},{"type":"diagonalDown","color":"A020F0","style":12},{"type":"diagonalUp","color":"A020F0","style":13}],"fill":{"type":"pattern","color":["#E0EBF5"],"pattern":1}}`)
	if err != nil {
		t.Log(err)
	}
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
}

func TestSetCellStyleNumberFormat(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook_2.xlsx")
	if err != nil {
		t.Log(err)
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
			err := xlsx.SetCellStyle("Sheet2", c, c, `{"fill":{"type":"gradient","color":["#FFFFFF","#E0EBF5"],"shading":5},"number_format": `+strconv.Itoa(d)+`}`)
			if err != nil {
				t.Log(err)
			}
			t.Log(xlsx.GetCellValue("Sheet2", c))
		}
	}
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
}

func TestSetCellStyleFill(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook_2.xlsx")
	if err != nil {
		t.Log(err)
	}
	// Test set fill for cell with invalid parameter.
	err = xlsx.SetCellStyle("Sheet1", "O23", "O23", `{"fill":{"type":"gradient","color":["#FFFFFF","#E0EBF5"],"shading":6}}`)
	if err != nil {
		t.Log(err)
	}
	err = xlsx.SetCellStyle("Sheet1", "O23", "O23", `{"fill":{"type":"gradient","color":["#FFFFFF"],"shading":1}}`)
	if err != nil {
		t.Log(err)
	}
	err = xlsx.SetCellStyle("Sheet1", "O23", "O23", `{"fill":{"type":"pattern","color":[],"pattern":1}}`)
	if err != nil {
		t.Log(err)
	}

	err = xlsx.SetCellStyle("Sheet1", "O23", "O23", `{"fill":{"type":"pattern","color":["#E0EBF5"],"pattern":19}}`)
	if err != nil {
		t.Log(err)
	}
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
}

func TestSetCellStyleFont(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook_2.xlsx")
	if err != nil {
		t.Log(err)
	}
	err = xlsx.SetCellStyle("Sheet2", "A1", "A1", `{"font":{"bold":true,"italic":true,"family":"Berlin Sans FB Demi","size":36,"color":"#777777","underline":"single"}}`)
	if err != nil {
		t.Log(err)
	}
	err = xlsx.SetCellStyle("Sheet2", "A2", "A2", `{"font":{"italic":true,"underline":"double"}}`)
	if err != nil {
		t.Log(err)
	}
	err = xlsx.SetCellStyle("Sheet2", "A3", "A3", `{"font":{"bold":true}}`)
	if err != nil {
		t.Log(err)
	}
	err = xlsx.SetCellStyle("Sheet2", "A4", "A4", `{"font":{"bold":true,"family":"","size":0,"color":"","underline":""}}`)
	if err != nil {
		t.Log(err)
	}
	err = xlsx.SetCellStyle("Sheet2", "A5", "A5", `{"font":{"color":"#777777"}}`)
	if err != nil {
		t.Log(err)
	}
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
}

func TestSetDeleteSheet(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook_3.xlsx")
	if err != nil {
		t.Log(err)
	}
	xlsx.DeleteSheet("XLSXSheet3")
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
	xlsx, err = OpenFile("./test/Workbook_4.xlsx")
	if err != nil {
		t.Log(err)
	}
	xlsx.DeleteSheet("Sheet1")
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
}

func TestGetPicture(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook_2.xlsx")
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
}

func TestSheetVisibility(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook_2.xlsx")
	if err != nil {
		t.Log(err)
	}
	xlsx.SetSheetVisible("Sheet2", false)
	xlsx.SetSheetVisible("Sheet1", false)
	xlsx.SetSheetVisible("Sheet1", true)
	xlsx.GetSheetVisible("Sheet1")
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
}

func TestRowVisibility(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook_2.xlsx")
	if err != nil {
		t.Log(err)
	}
	xlsx.SetRowVisible("Sheet3", 2, false)
	xlsx.SetRowVisible("Sheet3", 2, true)
	xlsx.GetRowVisible("Sheet3", 2)
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
}

func TestColumnVisibility(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook_2.xlsx")
	if err != nil {
		t.Log(err)
	}
	xlsx.SetColVisible("Sheet1", "F", false)
	xlsx.SetColVisible("Sheet1", "F", true)
	xlsx.GetColVisible("Sheet1", "F")
	xlsx.SetColVisible("Sheet3", "E", false)
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
	xlsx, err = OpenFile("./test/Workbook_3.xlsx")
	if err != nil {
		t.Log(err)
	}
	xlsx.GetColVisible("Sheet1", "B")
}

func TestCopySheet(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook_2.xlsx")
	if err != nil {
		t.Log(err)
	}
	err = xlsx.CopySheet(0, -1)
	if err != nil {
		t.Log(err)
	}
	xlsx.NewSheet(4, "CopySheet")
	err = xlsx.CopySheet(1, 4)
	if err != nil {
		t.Log(err)
	}
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
}

func TestAddTable(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook_2.xlsx")
	if err != nil {
		t.Log(err)
	}
	xlsx.AddTable("Sheet1", "B26", "A21", ``)
	xlsx.AddTable("Sheet2", "A2", "B5", `{"table_style":"TableStyleMedium2", "show_first_column":true,"show_last_column":true,"show_row_stripes":false,"show_column_stripes":true}`)
	xlsx.AddTable("Sheet2", "F1", "F1", `{"table_style":"TableStyleMedium8"}`)
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
}

func TestAddShape(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook_2.xlsx")
	if err != nil {
		t.Log(err)
	}
	xlsx.AddShape("Sheet1", "A30", `{"type":"rect","paragraph":[{"text":"Rectangle","font":{"color":"CD5C5C"}},{"text":"Shape","font":{"bold":true,"color":"2980B9"}}]}`)
	xlsx.AddShape("Sheet1", "B30", `{"type":"rect","paragraph":[{"text":"Rectangle"},{}]}`)
	xlsx.AddShape("Sheet1", "C30", `{"type":"rect","paragraph":[]}`)
	xlsx.AddShape("Sheet3", "H1", `{"type":"ellipseRibbon", "color":{"line":"#4286f4","fill":"#8eb9ff"}, "paragraph":[{"font":{"bold":true,"italic":true,"family":"Berlin Sans FB Demi","size":36,"color":"#777777","underline":"single"}}], "height": 90}`)
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
}

func TestAddComments(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook_2.xlsx")
	if err != nil {
		t.Log(err)
	}
	var s = "c"
	for i := 0; i < 32767; i++ {
		s += "c"
	}
	xlsx.AddComment("Sheet1", "A30", `{"author":"`+s+`","text":"`+s+`"}`)
	xlsx.AddComment("Sheet2", "B7", `{"author":"Excelize: ","text":"This is a comment."}`)
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
}

func TestAutoFilter(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook_2.xlsx")
	if err != nil {
		t.Log(err)
	}
	err = xlsx.AutoFilter("Sheet3", "D4", "B1", ``)
	t.Log(err)
	err = xlsx.AutoFilter("Sheet3", "D4", "B1", `{"column":"B","expression":"x != blanks"}`)
	t.Log(err)
	err = xlsx.AutoFilter("Sheet3", "D4", "B1", `{"column":"B","expression":"x == blanks"}`)
	t.Log(err)
	err = xlsx.AutoFilter("Sheet3", "D4", "B1", `{"column":"B","expression":"x != nonblanks"}`)
	t.Log(err)
	err = xlsx.AutoFilter("Sheet3", "D4", "B1", `{"column":"B","expression":"x == nonblanks"}`)
	t.Log(err)
	err = xlsx.AutoFilter("Sheet3", "D4", "B1", `{"column":"B","expression":"x <= 1 and x >= 2"}`)
	t.Log(err)
	err = xlsx.AutoFilter("Sheet3", "D4", "B1", `{"column":"B","expression":"x == 1 or x == 2"}`)
	t.Log(err)
	err = xlsx.AutoFilter("Sheet3", "D4", "B1", `{"column":"B","expression":"x == 1 or x == 2*"}`)
	t.Log(err)
	err = xlsx.AutoFilter("Sheet3", "D4", "B1", `{"column":"B","expression":"x <= 1 and x >= blanks"}`)
	t.Log(err)
	err = xlsx.AutoFilter("Sheet3", "D4", "B1", `{"column":"B","expression":"x -- y or x == *2*"}`)
	t.Log(err)
	err = xlsx.AutoFilter("Sheet3", "D4", "B1", `{"column":"B","expression":"x != y or x ? *2"}`)
	t.Log(err)
	err = xlsx.AutoFilter("Sheet3", "D4", "B1", `{"column":"B","expression":"x -- y o r x == *2"}`)
	t.Log(err)
	err = xlsx.AutoFilter("Sheet3", "D4", "B1", `{"column":"B","expression":"x -- y"}`)
	t.Log(err)
	err = xlsx.AutoFilter("Sheet3", "D4", "B1", `{"column":"A","expression":"x -- y"}`)
	t.Log(err)
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
}

func TestAddChart(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook1.xlsx")
	if err != nil {
		t.Log(err)
	}
	categories := map[string]string{"A30": "Small", "A31": "Normal", "A32": "Large", "B29": "Apple", "C29": "Orange", "D29": "Pear"}
	values := map[string]int{"B30": 2, "C30": 3, "D30": 3, "B31": 5, "C31": 2, "D31": 4, "B32": 6, "C32": 7, "D32": 8}
	for k, v := range categories {
		xlsx.SetCellValue("Sheet1", k, v)
	}
	for k, v := range values {
		xlsx.SetCellValue("Sheet1", k, v)
	}
	xlsx.AddChart("SHEET1", "P1", `{"type":"bar3D","series":[{"name":"=Sheet1!$A$30","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$30:$D$30"},{"name":"=Sheet1!$A$31","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$31:$D$31"},{"name":"=Sheet1!$A$32","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"bottom","show_legend_key":false},"title":{"name":"Fruit 3D Bar Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"zero"}`)
	xlsx.AddChart("SHEET1", "X1", `{"type":"bar","series":[{"name":"=Sheet1!$A$30","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$30:$D$30"},{"name":"=Sheet1!$A$31","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$31:$D$31"},{"name":"=Sheet1!$A$32","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"left","show_legend_key":false},"title":{"name":"Fruit Bar Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"zero"}`)
	xlsx.AddChart("SHEET1", "P16", `{"type":"doughnut","series":[{"name":"=Sheet1!$A$30","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$30:$D$30"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"right","show_legend_key":false},"title":{"name":"Fruit Doughnut Chart"},"plotarea":{"show_bubble_size":false,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":false,"show_val":false},"show_blanks_as":"zero"}`)
	xlsx.AddChart("SHEET1", "X16", `{"type":"line","series":[{"name":"=Sheet1!$A$30","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$30:$D$30"},{"name":"=Sheet1!$A$31","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$31:$D$31"},{"name":"=Sheet1!$A$32","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"top","show_legend_key":false},"title":{"name":"Fruit Line Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"zero"}`)
	xlsx.AddChart("SHEET1", "P30", `{"type":"pie3D","series":[{"name":"=Sheet1!$A$30","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$30:$D$30"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"bottom","show_legend_key":false},"title":{"name":"Fruit 3D Pie Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":false,"show_val":false},"show_blanks_as":"zero"}`)
	xlsx.AddChart("SHEET1", "X30", `{"type":"pie","series":[{"name":"=Sheet1!$A$30","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$30:$D$30"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"bottom","show_legend_key":false},"title":{"name":"Fruit Pie Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":false,"show_val":false},"show_blanks_as":"gap"}`)
	xlsx.AddChart("SHEET2", "P1", `{"type":"radar","series":[{"name":"=Sheet1!$A$30","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$30:$D$30"},{"name":"=Sheet1!$A$31","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$31:$D$31"},{"name":"=Sheet1!$A$32","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"top_right","show_legend_key":false},"title":{"name":"Fruit Radar Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"span"}`)
	xlsx.AddChart("SHEET2", "X1", `{"type":"scatter","series":[{"name":"=Sheet1!$A$30","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$30:$D$30"},{"name":"=Sheet1!$A$31","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$31:$D$31"},{"name":"=Sheet1!$A$32","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"bottom","show_legend_key":false},"title":{"name":"Fruit Scatter Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"zero"}`)
	// Save xlsx file by the given path.
	err = xlsx.SaveAs("./test/Workbook_6.xlsx")
	if err != nil {
		t.Log(err)
	}
}
