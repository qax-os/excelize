package excelize

import (
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"strconv"
	"testing"
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
	xlsx.SetCellValue("Sheet2", "F1", "Hello")
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
	// Test completion column.
	xlsx.SetCellValue("Sheet2", "M2", nil)
	// Test read cell value with given axis large than exists row.
	xlsx.GetCellValue("Sheet2", "E231")
	// Test get active sheet of XLSX and get sheet name of XLSX by given sheet index.
	xlsx.GetSheetName(xlsx.GetActiveSheetIndex())
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
	err = xlsx.WriteTo("")
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
	err = xlsx.WriteTo("./test/Workbook_2.xlsx")
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
	err = xlsx.WriteTo("./test/Workbook_3.xlsx")
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

func TestCreateFile(t *testing.T) {
	// Test create a XLSX file.
	xlsx := CreateFile()
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
	err = xlsx.WriteTo("./test/Workbook_3.xlsx")
	if err != nil {
		t.Log(err)
	}
}

func TestSetColWidth(t *testing.T) {
	xlsx := CreateFile()
	xlsx.SetColWidth("sheet1", "B", "A", 12)
	xlsx.SetColWidth("sheet1", "A", "B", 12)
	err := xlsx.WriteTo("./test/Workbook_4.xlsx")
	if err != nil {
		t.Log(err)
	}
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

func TestSetRowHeight(t *testing.T) {
	xlsx := CreateFile()
	xlsx.SetRowHeight("Sheet1", 0, 50)
	xlsx.SetRowHeight("Sheet1", 3, 90)
	err := xlsx.WriteTo("./test/Workbook_5.xlsx")
	if err != nil {
		t.Log(err)
	}
}

func TestSetBorder(t *testing.T) {
	xlsx, err := OpenFile("./test/Workbook_2.xlsx")
	if err != nil {
		t.Log(err)
	}
	// Test set border with invalid style parameter.
	err = xlsx.SetBorder("Sheet1", "J21", "L25", "")
	if err != nil {
		t.Log(err)
	}
	// Test set border with invalid style index number.
	err = xlsx.SetBorder("Sheet1", "J21", "L25", `{"border":[{"type":"left","color":"0000FF","style":-1},{"type":"top","color":"00FF00","style":14},{"type":"bottom","color":"FFFF00","style":5},{"type":"right","color":"FF0000","style":6},{"type":"diagonalDown","color":"A020F0","style":9},{"type":"diagonalUp","color":"A020F0","style":8}]}`)
	if err != nil {
		t.Log(err)
	}
	if err != nil {
		t.Log(err)
	}
	// Test set border on overlapping area.
	err = xlsx.SetBorder("Sheet1", "J21", "L25", `{"border":[{"type":"left","color":"0000FF","style":2},{"type":"top","color":"00FF00","style":12},{"type":"bottom","color":"FFFF00","style":5},{"type":"right","color":"FF0000","style":6},{"type":"diagonalDown","color":"A020F0","style":9},{"type":"diagonalUp","color":"A020F0","style":8}]}`)
	if err != nil {
		t.Log(err)
	}
	err = xlsx.SetBorder("Sheet1", "M28", "K24", `{"border":[{"type":"left","color":"0000FF","style":2},{"type":"top","color":"00FF00","style":3},{"type":"bottom","color":"FFFF00","style":4},{"type":"right","color":"FF0000","style":5},{"type":"diagonalDown","color":"A020F0","style":6},{"type":"diagonalUp","color":"A020F0","style":7}]}`)
	if err != nil {
		t.Log(err)
	}
	// Test set border for a single cell.
	err = xlsx.SetBorder("Sheet1", "O22", "O22", `{"border":[{"type":"left","color":"0000FF","style":8},{"type":"top","color":"00FF00","style":9},{"type":"bottom","color":"FFFF00","style":10},{"type":"right","color":"FF0000","style":11},{"type":"diagonalDown","color":"A020F0","style":12},{"type":"diagonalUp","color":"A020F0","style":13}]}`)
	if err != nil {
		t.Log(err)
	}
	err = xlsx.Save()
	if err != nil {
		t.Log(err)
	}
}
