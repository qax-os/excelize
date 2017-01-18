package excelize

import (
	"strconv"
	"testing"
)

func TestOpenFile(t *testing.T) {
	// Test update a XLSX file.
	f1, err := OpenFile("./test/Workbook1.xlsx")
	if err != nil {
		t.Log(err)
	}
	// Test get all the rows in a not exists sheet.
	rows := f1.GetRows("Sheet4")
	// Test get all the rows in a sheet.
	rows = f1.GetRows("Sheet2")
	for _, row := range rows {
		for _, cell := range row {
			t.Log(cell, "\t")
		}
		t.Log("\r\n")
	}
	f1.UpdateLinkedValue()
	f1.SetCellDefault("SHEET2", "A1", strconv.FormatFloat(float64(100.1588), 'f', -1, 32))
	f1.SetCellDefault("SHEET2", "A1", strconv.FormatFloat(float64(-100.1588), 'f', -1, 64))
	f1.SetCellInt("SHEET2", "A1", 100)
	f1.SetCellStr("SHEET2", "C11", "Knowns")
	f1.NewSheet(3, ":\\/?*[]Maximum 31 characters allowed in sheet title.")
	// Test set sheet name with illegal name.
	f1.SetSheetName("Maximum 31 characters allowed i", "[Rename]:\\/?* Maximum 31 characters allowed in sheet title.")
	f1.SetCellInt("Sheet3", "A23", 10)
	f1.SetCellStr("SHEET3", "b230", "10")
	f1.SetCellStr("SHEET10", "b230", "10")
	f1.SetActiveSheet(2)
	f1.GetCellFormula("Sheet1", "B19") // Test get cell formula with given rows number.
	f1.GetCellFormula("Sheet2", "B20") // Test get cell formula with illegal sheet index.
	f1.GetCellFormula("Sheet1", "B20") // Test get cell formula with illegal rows number.
	// Test read cell value with given illegal rows number.
	f1.GetCellValue("Sheet2", "a-1")
	// Test read cell value with given lowercase column number.
	f1.GetCellValue("Sheet2", "a5")
	f1.GetCellValue("Sheet2", "C11")
	f1.GetCellValue("Sheet2", "D11")
	f1.GetCellValue("Sheet2", "D12")
	// Test SetCellValue function.
	f1.SetCellValue("Sheet2", "F1", "Hello")
	f1.SetCellValue("Sheet2", "G1", []byte("World"))
	f1.SetCellValue("Sheet2", "F2", 42)
	f1.SetCellValue("Sheet2", "F2", int8(42))
	f1.SetCellValue("Sheet2", "F2", int16(42))
	f1.SetCellValue("Sheet2", "F2", int32(42))
	f1.SetCellValue("Sheet2", "F2", int64(42))
	f1.SetCellValue("Sheet2", "F2", float32(42.65418))
	f1.SetCellValue("Sheet2", "F2", float64(-42.65418))
	f1.SetCellValue("Sheet2", "F2", float32(42))
	f1.SetCellValue("Sheet2", "F2", float64(42))
	f1.SetCellValue("Sheet2", "G2", nil)
	// Test completion column.
	f1.SetCellValue("Sheet2", "M2", nil)
	// Test read cell value with given axis large than exists row.
	f1.GetCellValue("Sheet2", "E231")
	// Test get active sheet of XLSX and get sheet name of XLSX by given sheet index.
	f1.GetSheetName(f1.GetActiveSheetIndex())
	// Test get sheet name of XLSX by given invalid sheet index.
	f1.GetSheetName(4)
	// Test get sheet map of XLSX.
	f1.GetSheetMap()

	for i := 1; i <= 300; i++ {
		f1.SetCellStr("SHEET3", "c"+strconv.Itoa(i), strconv.Itoa(i))
	}
	err = f1.Save()
	if err != nil {
		t.Log(err)
	}
	// Test add picture to sheet.
	err = f1.AddPicture("Sheet2", "I1", "L10", "./test/images/excel.jpg")
	if err != nil {
		t.Log(err)
	}
	err = f1.AddPicture("Sheet1", "F21", "G25", "./test/images/excel.png")
	if err != nil {
		t.Log(err)
	}
	err = f1.AddPicture("Sheet2", "L1", "O10", "./test/images/excel.bmp")
	if err != nil {
		t.Log(err)
	}
	err = f1.AddPicture("Sheet1", "G21", "H25", "./test/images/excel.ico")
	if err != nil {
		t.Log(err)
	}
	// Test add picture to sheet with unsupport file type.
	err = f1.AddPicture("Sheet1", "G21", "H25", "./test/images/excel.icon")
	if err != nil {
		t.Log(err)
	}
	// Test add picture to sheet with invalid file path.
	err = f1.AddPicture("Sheet1", "G21", "H25", "./test/Workbook1.xlsx")
	if err != nil {
		t.Log(err)
	}

	// Test write file to given path.
	err = f1.WriteTo("./test/Workbook_2.xlsx")
	if err != nil {
		t.Log(err)
	}
	// Test write file to not exist directory.
	err = f1.WriteTo("")
	if err != nil {
		t.Log(err)
	}

	// Test write file with broken file struct.
	f2 := File{}
	err = f2.Save()
	if err != nil {
		t.Log(err)
	}
	// Test write file with broken file struct with given path.
	err = f2.WriteTo("./test/Workbook_3.xlsx")
	if err != nil {
		t.Log(err)
	}
}

func TestCreateFile(t *testing.T) {
	// Test create a XLSX file.
	f3 := CreateFile()
	f3.NewSheet(2, "XLSXSheet2")
	f3.NewSheet(3, "XLSXSheet3")
	f3.SetCellInt("Sheet2", "A23", 56)
	f3.SetCellStr("SHEET1", "B20", "42")
	f3.SetActiveSheet(0)
	// Test add picture to sheet.
	err := f3.AddPicture("Sheet1", "H2", "K12", "./test/images/excel.gif")
	if err != nil {
		t.Log(err)
	}
	err = f3.AddPicture("Sheet1", "C2", "F12", "./test/images/excel.tif")
	if err != nil {
		t.Log(err)
	}
	err = f3.WriteTo("./test/Workbook_3.xlsx")
	if err != nil {
		t.Log(err)
	}
}

func TestBrokenFile(t *testing.T) {
	// Test set active sheet without BookViews and Sheets maps in xl/workbook.xml.
	f4, err := OpenFile("./test/badWorkbook.xlsx")
	f4.SetActiveSheet(2)
	if err != nil {
		t.Log(err)
	}

	// Test open a XLSX file with given illegal path.
	_, err = OpenFile("./test/Workbook.xlsx")
	if err != nil {
		t.Log(err)
	}
}

func TestSetColWidth(t *testing.T) {
	f5, err := OpenFile("./test/Workbook1.xlsx")
	if err != nil {
		t.Log(err)
	}
	f5.SetColWidth("sheet1", "B", "A", 12)
	f5.SetColWidth("sheet1", "A", "B", 12)
	err = f5.Save()
	if err != nil {
		t.Log(err)
	}
}
