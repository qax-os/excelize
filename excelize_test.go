package excelize

import (
	"strconv"
	"testing"
)

func TestExcelize(t *testing.T) {
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
	f1.SetCellInt("SHEET2", "A1", 100)
	f1.SetCellStr("SHEET2", "C11", "Knowns")
	f1.NewSheet(3, "TestSheet")
	f1.SetCellInt("Sheet3", "A23", 10)
	f1.SetCellStr("SHEET3", "b230", "10")
	f1.SetCellStr("SHEET10", "b230", "10")
	f1.SetActiveSheet(2)
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
	f1.SetCellValue("Sheet2", "F2", float32(42))
	f1.SetCellValue("Sheet2", "F2", float64(42))
	f1.SetCellValue("Sheet2", "G2", nil)
	// Test completion column.
	f1.SetCellValue("Sheet2", "M2", nil)
	// Test read cell value with given axis large than exists row.
	f1.GetCellValue("Sheet2", "E231")

	for i := 1; i <= 300; i++ {
		f1.SetCellStr("SHEET3", "c"+strconv.Itoa(i), strconv.Itoa(i))
	}
	err = f1.Save()
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

	// Test create a XLSX file.
	f3 := CreateFile()
	f3.NewSheet(2, "XLSXSheet2")
	f3.NewSheet(3, "XLSXSheet3")
	f3.SetCellInt("Sheet2", "A23", 56)
	f3.SetCellStr("SHEET1", "B20", "42")
	f3.SetActiveSheet(0)
	err = f3.WriteTo("./test/Workbook_3.xlsx")
	if err != nil {
		t.Log(err)
	}

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
