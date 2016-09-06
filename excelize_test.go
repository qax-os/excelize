package excelize

import (
	"strconv"
	"testing"
)

func TestExcelize(t *testing.T) {
	// Test update a XLSX file
	file, err := OpenFile("./test/Workbook1.xlsx")
	if err != nil {
		t.Log(err)
	}
	file.UpdateLinkedValue()
	file.SetCellInt("SHEET2", "B2", 100)
	file.SetCellStr("SHEET2", "C11", "Knowns")
	file.NewSheet(3, "TestSheet")
	file.SetCellInt("Sheet3", "A23", 10)
	file.SetCellStr("SHEET3", "b230", "10")
	file.SetCellStr("SHEET10", "b230", "10")
	file.SetActiveSheet(2)
	// Test read cell value with given illegal rows number
	file.GetCellValue("Sheet2", "a-1")
	// Test read cell value with given lowercase column number
	file.GetCellValue("Sheet2", "a5")
	file.GetCellValue("Sheet2", "C11")
	file.GetCellValue("Sheet2", "D11")
	file.GetCellValue("Sheet2", "D12")
	// Test read cell value with given axis large than exists row
	file.GetCellValue("Sheet2", "E13")

	for i := 1; i <= 300; i++ {
		file.SetCellStr("SHEET3", "c"+strconv.Itoa(i), strconv.Itoa(i))
	}
	err = file.Save()
	if err != nil {
		t.Log(err)
	}
	// Test write file to given path
	err = file.WriteTo("./test/Workbook_2.xlsx")
	if err != nil {
		t.Log(err)
	}
	// Test write file to not exist directory
	err = file.WriteTo("")
	if err != nil {
		t.Log(err)
	}

	// Test create a XLSX file
	file2 := CreateFile()
	file2.NewSheet(2, "XLSXSheet2")
	file2.NewSheet(3, "XLSXSheet3")
	file2.SetCellInt("Sheet2", "A23", 56)
	file2.SetCellStr("SHEET1", "B20", "42")
	file2.SetActiveSheet(0)
	err = file2.WriteTo("./test/Workbook_3.xlsx")
	if err != nil {
		t.Log(err)
	}

	// Test open a XLSX file with given illegal path
	_, err = OpenFile("./test/Workbook.xlsx")
	if err != nil {
		t.Log(err)
	}
}
