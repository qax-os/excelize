package excelize

import (
	"strconv"
	"testing"
)

func TestExcelize(t *testing.T) {
	// Test update a XLSX file
	file, err := OpenFile("./test/Workbook1.xlsx")
	if err != nil {
		t.Error(err)
	}
	file = SetCellInt(file, "SHEET2", "B2", 100)
	file = SetCellStr(file, "SHEET2", "C11", "Knowns")
	file = NewSheet(file, 3, "TestSheet")
	file = SetCellInt(file, "Sheet3", "A23", 10)
	file = SetCellStr(file, "SHEET3", "b230", "10")
	file = SetActiveSheet(file, 2)
	if err != nil {
		t.Error(err)
	}
	for i := 1; i <= 300; i++ {
		file = SetCellStr(file, "SHEET3", "c"+strconv.Itoa(i), strconv.Itoa(i))
	}
	err = Save(file, "./test/Workbook_2.xlsx")

	// Test create a XLSX file
	file2 := CreateFile()
	file2 = NewSheet(file2, 2, "TestSheet2")
	file2 = NewSheet(file2, 3, "TestSheet3")
	file2 = SetCellInt(file2, "Sheet2", "A23", 10)
	file2 = SetCellStr(file2, "SHEET1", "B20", "10")
	err = Save(file2, "./test/Workbook_3.xlsx")
	if err != nil {
		t.Error(err)
	}

	// Test read cell value
	file, err = OpenFile("./test/Workbook1.xlsx")
	if err != nil {
		t.Error(err)
	}
	GetCellValue(file, "Sheet2", "a5")
	GetCellValue(file, "Sheet2", "D11")
	GetCellValue(file, "Sheet2", "D12")
	GetCellValue(file, "Sheet2", "E12")
}
