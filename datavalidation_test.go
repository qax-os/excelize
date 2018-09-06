package excelize

import (
	"testing"
)

func TestDataValidation(t *testing.T) {
	xlsx := NewFile()

	dvRange := NewDataValidation(true)
	dvRange.Sqref = "A1:B2"
	dvRange.SetRange(10, 20, DataValidationTypeWhole, DataValidationOperatorBetween)
	dvRange.SetError(DataValidationErrorStyleStop, "error title", "error body")
	xlsx.AddDataValidation("Sheet1", dvRange)

	dvRange = NewDataValidation(true)
	dvRange.Sqref = "A3:B4"
	dvRange.SetRange(10, 20, DataValidationTypeWhole, DataValidationOperatorGreaterThan)
	dvRange.SetInput("input title", "input body")
	xlsx.AddDataValidation("Sheet1", dvRange)

	dvRange = NewDataValidation(true)
	dvRange.Sqref = "A5:B6"
	dvRange.SetDropList([]string{"1", "2", "3"})
	xlsx.AddDataValidation("Sheet1", dvRange)

	xlsx.SetCellStr("Sheet1", "E1", "E1")
	xlsx.SetCellStr("Sheet1", "E2", "E2")
	xlsx.SetCellStr("Sheet1", "E3", "E3")
	dvRange = NewDataValidation(true)
	dvRange.Sqref = "A7:B8"
	dvRange.SetSqrefDropList("$E$1:$E$3", true)
	xlsx.AddDataValidation("Sheet1", dvRange)

	// Test write file to given path.
	err := xlsx.SaveAs("./test/Book_data_validation.xlsx")
	if err != nil {
		t.Error(err)
	}
}
