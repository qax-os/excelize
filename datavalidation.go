package excelize

import (
	"fmt"
	"strings"
)

type DataValidationType int

// Data validation types
const (
	_DataValidationType = iota
	typeNone            //inline use
	DataValidationTypeCustom
	DataValidationTypeDate
	DataValidationTypeDecimal
	typeList //inline use
	DataValidationTypeTextLeng
	DataValidationTypeTime
	// DataValidationTypeWhole Integer
	DataValidationTypeWhole
)

const (
	// dataValidationFormulaStrLen 255 characters+ 2 quotes
	dataValidationFormulaStrLen = 257
	// dataValidationFormulaStrLenErr
	dataValidationFormulaStrLenErr = "data validation must be 0-255 characters"
)

type DataValidationErrorStyle int

// Data validation error styles
const (
	_ DataValidationErrorStyle = iota
	DataValidationErrorStyleStop
	DataValidationErrorStyleWarning
	DataValidationErrorStyleInformation
)

// Data validation error styles
const (
	styleStop        = "stop"
	styleWarning     = "warning"
	styleInformation = "information"
)

// DataValidationOperator operator enum
type DataValidationOperator int

// Data validation operators
const (
	_DataValidationOperator = iota
	DataValidationOperatorBetween
	DataValidationOperatorEqual
	DataValidationOperatorGreaterThan
	DataValidationOperatorGreaterThanOrEqual
	DataValidationOperatorLessThan
	DataValidationOperatorLessThanOrEqual
	DataValidationOperatorNotBetween
	DataValidationOperatorNotEqual
)

// NewDataValidation return data validation struct
func NewDataValidation(allowBlank bool) *DataValidation {
	return &DataValidation{
		AllowBlank:       convBoolToStr(allowBlank),
		ShowErrorMessage: convBoolToStr(false),
		ShowInputMessage: convBoolToStr(false),
	}
}

// SetError set error notice
func (dd *DataValidation) SetError(style DataValidationErrorStyle, title, msg string) {
	dd.Error = &msg
	dd.ErrorTitle = &title
	strStyle := styleStop
	switch style {
	case DataValidationErrorStyleStop:
		strStyle = styleStop
	case DataValidationErrorStyleWarning:
		strStyle = styleWarning
	case DataValidationErrorStyleInformation:
		strStyle = styleInformation

	}
	dd.ShowErrorMessage = convBoolToStr(true)
	dd.ErrorStyle = &strStyle
}

// SetInput set prompt notice
func (dd *DataValidation) SetInput(title, msg string) {
	dd.ShowInputMessage = convBoolToStr(true)
	dd.PromptTitle = &title
	dd.Prompt = &msg
}

// SetDropList data validation list
func (dd *DataValidation) SetDropList(keys []string) error {
	dd.Formula1 = "\"" + strings.Join(keys, ",") + "\""
	dd.Type = convDataValidationType(typeList)
	return nil
}

// SetDropList data validation range
func (dd *DataValidation) SetRange(f1, f2 int, t DataValidationType, o DataValidationOperator) error {
	formula1 := fmt.Sprintf("%d", f1)
	formula2 := fmt.Sprintf("%d", f2)
	if dataValidationFormulaStrLen < len(dd.Formula1) || dataValidationFormulaStrLen < len(dd.Formula2) {
		return fmt.Errorf(dataValidationFormulaStrLenErr)
	}
	/*switch o {
	case DataValidationOperatorBetween:
		if f1 > f2 {
			tmp := formula1
			formula1 = formula2
			formula2 = tmp
		}
	case DataValidationOperatorNotBetween:
		if f1 > f2 {
			tmp := formula1
			formula1 = formula2
			formula2 = tmp
		}
	}*/

	dd.Formula1 = formula1
	dd.Formula2 = formula2
	dd.Type = convDataValidationType(t)
	dd.Operator = convDataValidationOperatior(o)
	return nil
}

// SetDropList data validation range
func (dd *DataValidation) SetSqref(sqref string) {
	if dd.Sqref == "" {
		dd.Sqref = sqref
	} else {
		dd.Sqref = fmt.Sprintf("%s %s", dd.Sqref, sqref)
	}
}

// convBoolToStr  convert boolean to string , false to 0, true to 1
func convBoolToStr(bl bool) string {
	if bl {
		return "1"
	}
	return "0"
}

// convDataValidationType get excel data validation type
func convDataValidationType(t DataValidationType) string {
	typeMap := map[DataValidationType]string{
		typeNone:                   "none",
		DataValidationTypeCustom:   "custom",
		DataValidationTypeDate:     "date",
		DataValidationTypeDecimal:  "decimal",
		typeList:                   "list",
		DataValidationTypeTextLeng: "textLength",
		DataValidationTypeTime:     "time",
		DataValidationTypeWhole:    "whole",
	}

	return typeMap[t]

}

// convDataValidationOperatior get excel data validation operator
func convDataValidationOperatior(o DataValidationOperator) string {
	typeMap := map[DataValidationOperator]string{
		DataValidationOperatorBetween:            "between",
		DataValidationOperatorEqual:              "equal",
		DataValidationOperatorGreaterThan:        "greaterThan",
		DataValidationOperatorGreaterThanOrEqual: "greaterThanOrEqual",
		DataValidationOperatorLessThan:           "lessThan",
		DataValidationOperatorLessThanOrEqual:    "lessThanOrEqual",
		DataValidationOperatorNotBetween:         "notBetween",
		DataValidationOperatorNotEqual:           "notEqual",
	}

	return typeMap[o]

}

func (f *File) AddDataValidation(sheet string, dv *DataValidation) {
	xlsx := f.workSheetReader(sheet)
	if nil == xlsx.DataValidations {
		xlsx.DataValidations = new(xlsxDataValidations)
	}
	xlsx.DataValidations.DataValidation = append(xlsx.DataValidations.DataValidation, dv)
	xlsx.DataValidations.Count = len(xlsx.DataValidations.DataValidation)
}
