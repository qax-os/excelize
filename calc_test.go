package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestCalcCellValue(t *testing.T) {
	prepareData := func() *File {
		f := NewFile()
		f.SetCellValue("Sheet1", "A1", 1)
		f.SetCellValue("Sheet1", "A2", 2)
		f.SetCellValue("Sheet1", "A3", 3)
		f.SetCellValue("Sheet1", "A4", 0)
		f.SetCellValue("Sheet1", "B1", 4)
		return f
	}

	mathCalc := map[string]string{
		// ABS
		"=ABS(-1)":    "1",
		"=ABS(-6.5)":  "6.5",
		"=ABS(6.5)":   "6.5",
		"=ABS(0)":     "0",
		"=ABS(2-4.5)": "2.5",
		// GCD
		"=GCD(1,5)":      "1",
		"=GCD(15,10,25)": "5",
		"=GCD(0,8,12)":   "4",
		"=GCD(7,2)":      "1",
		// LCM
		"=LCM(1,5)":      "5",
		"=LCM(15,10,25)": "150",
		"=LCM(1,8,12)":   "24",
		"=LCM(7,2)":      "14",
		// POWER
		"=POWER(4,2)": "16",
		// PRODUCT
		"=PRODUCT(3,6)": "18",
		// SIGN
		"=SIGN(9.5)":        "1",
		"=SIGN(-9.5)":       "-1",
		"=SIGN(0)":          "0",
		"=SIGN(0.00000001)": "1",
		"=SIGN(6-7)":        "-1",
		// SQRT
		"=SQRT(4)": "2",
		// SUM
		"=SUM(1,2)":                           "3",
		"=SUM(1,2+3)":                         "6",
		"=SUM(SUM(1,2),2)":                    "5",
		"=(-2-SUM(-4+7))*5":                   "-25",
		"SUM(1,2,3,4,5,6,7)":                  "28",
		"=SUM(1,2)+SUM(1,2)":                  "6",
		"=1+SUM(SUM(1,2*3),4)":                "12",
		"=1+SUM(SUM(1,-2*3),4)":               "0",
		"=(-2-SUM(-4*(7+7)))*5":               "270",
		"=SUM(SUM(1+2/1)*2-3/2,2)":            "6.5",
		"=((3+5*2)+3)/5+(-6)/4*2+3":           "3.2",
		"=1+SUM(SUM(1,2*3),4)*-4/2+5+(4+2)*3": "2",
		"=1+SUM(SUM(1,2*3),4)*4/3+5+(4+2)*3":  "38.666666666666664",
		// QUOTIENT
		"=QUOTIENT(5, 2)":     "2",
		"=QUOTIENT(4.5, 3.1)": "1",
		"=QUOTIENT(-10, 3)":   "-3",
	}
	for formula, expected := range mathCalc {
		f := prepareData()
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err)
		assert.Equal(t, expected, result)
	}
	mathCalcError := map[string]string{
		// ABS
		"=ABS(1,2)": "ABS requires 1 numeric arguments",
		"=ABS(~)":   `cannot convert cell "~" to coordinates: invalid cell name "~"`,
		// GCD
		"=GCD()":     "GCD requires at least 1 argument",
		"=GCD(-1)":   "GCD only accepts positive arguments",
		"=GCD(1,-1)": "GCD only accepts positive arguments",
		// LCM
		"=LCM()":     "LCM requires at least 1 argument",
		"=LCM(-1)":   "LCM only accepts positive arguments",
		"=LCM(1,-1)": "LCM only accepts positive arguments",
		// POWER
		"=POWER(0,0)":  "#NUM!",
		"=POWER(0,-1)": "#DIV/0!",
		"=POWER(1)":    "POWER requires 2 numeric arguments",
		// SIGN
		"=SIGN()": "SIGN requires 1 numeric arguments",
		// SQRT
		"=SQRT(-1)":  "#NUM!",
		"=SQRT(1,2)": "SQRT requires 1 numeric arguments",
		// QUOTIENT
		"=QUOTIENT(1,0)": "#DIV/0!",
		"=QUOTIENT(1)":   "QUOTIENT requires 2 numeric arguments",
	}
	for formula, expected := range mathCalcError {
		f := prepareData()
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.EqualError(t, err, expected)
		assert.Equal(t, "", result)
	}

	referenceCalc := map[string]string{
		// PRODUCT
		"=PRODUCT(Sheet1!A1:Sheet1!A1:A2,A2)": "4",
		// SUM
		"=A1/A3":                          "0.3333333333333333",
		"=SUM(A1:A2)":                     "3",
		"=SUM(Sheet1!A1,A2)":              "3",
		"=(-2-SUM(-4+A2))*5":              "0",
		"=SUM(Sheet1!A1:Sheet1!A1:A2,A2)": "5",
		"=SUM(A1,A2,A3)*SUM(2,3)":         "30",
		"=1+SUM(SUM(A1+A2/A3)*(2-3),2)":   "1.3333333333333335",
		"=A1/A2/SUM(A1:A2:B1)":            "0.07142857142857142",
		"=A1/A2/SUM(A1:A2:B1)*A3":         "0.21428571428571427",
	}
	for formula, expected := range referenceCalc {
		f := prepareData()
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err)
		assert.Equal(t, expected, result)
	}

	referenceCalcError := map[string]string{
		"=1+SUM(SUM(A1+A2/A4)*(2-3),2)": "#DIV/0!",
	}
	for formula, expected := range referenceCalcError {
		f := prepareData()
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.EqualError(t, err, expected)
		assert.Equal(t, "", result)
	}

	// Test get calculated cell value on not formula cell.
	f := prepareData()
	result, err := f.CalcCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "", result)
	// Test get calculated cell value on not exists worksheet.
	f = prepareData()
	_, err = f.CalcCellValue("SheetN", "A1")
	assert.EqualError(t, err, "sheet SheetN is not exist")
	// Test get calculated cell value with not support formula.
	f = prepareData()
	assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "=UNSUPPORT(A1)"))
	_, err = f.CalcCellValue("Sheet1", "A1")
	assert.EqualError(t, err, "not support UNSUPPORT function")
}
