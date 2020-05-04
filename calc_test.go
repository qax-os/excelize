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
		// ACOS
		"=ACOS(-1)": "3.141592653589793",
		"=ACOS(0)":  "1.5707963267948966",
		// ACOSH
		"=ACOSH(1)":   "0",
		"=ACOSH(2.5)": "1.566799236972411",
		"=ACOSH(5)":   "2.2924316695611777",
		// ACOT
		"=_xlfn.ACOT(1)":  "0.7853981633974483",
		"=_xlfn.ACOT(-2)": "2.677945044588987",
		"=_xlfn.ACOT(0)":  "1.5707963267948966",
		// ACOTH
		"=_xlfn.ACOTH(-5)":  "-0.2027325540540822",
		"=_xlfn.ACOTH(1.1)": "1.5222612188617113",
		"=_xlfn.ACOTH(2)":   "0.5493061443340548",
		// ARABIC
		`=_xlfn.ARABIC("IV")`:   "4",
		`=_xlfn.ARABIC("-IV")`:  "-4",
		`=_xlfn.ARABIC("MCXX")`: "1120",
		`=_xlfn.ARABIC("")`:     "0",
		// ASIN
		"=ASIN(-1)": "-1.5707963267948966",
		"=ASIN(0)":  "0",
		// ASINH
		"=ASINH(0)":    "0",
		"=ASINH(-0.5)": "-0.48121182505960347",
		"=ASINH(2)":    "1.4436354751788103",
		// ATAN
		"=ATAN(-1)": "-0.7853981633974483",
		"=ATAN(0)":  "0",
		"=ATAN(1)":  "0.7853981633974483",
		// ATANH
		"=ATANH(-0.8)": "-1.0986122886681098",
		"=ATANH(0)":    "0",
		"=ATANH(0.5)":  "0.5493061443340548",
		// ATAN2
		"=ATAN2(1,1)":  "0.7853981633974483",
		"=ATAN2(1,-1)": "-0.7853981633974483",
		"=ATAN2(4,0)":  "0",
		// BASE
		"=BASE(12,2)":      "1100",
		"=BASE(12,2,8)":    "00001100",
		"=BASE(100000,16)": "186A0",
		// CEILING
		"=CEILING(22.25,0.1)":   "22.3",
		"=CEILING(22.25,0.5)":   "22.5",
		"=CEILING(22.25,1)":     "23",
		"=CEILING(22.25,10)":    "30",
		"=CEILING(22.25,20)":    "40",
		"=CEILING(-22.25,-0.1)": "-22.3",
		"=CEILING(-22.25,-1)":   "-23",
		"=CEILING(-22.25,-5)":   "-25",
		// _xlfn.CEILING.MATH
		"=_xlfn.CEILING.MATH(15.25,1)":    "16",
		"=_xlfn.CEILING.MATH(15.25,0.1)":  "15.3",
		"=_xlfn.CEILING.MATH(15.25,5)":    "20",
		"=_xlfn.CEILING.MATH(-15.25,1)":   "-15",
		"=_xlfn.CEILING.MATH(-15.25,1,1)": "-15", // should be 16
		"=_xlfn.CEILING.MATH(-15.25,10)":  "-10",
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
		"=ABS()":  "ABS requires 1 numeric arguments",
		"=ABS(~)": `cannot convert cell "~" to coordinates: invalid cell name "~"`,
		// ACOS
		"=ACOS()": "ACOS requires 1 numeric arguments",
		// ACOSH
		"=ACOSH()": "ACOSH requires 1 numeric arguments",
		// _xlfn.ACOT
		"=_xlfn.ACOT()": "ACOT requires 1 numeric arguments",
		// _xlfn.ACOTH
		"=_xlfn.ACOTH()": "ACOTH requires 1 numeric arguments",
		// _xlfn.ARABIC
		"_xlfn.ARABIC()": "ARABIC requires 1 numeric arguments",
		// ASIN
		"=ASIN()": "ASIN requires 1 numeric arguments",
		// ASINH
		"=ASINH()": "ASINH requires 1 numeric arguments",
		// ATAN
		"=ATAN()": "ATAN requires 1 numeric arguments",
		// ATANH
		"=ATANH()": "ATANH requires 1 numeric arguments",
		// ATAN2
		"=ATAN2()": "ATAN2 requires 2 numeric arguments",
		// BASE
		"=BASE()":        "BASE requires at least 2 arguments",
		"=BASE(1,2,3,4)": "BASE allows at most 3 arguments",
		"=BASE(1,1)":     "radix must be an integer ≥ 2 and ≤ 36",
		// CEILING
		"=CEILING()":      "CEILING requires at least 1 argument",
		"=CEILING(1,2,3)": "CEILING allows at most 2 arguments",
		"=CEILING(1,-1)":  "negative sig to CEILING invalid",
		// _xlfn.CEILING.MATH
		"=_xlfn.CEILING.MATH()":        "CEILING.MATH requires at least 1 argument",
		"=_xlfn.CEILING.MATH(1,2,3,4)": "CEILING.MATH allows at most 3 arguments",
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
