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
		f.SetCellValue("Sheet1", "B2", 5)
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
		// _xlfn.CEILING.PRECISE
		"=_xlfn.CEILING.PRECISE(22.25,0.1)": "22.3",
		"=_xlfn.CEILING.PRECISE(22.25,0.5)": "22.5",
		"=_xlfn.CEILING.PRECISE(22.25,1)":   "23",
		"=_xlfn.CEILING.PRECISE(22.25)":     "23",
		"=_xlfn.CEILING.PRECISE(22.25,10)":  "30",
		"=_xlfn.CEILING.PRECISE(22.25,0)":   "0",
		"=_xlfn.CEILING.PRECISE(-22.25,1)":  "-22",
		"=_xlfn.CEILING.PRECISE(-22.25,-1)": "-22",
		"=_xlfn.CEILING.PRECISE(-22.25,5)":  "-20",
		// COMBIN
		"=COMBIN(6,1)": "6",
		"=COMBIN(6,2)": "15",
		"=COMBIN(6,3)": "20",
		"=COMBIN(6,4)": "15",
		"=COMBIN(6,5)": "6",
		"=COMBIN(6,6)": "1",
		// _xlfn.COMBINA
		"=_xlfn.COMBINA(6,1)": "6",
		"=_xlfn.COMBINA(6,2)": "21",
		"=_xlfn.COMBINA(6,3)": "56",
		"=_xlfn.COMBINA(6,4)": "126",
		"=_xlfn.COMBINA(6,5)": "252",
		"=_xlfn.COMBINA(6,6)": "462",
		// COS
		"=COS(0.785398163)": "0.707106781467586",
		"=COS(0)":           "1",
		// COSH
		"=COSH(0)":   "1",
		"=COSH(0.5)": "1.1276259652063807",
		"=COSH(-2)":  "3.7621956910836314",
		// _xlfn.COT
		"=_xlfn.COT(0.785398163397448)": "0.9999999999999992",
		// _xlfn.COTH
		"=_xlfn.COTH(-3.14159265358979)": "-0.9962720762207499",
		// _xlfn.CSC
		"=_xlfn.CSC(-6)":              "3.5788995472544056",
		"=_xlfn.CSC(1.5707963267949)": "1",
		// _xlfn.CSCH
		"=_xlfn.CSCH(-3.14159265358979)": "-0.08658953753004724",
		// _xlfn.DECIMAL
		`=_xlfn.DECIMAL("1100",2)`:   "12",
		`=_xlfn.DECIMAL("186A0",16)`: "100000",
		`=_xlfn.DECIMAL("31L0",32)`:  "100000",
		`=_xlfn.DECIMAL("70122",8)`:  "28754",
		// DEGREES
		"=DEGREES(1)":   "57.29577951308232",
		"=DEGREES(2.5)": "143.2394487827058",
		// EVEN
		"=EVEN(23)":   "24",
		"=EVEN(2.22)": "4",
		"=EVEN(0)":    "0",
		"=EVEN(-0.3)": "-2",
		"=EVEN(-11)":  "-12",
		"=EVEN(-4)":   "-4",
		// EXP
		"=EXP(100)": "2.6881171418161356E+43",
		"=EXP(0.1)": "1.1051709180756477",
		"=EXP(0)":   "1",
		"=EXP(-5)":  "0.006737946999085467",
		// FACT
		"=FACT(3)":  "6",
		"=FACT(6)":  "720",
		"=FACT(10)": "3.6288E+06",
		// FACTDOUBLE
		"=FACTDOUBLE(5)":  "15",
		"=FACTDOUBLE(8)":  "384",
		"=FACTDOUBLE(13)": "135135",
		// FLOOR
		"=FLOOR(26.75,0.1)":   "26.700000000000003",
		"=FLOOR(26.75,0.5)":   "26.5",
		"=FLOOR(26.75,1)":     "26",
		"=FLOOR(26.75,10)":    "20",
		"=FLOOR(26.75,20)":    "20",
		"=FLOOR(-26.75,-0.1)": "-26.700000000000003",
		"=FLOOR(-26.75,-1)":   "-26",
		"=FLOOR(-26.75,-5)":   "-25",
		// _xlfn.FLOOR.MATH
		"=_xlfn.FLOOR.MATH(58.55)":       "58",
		"=_xlfn.FLOOR.MATH(58.55,0.1)":   "58.5",
		"=_xlfn.FLOOR.MATH(58.55,5)":     "55",
		"=_xlfn.FLOOR.MATH(58.55,1,1)":   "58",
		"=_xlfn.FLOOR.MATH(-58.55,1)":    "-59",
		"=_xlfn.FLOOR.MATH(-58.55,1,-1)": "-58",
		"=_xlfn.FLOOR.MATH(-58.55,1,1)":  "-59", // should be -58
		"=_xlfn.FLOOR.MATH(-58.55,10)":   "-60",
		// _xlfn.FLOOR.PRECISE
		"=_xlfn.FLOOR.PRECISE(26.75,0.1)": "26.700000000000003",
		"=_xlfn.FLOOR.PRECISE(26.75,0.5)": "26.5",
		"=_xlfn.FLOOR.PRECISE(26.75,1)":   "26",
		"=_xlfn.FLOOR.PRECISE(26.75)":     "26",
		"=_xlfn.FLOOR.PRECISE(26.75,10)":  "20",
		"=_xlfn.FLOOR.PRECISE(26.75,0)":   "0",
		"=_xlfn.FLOOR.PRECISE(-26.75,1)":  "-27",
		"=_xlfn.FLOOR.PRECISE(-26.75,-1)": "-27",
		"=_xlfn.FLOOR.PRECISE(-26.75,-5)": "-30",
		// GCD
		"=GCD(1,5)":      "1",
		"=GCD(15,10,25)": "5",
		"=GCD(0,8,12)":   "4",
		"=GCD(7,2)":      "1",
		// INT
		"=INT(100.9)":  "100",
		"=INT(5.22)":   "5",
		"=INT(5.99)":   "5",
		"=INT(-6.1)":   "-7",
		"=INT(-100.9)": "-101",
		// ISO.CEILING
		"=ISO.CEILING(22.25)":      "23",
		"=ISO.CEILING(22.25,1)":    "23",
		"=ISO.CEILING(22.25,0.1)":  "22.3",
		"=ISO.CEILING(22.25,10)":   "30",
		"=ISO.CEILING(-22.25,1)":   "-22",
		"=ISO.CEILING(-22.25,0.1)": "-22.200000000000003",
		"=ISO.CEILING(-22.25,5)":   "-20",
		// LCM
		"=LCM(1,5)":      "5",
		"=LCM(15,10,25)": "150",
		"=LCM(1,8,12)":   "24",
		"=LCM(7,2)":      "14",
		// LN
		"=LN(1)":   "0",
		"=LN(100)": "4.605170185988092",
		"=LN(0.5)": "-0.6931471805599453",
		// LOG
		"=LOG(64,2)":  "6",
		"=LOG(100)":   "2",
		"=LOG(4,0.5)": "-2",
		"=LOG(500)":   "2.6989700043360183",
		// LOG10
		"=LOG10(100)":   "2",
		"=LOG10(1000)":  "3",
		"=LOG10(0.001)": "-3",
		"=LOG10(25)":    "1.3979400086720375",
		// MOD
		"=MOD(6,4)":     "2",
		"=MOD(6,3)":     "0",
		"=MOD(6,2.5)":   "1",
		"=MOD(6,1.333)": "0.6680000000000001",
		// MROUND
		"=MROUND(333.7,0.5)":   "333.5",
		"=MROUND(333.8,1)":     "334",
		"=MROUND(333.3,2)":     "334",
		"=MROUND(555.3,400)":   "400",
		"=MROUND(555,1000)":    "1000",
		"=MROUND(-555.7,-1)":   "-556",
		"=MROUND(-555.4,-1)":   "-555",
		"=MROUND(-1555,-1000)": "-2000",
		// MULTINOMIAL
		"=MULTINOMIAL(3,1,2,5)": "27720",
		// _xlfn.MUNIT
		"=_xlfn.MUNIT(4)": "", // not support currently
		// ODD
		"=ODD(22)":     "23",
		"=ODD(1.22)":   "3",
		"=ODD(1.22+4)": "7",
		"=ODD(0)":      "1",
		"=ODD(-1.3)":   "-3",
		"=ODD(-10)":    "-11",
		"=ODD(-3)":     "-3",
		// PI
		"=PI()": "3.141592653589793",
		// POWER
		"=POWER(4,2)": "16",
		// PRODUCT
		"=PRODUCT(3,6)": "18",
		// QUOTIENT
		"=QUOTIENT(5,2)":     "2",
		"=QUOTIENT(4.5,3.1)": "1",
		"=QUOTIENT(-10,3)":   "-3",
		// RADIANS
		"=RADIANS(50)":   "0.8726646259971648",
		"=RADIANS(-180)": "-3.141592653589793",
		"=RADIANS(180)":  "3.141592653589793",
		"=RADIANS(360)":  "6.283185307179586",
		// ROMAN
		"=ROMAN(499,0)":  "CDXCIX",
		"=ROMAN(1999,0)": "MCMXCIX",
		"=ROMAN(1999,1)": "MLMVLIV",
		"=ROMAN(1999,2)": "MXMIX",
		"=ROMAN(1999,3)": "MVMIV",
		"=ROMAN(1999,4)": "MIM",
		// ROUND
		"=ROUND(100.319,1)": "100.30000000000001",
		"=ROUND(5.28,1)":    "5.300000000000001",
		"=ROUND(5.9999,3)":  "6.000000000000002",
		"=ROUND(99.5,0)":    "100",
		"=ROUND(-6.3,0)":    "-6",
		"=ROUND(-100.5,0)":  "-101",
		"=ROUND(-22.45,1)":  "-22.5",
		"=ROUND(999,-1)":    "1000",
		"=ROUND(991,-1)":    "990",
		// ROUNDDOWN
		"=ROUNDDOWN(99.999,1)":   "99.9",
		"=ROUNDDOWN(99.999,2)":   "99.99000000000002",
		"=ROUNDDOWN(99.999,0)":   "99",
		"=ROUNDDOWN(99.999,-1)":  "90",
		"=ROUNDDOWN(-99.999,2)":  "-99.99000000000002",
		"=ROUNDDOWN(-99.999,-1)": "-90",
		// ROUNDUP
		"=ROUNDUP(11.111,1)":   "11.200000000000001",
		"=ROUNDUP(11.111,2)":   "11.120000000000003",
		"=ROUNDUP(11.111,0)":   "12",
		"=ROUNDUP(11.111,-1)":  "20",
		"=ROUNDUP(-11.111,2)":  "-11.120000000000003",
		"=ROUNDUP(-11.111,-1)": "-20",
		// SEC
		"=_xlfn.SEC(-3.14159265358979)": "-1",
		"=_xlfn.SEC(0)":                 "1",
		// SECH
		"=_xlfn.SECH(-3.14159265358979)": "0.0862667383340547",
		"=_xlfn.SECH(0)":                 "1",
		// SIGN
		"=SIGN(9.5)":        "1",
		"=SIGN(-9.5)":       "-1",
		"=SIGN(0)":          "0",
		"=SIGN(0.00000001)": "1",
		"=SIGN(6-7)":        "-1",
		// SIN
		"=SIN(0.785398163)": "0.7071067809055092",
		// SINH
		"=SINH(0)":   "0",
		"=SINH(0.5)": "0.5210953054937474",
		"=SINH(-2)":  "-3.626860407847019",
		// SQRT
		"=SQRT(4)": "2",
		// SQRTPI
		"=SQRTPI(5)":   "3.963327297606011",
		"=SQRTPI(0.2)": "0.7926654595212022",
		"=SQRTPI(100)": "17.72453850905516",
		"=SQRTPI(0)":   "0",
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
		// SUMSQ
		"=SUMSQ(A1:A4)":         "14",
		"=SUMSQ(A1,B1,A2,B2,6)": "82",
		// TAN
		"=TAN(1.047197551)": "1.732050806782486",
		"=TAN(0)":           "0",
		// TANH
		"=TANH(0)":   "0",
		"=TANH(0.5)": "0.46211715726000974",
		"=TANH(-2)":  "-0.9640275800758169",
		// TRUNC
		"=TRUNC(99.999,1)":   "99.9",
		"=TRUNC(99.999,2)":   "99.99",
		"=TRUNC(99.999)":     "99",
		"=TRUNC(99.999,-1)":  "90",
		"=TRUNC(-99.999,2)":  "-99.99",
		"=TRUNC(-99.999,-1)": "-90",
	}
	for formula, expected := range mathCalc {
		f := prepareData()
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err)
		assert.Equal(t, expected, result, formula)
	}
	mathCalcError := map[string]string{
		// ABS
		"=ABS()":  "ABS requires 1 numeric argument",
		"=ABS(~)": `cannot convert cell "~" to coordinates: invalid cell name "~"`,
		// ACOS
		"=ACOS()": "ACOS requires 1 numeric argument",
		// ACOSH
		"=ACOSH()": "ACOSH requires 1 numeric argument",
		// _xlfn.ACOT
		"=_xlfn.ACOT()": "ACOT requires 1 numeric argument",
		// _xlfn.ACOTH
		"=_xlfn.ACOTH()": "ACOTH requires 1 numeric argument",
		// _xlfn.ARABIC
		"=_xlfn.ARABIC()": "ARABIC requires 1 numeric argument",
		// ASIN
		"=ASIN()": "ASIN requires 1 numeric argument",
		// ASINH
		"=ASINH()": "ASINH requires 1 numeric argument",
		// ATAN
		"=ATAN()": "ATAN requires 1 numeric argument",
		// ATANH
		"=ATANH()": "ATANH requires 1 numeric argument",
		// ATAN2
		"=ATAN2()": "ATAN2 requires 2 numeric arguments",
		// BASE
		"=BASE()":        "BASE requires at least 2 arguments",
		"=BASE(1,2,3,4)": "BASE allows at most 3 arguments",
		"=BASE(1,1)":     "radix must be an integer >= 2 and <= 36",
		// CEILING
		"=CEILING()":      "CEILING requires at least 1 argument",
		"=CEILING(1,2,3)": "CEILING allows at most 2 arguments",
		"=CEILING(1,-1)":  "negative sig to CEILING invalid",
		// _xlfn.CEILING.MATH
		"=_xlfn.CEILING.MATH()":        "CEILING.MATH requires at least 1 argument",
		"=_xlfn.CEILING.MATH(1,2,3,4)": "CEILING.MATH allows at most 3 arguments",
		// _xlfn.CEILING.PRECISE
		"=_xlfn.CEILING.PRECISE()":      "CEILING.PRECISE requires at least 1 argument",
		"=_xlfn.CEILING.PRECISE(1,2,3)": "CEILING.PRECISE allows at most 2 arguments",
		// COMBIN
		"=COMBIN()":     "COMBIN requires 2 argument",
		"=COMBIN(-1,1)": "COMBIN requires number >= number_chosen",
		// _xlfn.COMBINA
		"=_xlfn.COMBINA()":      "COMBINA requires 2 argument",
		"=_xlfn.COMBINA(-1,1)":  "COMBINA requires number > number_chosen",
		"=_xlfn.COMBINA(-1,-1)": "COMBIN requires number >= number_chosen",
		// COS
		"=COS()": "COS requires 1 numeric argument",
		// COSH
		"=COSH()": "COSH requires 1 numeric argument",
		// _xlfn.COT
		"=COT()": "COT requires 1 numeric argument",
		// _xlfn.COTH
		"=COTH()": "COTH requires 1 numeric argument",
		// _xlfn.CSC
		"=_xlfn.CSC()":  "CSC requires 1 numeric argument",
		"=_xlfn.CSC(0)": "#NAME?",
		// _xlfn.CSCH
		"=_xlfn.CSCH()":  "CSCH requires 1 numeric argument",
		"=_xlfn.CSCH(0)": "#NAME?",
		// _xlfn.DECIMAL
		"=_xlfn.DECIMAL()":          "DECIMAL requires 2 numeric arguments",
		`=_xlfn.DECIMAL("2000", 2)`: "#NUM!",
		// DEGREES
		"=DEGREES()": "DEGREES requires 1 numeric argument",
		// EVEN
		"=EVEN()": "EVEN requires 1 numeric argument",
		// EXP
		"=EXP()": "EXP requires 1 numeric argument",
		// FACT
		"=FACT()":   "FACT requires 1 numeric argument",
		"=FACT(-1)": "#NUM!",
		// FACTDOUBLE
		"=FACTDOUBLE()":   "FACTDOUBLE requires 1 numeric argument",
		"=FACTDOUBLE(-1)": "#NUM!",
		// FLOOR
		"=FLOOR()":     "FLOOR requires 2 numeric arguments",
		"=FLOOR(1,-1)": "#NUM!",
		// _xlfn.FLOOR.MATH
		"=_xlfn.FLOOR.MATH()":        "FLOOR.MATH requires at least 1 argument",
		"=_xlfn.FLOOR.MATH(1,2,3,4)": "FLOOR.MATH allows at most 3 arguments",
		// _xlfn.FLOOR.PRECISE
		"=_xlfn.FLOOR.PRECISE()":      "FLOOR.PRECISE requires at least 1 argument",
		"=_xlfn.FLOOR.PRECISE(1,2,3)": "FLOOR.PRECISE allows at most 2 arguments",
		// GCD
		"=GCD()":     "GCD requires at least 1 argument",
		"=GCD(-1)":   "GCD only accepts positive arguments",
		"=GCD(1,-1)": "GCD only accepts positive arguments",
		// INT
		"=INT()": "INT requires 1 numeric argument",
		// ISO.CEILING
		"=ISO.CEILING()":      "ISO.CEILING requires at least 1 argument",
		"=ISO.CEILING(1,2,3)": "ISO.CEILING allows at most 2 arguments",
		// LCM
		"=LCM()":     "LCM requires at least 1 argument",
		"=LCM(-1)":   "LCM only accepts positive arguments",
		"=LCM(1,-1)": "LCM only accepts positive arguments",
		// LN
		"=LN()": "LN requires 1 numeric argument",
		// LOG
		"=LOG()":      "LOG requires at least 1 argument",
		"=LOG(1,2,3)": "LOG allows at most 2 arguments",
		"=LOG(0,0)":   "#NUM!",
		"=LOG(1,0)":   "#NUM!",
		"=LOG(1,1)":   "#DIV/0!",
		// LOG10
		"=LOG10()": "LOG10 requires 1 numeric argument",
		// MOD
		"=MOD()":    "MOD requires 2 numeric arguments",
		"=MOD(6,0)": "#DIV/0!",
		// MROUND
		"=MROUND()":    "MROUND requires 2 numeric arguments",
		"=MROUND(1,0)": "#NUM!",
		// _xlfn.MUNIT
		"=_xlfn.MUNIT()": "MUNIT requires 1 numeric argument", // not support currently
		// ODD
		"=ODD()": "ODD requires 1 numeric argument",
		// PI
		"=PI(1)": "PI accepts no arguments",
		// POWER
		"=POWER(0,0)":  "#NUM!",
		"=POWER(0,-1)": "#DIV/0!",
		"=POWER(1)":    "POWER requires 2 numeric arguments",
		// QUOTIENT
		"=QUOTIENT(1,0)": "#DIV/0!",
		"=QUOTIENT(1)":   "QUOTIENT requires 2 numeric arguments",
		// RADIANS
		"=RADIANS()": "RADIANS requires 1 numeric argument",
		// RAND
		"=RAND(1)": "RAND accepts no arguments",
		// RANDBETWEEN
		"=RANDBETWEEN()":    "RANDBETWEEN requires 2 numeric arguments",
		"=RANDBETWEEN(2,1)": "#NUM!",
		// ROMAN
		"=ROMAN()":      "ROMAN requires at least 1 argument",
		"=ROMAN(1,2,3)": "ROMAN allows at most 2 arguments",
		// ROUND
		"=ROUND()": "ROUND requires 2 numeric arguments",
		// ROUNDDOWN
		"=ROUNDDOWN()": "ROUNDDOWN requires 2 numeric arguments",
		// ROUNDUP
		"=ROUNDUP()": "ROUNDUP requires 2 numeric arguments",
		// SEC
		"=_xlfn.SEC()": "SEC requires 1 numeric argument",
		// _xlfn.SECH
		"=_xlfn.SECH()": "SECH requires 1 numeric argument",
		// SIGN
		"=SIGN()": "SIGN requires 1 numeric argument",
		// SIN
		"=SIN()": "SIN requires 1 numeric argument",
		// SINH
		"=SINH()": "SINH requires 1 numeric argument",
		// SQRT
		"=SQRT()":   "SQRT requires 1 numeric argument",
		"=SQRT(-1)": "#NUM!",
		// SQRTPI
		"=SQRTPI()": "SQRTPI requires 1 numeric argument",
		// SUM
		"=SUM((":   "formula not valid",
		"=SUM(-)":  "formula not valid",
		"=SUM(1+)": "formula not valid",
		"=SUM(1-)": "formula not valid",
		"=SUM(1*)": "formula not valid",
		"=SUM(1/)": "formula not valid",
		// TAN
		"=TAN()": "TAN requires 1 numeric argument",
		// TANH
		"=TANH()": "TANH requires 1 numeric argument",
		// TRUNC
		"=TRUNC()": "TRUNC requires at least 1 argument",
	}
	for formula, expected := range mathCalcError {
		f := prepareData()
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.EqualError(t, err, expected)
		assert.Equal(t, "", result, formula)
	}

	referenceCalc := map[string]string{
		// MDETERM
		"=MDETERM(A1:B2)": "-3",
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
		"=A1/A2/SUM(A1:A2:B1)":            "0.041666666666666664",
		"=A1/A2/SUM(A1:A2:B1)*A3":         "0.125",
	}
	for formula, expected := range referenceCalc {
		f := prepareData()
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err)
		assert.Equal(t, expected, result, formula)
	}

	referenceCalcError := map[string]string{
		// MDETERM
		"=MDETERM(A1:B3)": "#VALUE!",
		// SUM
		"=1+SUM(SUM(A1+A2/A4)*(2-3),2)": "#DIV/0!",
	}
	for formula, expected := range referenceCalcError {
		f := prepareData()
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.EqualError(t, err, expected)
		assert.Equal(t, "", result, formula)
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
