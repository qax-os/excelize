package excelize

import (
	"container/list"
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/xuri/efp"
)

func TestCalcCellValue(t *testing.T) {
	cellData := [][]interface{}{
		{1, 4, nil, "Month", "Team", "Sales"},
		{2, 5, nil, "Jan", "North 1", 36693},
		{3, nil, nil, "Jan", "North 2", 22100},
		{0, nil, nil, "Jan", "South 1", 53321},
		{nil, nil, nil, "Jan", "South 2", 34440},
		{nil, nil, nil, "Feb", "North 1", 29889},
		{nil, nil, nil, "Feb", "North 2", 50090},
		{nil, nil, nil, "Feb", "South 1", 32080},
		{nil, nil, nil, "Feb", "South 2", 45500},
	}
	prepareData := func() *File {
		f := NewFile()
		for r, row := range cellData {
			for c, value := range row {
				cell, _ := CoordinatesToCellName(c+1, r+1)
				assert.NoError(t, f.SetCellValue("Sheet1", cell, value))
			}
		}
		return f
	}

	mathCalc := map[string]string{
		"=2^3":  "8",
		"=1=1":  "TRUE",
		"=1=2":  "FALSE",
		"=1<2":  "TRUE",
		"=3<2":  "FALSE",
		"=2<=3": "TRUE",
		"=2<=1": "FALSE",
		"=2>1":  "TRUE",
		"=2>3":  "FALSE",
		"=2>=1": "TRUE",
		"=2>=3": "FALSE",
		"=1&2":  "12",
		// ABS
		"=ABS(-1)":    "1",
		"=ABS(-6.5)":  "6.5",
		"=ABS(6.5)":   "6.5",
		"=ABS(0)":     "0",
		"=ABS(2-4.5)": "2.5",
		// ACOS
		"=ACOS(-1)": "3.141592653589793",
		"=ACOS(0)":  "1.570796326794897",
		// ACOSH
		"=ACOSH(1)":   "0",
		"=ACOSH(2.5)": "1.566799236972411",
		"=ACOSH(5)":   "2.292431669561178",
		// ACOT
		"=_xlfn.ACOT(1)":  "0.785398163397448",
		"=_xlfn.ACOT(-2)": "2.677945044588987",
		"=_xlfn.ACOT(0)":  "1.570796326794897",
		// ACOTH
		"=_xlfn.ACOTH(-5)":  "-0.202732554054082",
		"=_xlfn.ACOTH(1.1)": "1.522261218861711",
		"=_xlfn.ACOTH(2)":   "0.549306144334055",
		// ARABIC
		`=_xlfn.ARABIC("IV")`:   "4",
		`=_xlfn.ARABIC("-IV")`:  "-4",
		`=_xlfn.ARABIC("MCXX")`: "1120",
		`=_xlfn.ARABIC("")`:     "0",
		// ASIN
		"=ASIN(-1)": "-1.570796326794897",
		"=ASIN(0)":  "0",
		// ASINH
		"=ASINH(0)":    "0",
		"=ASINH(-0.5)": "-0.481211825059604",
		"=ASINH(2)":    "1.44363547517881",
		// ATAN
		"=ATAN(-1)": "-0.785398163397448",
		"=ATAN(0)":  "0",
		"=ATAN(1)":  "0.785398163397448",
		// ATANH
		"=ATANH(-0.8)": "-1.09861228866811",
		"=ATANH(0)":    "0",
		"=ATANH(0.5)":  "0.549306144334055",
		// ATAN2
		"=ATAN2(1,1)":  "0.785398163397448",
		"=ATAN2(1,-1)": "-0.785398163397448",
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
		"=CEILING(22.25)":       "23",
		// _xlfn.CEILING.MATH
		"=_xlfn.CEILING.MATH(15.25,1)":      "16",
		"=_xlfn.CEILING.MATH(15.25,0.1)":    "15.3",
		"=_xlfn.CEILING.MATH(15.25,5)":      "20",
		"=_xlfn.CEILING.MATH(-15.25,1)":     "-15",
		"=_xlfn.CEILING.MATH(-15.25,1,1)":   "-15", // should be 16
		"=_xlfn.CEILING.MATH(-15.25,10)":    "-10",
		"=_xlfn.CEILING.MATH(-15.25)":       "-15",
		"=_xlfn.CEILING.MATH(-15.25,-5,-1)": "-10",
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
		"=COMBIN(0,0)": "1",
		// _xlfn.COMBINA
		"=_xlfn.COMBINA(6,1)": "6",
		"=_xlfn.COMBINA(6,2)": "21",
		"=_xlfn.COMBINA(6,3)": "56",
		"=_xlfn.COMBINA(6,4)": "126",
		"=_xlfn.COMBINA(6,5)": "252",
		"=_xlfn.COMBINA(6,6)": "462",
		"=_xlfn.COMBINA(0,0)": "0",
		// COS
		"=COS(0.785398163)": "0.707106781467586",
		"=COS(0)":           "1",
		// COSH
		"=COSH(0)":   "1",
		"=COSH(0.5)": "1.127625965206381",
		"=COSH(-2)":  "3.762195691083632",
		// _xlfn.COT
		"=_xlfn.COT(0.785398163397448)": "0.999999999999999",
		// _xlfn.COTH
		"=_xlfn.COTH(-3.14159265358979)": "-0.99627207622075",
		// _xlfn.CSC
		"=_xlfn.CSC(-6)":              "3.578899547254406",
		"=_xlfn.CSC(1.5707963267949)": "1",
		// _xlfn.CSCH
		"=_xlfn.CSCH(-3.14159265358979)": "-0.086589537530047",
		// _xlfn.DECIMAL
		`=_xlfn.DECIMAL("1100",2)`:    "12",
		`=_xlfn.DECIMAL("186A0",16)`:  "100000",
		`=_xlfn.DECIMAL("31L0",32)`:   "100000",
		`=_xlfn.DECIMAL("70122",8)`:   "28754",
		`=_xlfn.DECIMAL("0x70122",8)`: "28754",
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
		"=EXP(0.1)": "1.105170918075648",
		"=EXP(0)":   "1",
		"=EXP(-5)":  "0.006737946999085",
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
		"=GCD(0)":        "0",
		`=GCD("",1)`:     "1",
		"=GCD(1,0)":      "1",
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
		"=ISO.CEILING(-22.25,0)":   "0",
		// LCM
		"=LCM(1,5)":      "5",
		"=LCM(15,10,25)": "150",
		"=LCM(1,8,12)":   "24",
		"=LCM(7,2)":      "14",
		"=LCM(7)":        "7",
		`=LCM("",1)`:     "1",
		`=LCM(0,0)`:      "0",
		// LN
		"=LN(1)":   "0",
		"=LN(100)": "4.605170185988092",
		"=LN(0.5)": "-0.693147180559945",
		// LOG
		"=LOG(64,2)":  "6",
		"=LOG(100)":   "2",
		"=LOG(4,0.5)": "-2",
		"=LOG(500)":   "2.698970004336019",
		// LOG10
		"=LOG10(100)":   "2",
		"=LOG10(1000)":  "3",
		"=LOG10(0.001)": "-3",
		"=LOG10(25)":    "1.397940008672038",
		// MOD
		"=MOD(6,4)":      "2",
		"=MOD(6,3)":      "0",
		"=MOD(6,2.5)":    "1",
		"=MOD(6,1.333)":  "0.668",
		"=MOD(-10.23,1)": "0.77",
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
		"=MULTINOMIAL(3,1,2,5)":    "27720",
		`=MULTINOMIAL("",3,1,2,5)`: "27720",
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
		"=PRODUCT(3,6)":    "18",
		`=PRODUCT("",3,6)`: "18",
		// QUOTIENT
		"=QUOTIENT(5,2)":     "2",
		"=QUOTIENT(4.5,3.1)": "1",
		"=QUOTIENT(-10,3)":   "-3",
		// RADIANS
		"=RADIANS(50)":   "0.872664625997165",
		"=RADIANS(-180)": "-3.141592653589793",
		"=RADIANS(180)":  "3.141592653589793",
		"=RADIANS(360)":  "6.283185307179586",
		// ROMAN
		"=ROMAN(499,0)":   "CDXCIX",
		"=ROMAN(1999,0)":  "MCMXCIX",
		"=ROMAN(1999,1)":  "MLMVLIV",
		"=ROMAN(1999,2)":  "MXMIX",
		"=ROMAN(1999,3)":  "MVMIV",
		"=ROMAN(1999,4)":  "MIM",
		"=ROMAN(1999,-1)": "MCMXCIX",
		"=ROMAN(1999,5)":  "MIM",
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
		// ROUNDUP`
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
		"=_xlfn.SECH(-3.14159265358979)": "0.086266738334055",
		"=_xlfn.SECH(0)":                 "1",
		// SIGN
		"=SIGN(9.5)":        "1",
		"=SIGN(-9.5)":       "-1",
		"=SIGN(0)":          "0",
		"=SIGN(0.00000001)": "1",
		"=SIGN(6-7)":        "-1",
		// SIN
		"=SIN(0.785398163)": "0.707106780905509",
		// SINH
		"=SINH(0)":   "0",
		"=SINH(0.5)": "0.521095305493747",
		"=SINH(-2)":  "-3.626860407847019",
		// SQRT
		"=SQRT(4)":  "2",
		`=SQRT("")`: "0",
		// SQRTPI
		"=SQRTPI(5)":   "3.963327297606011",
		"=SQRTPI(0.2)": "0.792665459521202",
		"=SQRTPI(100)": "17.72453850905516",
		"=SQRTPI(0)":   "0",
		// SUM
		"=SUM(1,2)":                           "3",
		`=SUM("",1,2)`:                        "3",
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
		// SUMIF
		`=SUMIF(F1:F5, "")`:             "0",
		`=SUMIF(A1:A5, "3")`:            "3",
		`=SUMIF(F1:F5, "=36693")`:       "36693",
		`=SUMIF(F1:F5, "<100")`:         "0",
		`=SUMIF(F1:F5, "<=36693")`:      "93233",
		`=SUMIF(F1:F5, ">100")`:         "146554",
		`=SUMIF(F1:F5, ">=100")`:        "146554",
		`=SUMIF(F1:F5, ">=text")`:       "0",
		`=SUMIF(F1:F5, "*Jan",F2:F5)`:   "0",
		`=SUMIF(D3:D7,"Jan",F2:F5)`:     "112114",
		`=SUMIF(D2:D9,"Feb",F2:F9)`:     "157559",
		`=SUMIF(E2:E9,"North 1",F2:F9)`: "66582",
		`=SUMIF(E2:E9,"North*",F2:F9)`:  "138772",
		// SUMSQ
		"=SUMSQ(A1:A4)":            "14",
		"=SUMSQ(A1,B1,A2,B2,6)":    "82",
		`=SUMSQ("",A1,B1,A2,B2,6)`: "82",
		// TAN
		"=TAN(1.047197551)": "1.732050806782486",
		"=TAN(0)":           "0",
		// TANH
		"=TANH(0)":   "0",
		"=TANH(0.5)": "0.46211715726001",
		"=TANH(-2)":  "-0.964027580075817",
		// TRUNC
		"=TRUNC(99.999,1)":   "99.9",
		"=TRUNC(99.999,2)":   "99.99",
		"=TRUNC(99.999)":     "99",
		"=TRUNC(99.999,-1)":  "90",
		"=TRUNC(-99.999,2)":  "-99.99",
		"=TRUNC(-99.999,-1)": "-90",
		// Statistical Functions
		// COUNTA
		`=COUNTA()`:                       "0",
		`=COUNTA(A1:A5,B2:B5,"text",1,2)`: "8",
		// MEDIAN
		"=MEDIAN(A1:A5,12)": "2",
		"=MEDIAN(A1:A5)":    "1.5",
		// Information Functions
		// ISBLANK
		"=ISBLANK(A1)": "FALSE",
		"=ISBLANK(A5)": "TRUE",
		// ISERR
		"=ISERR(A1)":   "FALSE",
		"=ISERR(NA())": "FALSE",
		// ISERROR
		"=ISERROR(A1)":   "FALSE",
		"=ISERROR(NA())": "TRUE",
		// ISEVEN
		"=ISEVEN(A1)": "FALSE",
		"=ISEVEN(A2)": "TRUE",
		// ISNA
		"=ISNA(A1)":   "FALSE",
		"=ISNA(NA())": "TRUE",
		// ISNONTEXT
		"=ISNONTEXT(A1)":         "FALSE",
		"=ISNONTEXT(A5)":         "TRUE",
		`=ISNONTEXT("Excelize")`: "FALSE",
		"=ISNONTEXT(NA())":       "FALSE",
		// ISNUMBER
		"=ISNUMBER(A1)": "TRUE",
		"=ISNUMBER(D1)": "FALSE",
		// ISODD
		"=ISODD(A1)": "TRUE",
		"=ISODD(A2)": "FALSE",
		// NA
		"=NA()": "#N/A",
		// Logical Functions
		// AND
		"=AND(0)":               "FALSE",
		"=AND(1)":               "TRUE",
		"=AND(1,0)":             "FALSE",
		"=AND(0,1)":             "FALSE",
		"=AND(1=1)":             "TRUE",
		"=AND(1<2)":             "TRUE",
		"=AND(1>2,2<3,2>0,3>1)": "FALSE",
		"=AND(1=1),1=1":         "TRUE",
		// OR
		"=OR(1)":       "TRUE",
		"=OR(0)":       "FALSE",
		"=OR(1=2,2=2)": "TRUE",
		"=OR(1=2,2=3)": "FALSE",
		// Date and Time Functions
		// DATE
		"=DATE(2020,10,21)": "2020-10-21 00:00:00 +0000 UTC",
		"=DATE(1900,1,1)":   "1899-12-31 00:00:00 +0000 UTC",
		// Text Functions
		// CLEAN
		"=CLEAN(\"\u0009clean text\")": "clean text",
		"=CLEAN(0)":                    "0",
		// LEN
		"=LEN(\"\")": "0",
		"=LEN(D1)":   "5",
		// TRIM
		"=TRIM(\" trim text \")": "trim text",
		"=TRIM(0)":               "0",
		// LOWER
		"=LOWER(\"test\")":     "test",
		"=LOWER(\"TEST\")":     "test",
		"=LOWER(\"Test\")":     "test",
		"=LOWER(\"TEST 123\")": "test 123",
		// PROPER
		"=PROPER(\"this is a test sentence\")": "This Is A Test Sentence",
		"=PROPER(\"THIS IS A TEST SENTENCE\")": "This Is A Test Sentence",
		"=PROPER(\"123tEST teXT\")":            "123Test Text",
		"=PROPER(\"Mr. SMITH's address\")":     "Mr. Smith'S Address",
		// UPPER
		"=UPPER(\"test\")":     "TEST",
		"=UPPER(\"TEST\")":     "TEST",
		"=UPPER(\"Test\")":     "TEST",
		"=UPPER(\"TEST 123\")": "TEST 123",
		// Conditional Functions
		// IF
		"=IF(1=1)":                              "TRUE",
		"=IF(1<>1)":                             "FALSE",
		"=IF(5<0, \"negative\", \"positive\")":  "positive",
		"=IF(-2<0, \"negative\", \"positive\")": "negative",
		// Excel Lookup and Reference Functions
		// CHOOSE
		"=CHOOSE(4,\"red\",\"blue\",\"green\",\"brown\")": "brown",
		"=CHOOSE(1,\"red\",\"blue\",\"green\",\"brown\")": "red",
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
		"=ABS()":    "ABS requires 1 numeric argument",
		`=ABS("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		"=ABS(~)":   `cannot convert cell "~" to coordinates: invalid cell name "~"`,
		// ACOS
		"=ACOS()":    "ACOS requires 1 numeric argument",
		`=ACOS("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// ACOSH
		"=ACOSH()":    "ACOSH requires 1 numeric argument",
		`=ACOSH("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// _xlfn.ACOT
		"=_xlfn.ACOT()":    "ACOT requires 1 numeric argument",
		`=_xlfn.ACOT("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// _xlfn.ACOTH
		"=_xlfn.ACOTH()":    "ACOTH requires 1 numeric argument",
		`=_xlfn.ACOTH("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// _xlfn.ARABIC
		"=_xlfn.ARABIC()": "ARABIC requires 1 numeric argument",
		// ASIN
		"=ASIN()":    "ASIN requires 1 numeric argument",
		`=ASIN("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// ASINH
		"=ASINH()":    "ASINH requires 1 numeric argument",
		`=ASINH("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// ATAN
		"=ATAN()":    "ATAN requires 1 numeric argument",
		`=ATAN("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// ATANH
		"=ATANH()":    "ATANH requires 1 numeric argument",
		`=ATANH("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// ATAN2
		"=ATAN2()":      "ATAN2 requires 2 numeric arguments",
		`=ATAN2("X",0)`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=ATAN2(0,"X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// BASE
		"=BASE()":        "BASE requires at least 2 arguments",
		"=BASE(1,2,3,4)": "BASE allows at most 3 arguments",
		"=BASE(1,1)":     "radix must be an integer >= 2 and <= 36",
		`=BASE("X",2)`:   "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=BASE(1,"X")`:   "strconv.Atoi: parsing \"X\": invalid syntax",
		`=BASE(1,2,"X")`: "strconv.Atoi: parsing \"X\": invalid syntax",
		// CEILING
		"=CEILING()":      "CEILING requires at least 1 argument",
		"=CEILING(1,2,3)": "CEILING allows at most 2 arguments",
		"=CEILING(1,-1)":  "negative sig to CEILING invalid",
		`=CEILING("X",0)`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=CEILING(0,"X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// _xlfn.CEILING.MATH
		"=_xlfn.CEILING.MATH()":        "CEILING.MATH requires at least 1 argument",
		"=_xlfn.CEILING.MATH(1,2,3,4)": "CEILING.MATH allows at most 3 arguments",
		`=_xlfn.CEILING.MATH("X")`:     "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=_xlfn.CEILING.MATH(1,"X")`:   "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=_xlfn.CEILING.MATH(1,2,"X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// _xlfn.CEILING.PRECISE
		"=_xlfn.CEILING.PRECISE()":      "CEILING.PRECISE requires at least 1 argument",
		"=_xlfn.CEILING.PRECISE(1,2,3)": "CEILING.PRECISE allows at most 2 arguments",
		`=_xlfn.CEILING.PRECISE("X",2)`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=_xlfn.CEILING.PRECISE(1,"X")`: "#VALUE!",
		// COMBIN
		"=COMBIN()":       "COMBIN requires 2 argument",
		"=COMBIN(-1,1)":   "COMBIN requires number >= number_chosen",
		`=COMBIN("X",1)`:  "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=COMBIN(-1,"X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// _xlfn.COMBINA
		"=_xlfn.COMBINA()":       "COMBINA requires 2 argument",
		"=_xlfn.COMBINA(-1,1)":   "COMBINA requires number > number_chosen",
		"=_xlfn.COMBINA(-1,-1)":  "COMBIN requires number >= number_chosen",
		`=_xlfn.COMBINA("X",1)`:  "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=_xlfn.COMBINA(-1,"X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// COS
		"=COS()":    "COS requires 1 numeric argument",
		`=COS("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// COSH
		"=COSH()":    "COSH requires 1 numeric argument",
		`=COSH("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// _xlfn.COT
		"=COT()":    "COT requires 1 numeric argument",
		`=COT("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		"=COT(0)":   "#DIV/0!",
		// _xlfn.COTH
		"=COTH()":    "COTH requires 1 numeric argument",
		`=COTH("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		"=COTH(0)":   "#DIV/0!",
		// _xlfn.CSC
		"=_xlfn.CSC()":    "CSC requires 1 numeric argument",
		`=_xlfn.CSC("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		"=_xlfn.CSC(0)":   "#DIV/0!",
		// _xlfn.CSCH
		"=_xlfn.CSCH()":    "CSCH requires 1 numeric argument",
		`=_xlfn.CSCH("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		"=_xlfn.CSCH(0)":   "#DIV/0!",
		// _xlfn.DECIMAL
		"=_xlfn.DECIMAL()":          "DECIMAL requires 2 numeric arguments",
		`=_xlfn.DECIMAL("X", 2)`:    "strconv.ParseInt: parsing \"X\": invalid syntax",
		`=_xlfn.DECIMAL(2000, "X")`: "strconv.Atoi: parsing \"X\": invalid syntax",
		// DEGREES
		"=DEGREES()":    "DEGREES requires 1 numeric argument",
		`=DEGREES("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		"=DEGREES(0)":   "#DIV/0!",
		// EVEN
		"=EVEN()":    "EVEN requires 1 numeric argument",
		`=EVEN("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// EXP
		"=EXP()":    "EXP requires 1 numeric argument",
		`=EXP("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// FACT
		"=FACT()":    "FACT requires 1 numeric argument",
		`=FACT("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		"=FACT(-1)":  "#NUM!",
		// FACTDOUBLE
		"=FACTDOUBLE()":    "FACTDOUBLE requires 1 numeric argument",
		`=FACTDOUBLE("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		"=FACTDOUBLE(-1)":  "#NUM!",
		// FLOOR
		"=FLOOR()":       "FLOOR requires 2 numeric arguments",
		`=FLOOR("X",-1)`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=FLOOR(1,"X")`:  "strconv.ParseFloat: parsing \"X\": invalid syntax",
		"=FLOOR(1,-1)":   "invalid arguments to FLOOR",
		// _xlfn.FLOOR.MATH
		"=_xlfn.FLOOR.MATH()":        "FLOOR.MATH requires at least 1 argument",
		"=_xlfn.FLOOR.MATH(1,2,3,4)": "FLOOR.MATH allows at most 3 arguments",
		`=_xlfn.FLOOR.MATH("X",2,3)`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=_xlfn.FLOOR.MATH(1,"X",3)`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=_xlfn.FLOOR.MATH(1,2,"X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// _xlfn.FLOOR.PRECISE
		"=_xlfn.FLOOR.PRECISE()":      "FLOOR.PRECISE requires at least 1 argument",
		"=_xlfn.FLOOR.PRECISE(1,2,3)": "FLOOR.PRECISE allows at most 2 arguments",
		`=_xlfn.FLOOR.PRECISE("X",2)`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=_xlfn.FLOOR.PRECISE(1,"X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// GCD
		"=GCD()":     "GCD requires at least 1 argument",
		"=GCD(-1)":   "GCD only accepts positive arguments",
		"=GCD(1,-1)": "GCD only accepts positive arguments",
		`=GCD("X")`:  "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// INT
		"=INT()":    "INT requires 1 numeric argument",
		`=INT("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// ISO.CEILING
		"=ISO.CEILING()":      "ISO.CEILING requires at least 1 argument",
		"=ISO.CEILING(1,2,3)": "ISO.CEILING allows at most 2 arguments",
		`=ISO.CEILING("X",2)`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=ISO.CEILING(1,"X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// LCM
		"=LCM()":     "LCM requires at least 1 argument",
		"=LCM(-1)":   "LCM only accepts positive arguments",
		"=LCM(1,-1)": "LCM only accepts positive arguments",
		`=LCM("X")`:  "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// LN
		"=LN()":    "LN requires 1 numeric argument",
		`=LN("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// LOG
		"=LOG()":      "LOG requires at least 1 argument",
		"=LOG(1,2,3)": "LOG allows at most 2 arguments",
		`=LOG("X",1)`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=LOG(1,"X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		"=LOG(0,0)":   "#DIV/0!",
		"=LOG(1,0)":   "#DIV/0!",
		"=LOG(1,1)":   "#DIV/0!",
		// LOG10
		"=LOG10()":    "LOG10 requires 1 numeric argument",
		`=LOG10("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// MOD
		"=MOD()":      "MOD requires 2 numeric arguments",
		"=MOD(6,0)":   "MOD divide by zero",
		`=MOD("X",0)`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=MOD(6,"X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// MROUND
		"=MROUND()":      "MROUND requires 2 numeric arguments",
		"=MROUND(1,0)":   "#NUM!",
		"=MROUND(1,-1)":  "#NUM!",
		`=MROUND("X",0)`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=MROUND(1,"X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// MULTINOMIAL
		`=MULTINOMIAL("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// _xlfn.MUNIT
		"=_xlfn.MUNIT()":    "MUNIT requires 1 numeric argument",           // not support currently
		`=_xlfn.MUNIT("X")`: "strconv.Atoi: parsing \"X\": invalid syntax", // not support currently
		// ODD
		"=ODD()":    "ODD requires 1 numeric argument",
		`=ODD("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// PI
		"=PI(1)": "PI accepts no arguments",
		// POWER
		`=POWER("X",1)`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=POWER(1,"X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		"=POWER(0,0)":   "#NUM!",
		"=POWER(0,-1)":  "#DIV/0!",
		"=POWER(1)":     "POWER requires 2 numeric arguments",
		// PRODUCT
		`=PRODUCT("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// QUOTIENT
		`=QUOTIENT("X",1)`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=QUOTIENT(1,"X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		"=QUOTIENT(1,0)":   "#DIV/0!",
		"=QUOTIENT(1)":     "QUOTIENT requires 2 numeric arguments",
		// RADIANS
		`=RADIANS("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		"=RADIANS()":    "RADIANS requires 1 numeric argument",
		// RAND
		"=RAND(1)": "RAND accepts no arguments",
		// RANDBETWEEN
		`=RANDBETWEEN("X",1)`: "strconv.ParseInt: parsing \"X\": invalid syntax",
		`=RANDBETWEEN(1,"X")`: "strconv.ParseInt: parsing \"X\": invalid syntax",
		"=RANDBETWEEN()":      "RANDBETWEEN requires 2 numeric arguments",
		"=RANDBETWEEN(2,1)":   "#NUM!",
		// ROMAN
		"=ROMAN()":      "ROMAN requires at least 1 argument",
		"=ROMAN(1,2,3)": "ROMAN allows at most 2 arguments",
		`=ROMAN("X")`:   "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=ROMAN("X",1)`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// ROUND
		"=ROUND()":      "ROUND requires 2 numeric arguments",
		`=ROUND("X",1)`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=ROUND(1,"X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// ROUNDDOWN
		"=ROUNDDOWN()":      "ROUNDDOWN requires 2 numeric arguments",
		`=ROUNDDOWN("X",1)`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=ROUNDDOWN(1,"X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// ROUNDUP
		"=ROUNDUP()":      "ROUNDUP requires 2 numeric arguments",
		`=ROUNDUP("X",1)`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=ROUNDUP(1,"X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// SEC
		"=_xlfn.SEC()":    "SEC requires 1 numeric argument",
		`=_xlfn.SEC("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// _xlfn.SECH
		"=_xlfn.SECH()":    "SECH requires 1 numeric argument",
		`=_xlfn.SECH("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// SIGN
		"=SIGN()":    "SIGN requires 1 numeric argument",
		`=SIGN("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// SIN
		"=SIN()":    "SIN requires 1 numeric argument",
		`=SIN("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// SINH
		"=SINH()":    "SINH requires 1 numeric argument",
		`=SINH("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// SQRT
		"=SQRT()":    "SQRT requires 1 numeric argument",
		`=SQRT("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		"=SQRT(-1)":  "#NUM!",
		// SQRTPI
		"=SQRTPI()":    "SQRTPI requires 1 numeric argument",
		`=SQRTPI("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// SUM
		"=SUM((":    "formula not valid",
		"=SUM(-)":   "formula not valid",
		"=SUM(1+)":  "formula not valid",
		"=SUM(1-)":  "formula not valid",
		"=SUM(1*)":  "formula not valid",
		"=SUM(1/)":  "formula not valid",
		`=SUM("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// SUMIF
		"=SUMIF()": "SUMIF requires at least 2 argument",
		// SUMSQ
		`=SUMSQ("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// TAN
		"=TAN()":    "TAN requires 1 numeric argument",
		`=TAN("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// TANH
		"=TANH()":    "TANH requires 1 numeric argument",
		`=TANH("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// TRUNC
		"=TRUNC()":      "TRUNC requires at least 1 argument",
		`=TRUNC("X")`:   "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=TRUNC(1,"X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// Statistical Functions
		// MEDIAN
		"=MEDIAN()": "MEDIAN requires at least 1 argument",
		// Information Functions
		// ISBLANK
		"=ISBLANK(A1,A2)": "ISBLANK requires 1 argument",
		// ISERR
		"=ISERR()": "ISERR requires 1 argument",
		// ISERROR
		"=ISERROR()": "ISERROR requires 1 argument",
		// ISEVEN
		"=ISEVEN()":       "ISEVEN requires 1 argument",
		`=ISEVEN("text")`: "strconv.Atoi: parsing \"text\": invalid syntax",
		// ISNA
		"=ISNA()": "ISNA requires 1 argument",
		// ISNONTEXT
		"=ISNONTEXT()": "ISNONTEXT requires 1 argument",
		// ISNUMBER
		"=ISNUMBER()": "ISNUMBER requires 1 argument",
		// ISODD
		"=ISODD()":       "ISODD requires 1 argument",
		`=ISODD("text")`: "strconv.Atoi: parsing \"text\": invalid syntax",
		// NA
		"=NA(1)": "NA accepts no arguments",
		// Logical Functions
		// AND
		`=AND("text")`: "strconv.ParseFloat: parsing \"text\": invalid syntax",
		`=AND(A1:B1)`:  "#VALUE!",
		"=AND()":       "AND requires at least 1 argument",
		"=AND(1" + strings.Repeat(",1", 30) + ")": "AND accepts at most 30 arguments",
		// OR
		`=OR("text")`:                            "strconv.ParseFloat: parsing \"text\": invalid syntax",
		`=OR(A1:B1)`:                             "#VALUE!",
		"=OR()":                                  "OR requires at least 1 argument",
		"=OR(1" + strings.Repeat(",1", 30) + ")": "OR accepts at most 30 arguments",
		// Date and Time Functions
		// DATE
		"=DATE()":               "DATE requires 3 number arguments",
		`=DATE("text",10,21)`:   "DATE requires 3 number arguments",
		`=DATE(2020,"text",21)`: "DATE requires 3 number arguments",
		`=DATE(2020,10,"text")`: "DATE requires 3 number arguments",
		// Text Functions
		// CLEAN
		"=CLEAN()":    "CLEAN requires 1 argument",
		"=CLEAN(1,2)": "CLEAN requires 1 argument",
		// LEN
		"=LEN()": "LEN requires 1 string argument",
		// TRIM
		"=TRIM()":    "TRIM requires 1 argument",
		"=TRIM(1,2)": "TRIM requires 1 argument",
		// LOWER
		"=LOWER()":    "LOWER requires 1 argument",
		"=LOWER(1,2)": "LOWER requires 1 argument",
		// UPPER
		"=UPPER()":    "UPPER requires 1 argument",
		"=UPPER(1,2)": "UPPER requires 1 argument",
		// PROPER
		"=PROPER()":    "PROPER requires 1 argument",
		"=PROPER(1,2)": "PROPER requires 1 argument",
		// Conditional Functions
		// IF
		"=IF()":        "IF requires at least 1 argument",
		"=IF(0,1,2,3)": "IF accepts at most 3 arguments",
		"=IF(D1,1,2)":  "strconv.ParseBool: parsing \"Month\": invalid syntax",
		// Excel Lookup and Reference Functions
		// CHOOSE
		"=CHOOSE()":                "CHOOSE requires 2 arguments",
		"=CHOOSE(\"index_num\",0)": "CHOOSE requires first argument of type number",
		"=CHOOSE(2,0)":             "index_num should be <= to the number of values",
	}
	for formula, expected := range mathCalcError {
		f := prepareData()
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.EqualError(t, err, expected, formula)
		assert.Equal(t, "", result, formula)
	}

	referenceCalc := map[string]string{
		// MDETERM
		"=MDETERM(A1:B2)": "-3",
		// PRODUCT
		"=PRODUCT(Sheet1!A1:Sheet1!A1:A2,A2)": "4",
		// SUM
		"=A1/A3":                          "0.333333333333333",
		"=SUM(A1:A2)":                     "3",
		"=SUM(Sheet1!A1,A2)":              "3",
		"=(-2-SUM(-4+A2))*5":              "0",
		"=SUM(Sheet1!A1:Sheet1!A1:A2,A2)": "5",
		"=SUM(A1,A2,A3)*SUM(2,3)":         "30",
		"=1+SUM(SUM(A1+A2/A3)*(2-3),2)":   "1.333333333333334",
		"=A1/A2/SUM(A1:A2:B1)":            "0.041666666666667",
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

	volatileFuncs := []string{
		"=RAND()",
		"=RANDBETWEEN(1,2)",
	}
	for _, formula := range volatileFuncs {
		f := prepareData()
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		_, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err)
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
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestCalcCellValue.xlsx")))

}

func TestCalculate(t *testing.T) {
	err := `strconv.ParseFloat: parsing "string": invalid syntax`
	opd := NewStack()
	opd.Push(efp.Token{TValue: "string"})
	opt := efp.Token{TValue: "-", TType: efp.TokenTypeOperatorPrefix}
	assert.EqualError(t, calculate(opd, opt), err)
	opd.Push(efp.Token{TValue: "string"})
	opd.Push(efp.Token{TValue: "string"})
	opt = efp.Token{TValue: "-", TType: efp.TokenTypeOperatorInfix}
	assert.EqualError(t, calculate(opd, opt), err)
}

func TestCalcCellValueWithDefinedName(t *testing.T) {
	cellData := [][]interface{}{
		{"A1 value", "B1 value", nil},
	}
	prepareData := func() *File {
		f := NewFile()
		for r, row := range cellData {
			for c, value := range row {
				cell, _ := CoordinatesToCellName(c+1, r+1)
				assert.NoError(t, f.SetCellValue("Sheet1", cell, value))
			}
		}
		assert.NoError(t, f.SetDefinedName(&DefinedName{Name: "defined_name1", RefersTo: "Sheet1!A1", Scope: "Workbook"}))
		assert.NoError(t, f.SetDefinedName(&DefinedName{Name: "defined_name1", RefersTo: "Sheet1!B1", Scope: "Sheet1"}))

		return f
	}
	f := prepareData()
	assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "=defined_name1"))
	result, err := f.CalcCellValue("Sheet1", "C1")
	assert.NoError(t, err)
	// DefinedName with scope WorkSheet takes precedence over DefinedName with scope Workbook, so we should get B1 value
	assert.Equal(t, "B1 value", result, "=defined_name1")
}

func TestCalcPow(t *testing.T) {
	err := `strconv.ParseFloat: parsing "text": invalid syntax`
	assert.EqualError(t, calcPow("1", "text", nil), err)
	assert.EqualError(t, calcPow("text", "1", nil), err)
	assert.EqualError(t, calcL("1", "text", nil), err)
	assert.EqualError(t, calcL("text", "1", nil), err)
	assert.EqualError(t, calcLe("1", "text", nil), err)
	assert.EqualError(t, calcLe("text", "1", nil), err)
	assert.EqualError(t, calcG("1", "text", nil), err)
	assert.EqualError(t, calcG("text", "1", nil), err)
	assert.EqualError(t, calcGe("1", "text", nil), err)
	assert.EqualError(t, calcGe("text", "1", nil), err)
	assert.EqualError(t, calcAdd("1", "text", nil), err)
	assert.EqualError(t, calcAdd("text", "1", nil), err)
	assert.EqualError(t, calcAdd("1", "text", nil), err)
	assert.EqualError(t, calcAdd("text", "1", nil), err)
	assert.EqualError(t, calcSubtract("1", "text", nil), err)
	assert.EqualError(t, calcSubtract("text", "1", nil), err)
	assert.EqualError(t, calcMultiply("1", "text", nil), err)
	assert.EqualError(t, calcMultiply("text", "1", nil), err)
	assert.EqualError(t, calcDiv("1", "text", nil), err)
	assert.EqualError(t, calcDiv("text", "1", nil), err)
}

func TestISBLANK(t *testing.T) {
	argsList := list.New()
	argsList.PushBack(formulaArg{
		Type: ArgUnknown,
	})
	fn := formulaFuncs{}
	result := fn.ISBLANK(argsList)
	assert.Equal(t, result.String, "TRUE")
	assert.Empty(t, result.Error)
}

func TestAND(t *testing.T) {
	argsList := list.New()
	argsList.PushBack(formulaArg{
		Type: ArgUnknown,
	})
	fn := formulaFuncs{}
	result := fn.AND(argsList)
	assert.Equal(t, result.String, "TRUE")
	assert.Empty(t, result.Error)
}

func TestOR(t *testing.T) {
	argsList := list.New()
	argsList.PushBack(formulaArg{
		Type: ArgUnknown,
	})
	fn := formulaFuncs{}
	result := fn.OR(argsList)
	assert.Equal(t, result.String, "FALSE")
	assert.Empty(t, result.Error)
}

func TestDet(t *testing.T) {
	assert.Equal(t, det([][]float64{
		{1, 2, 3, 4},
		{2, 3, 4, 5},
		{3, 4, 5, 6},
		{4, 5, 6, 7},
	}), float64(0))
}
