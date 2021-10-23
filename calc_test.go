package excelize

import (
	"container/list"
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/xuri/efp"
)

func prepareCalcData(cellData [][]interface{}) *File {
	f := NewFile()
	for r, row := range cellData {
		for c, value := range row {
			cell, _ := CoordinatesToCellName(c+1, r+1)
			_ = f.SetCellValue("Sheet1", cell, value)
		}
	}
	return f
}

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
	mathCalc := map[string]string{
		"=2^3":            "8",
		"=1=1":            "TRUE",
		"=1=2":            "FALSE",
		"=1<2":            "TRUE",
		"=3<2":            "FALSE",
		"=1<\"-1\"":       "TRUE",
		"=\"-1\"<1":       "FALSE",
		"=\"-1\"<\"-2\"":  "TRUE",
		"=2<=3":           "TRUE",
		"=2<=1":           "FALSE",
		"=1<=\"-1\"":      "TRUE",
		"=\"-1\"<=1":      "FALSE",
		"=\"-1\"<=\"-2\"": "TRUE",
		"=2>1":            "TRUE",
		"=2>3":            "FALSE",
		"=1>\"-1\"":       "FALSE",
		"=\"-1\">-1":      "TRUE",
		"=\"-1\">\"-2\"":  "FALSE",
		"=2>=1":           "TRUE",
		"=2>=3":           "FALSE",
		"=1>=\"-1\"":      "FALSE",
		"=\"-1\">=-1":     "TRUE",
		"=\"-1\">=\"-2\"": "FALSE",
		"=1&2":            "12",
		"=15%":            "0.15",
		"=1+20%":          "1.2",
		`="A"="A"`:        "TRUE",
		`="A"<>"A"`:       "FALSE",
		// Engineering Functions
		// BESSELI
		"=BESSELI(4.5,1)": "15.389222753735925",
		"=BESSELI(32,1)":  "5.502845511211247e+12",
		// BESSELJ
		"=BESSELJ(1.9,2)": "0.329925727692387",
		// BESSELK
		"=BESSELK(0.05,0)": "3.11423403428966",
		"=BESSELK(0.05,1)": "19.90967432724863",
		"=BESSELK(0.05,2)": "799.501207124235",
		"=BESSELK(3,2)":    "0.0615104585619118",
		// BESSELY
		"=BESSELY(0.05,0)": "-1.97931100684153",
		"=BESSELY(0.05,1)": "-12.789855163794034",
		"=BESSELY(0.05,2)": "-509.61489554491976",
		"=BESSELY(9,2)":    "-0.229082087487741",
		// BIN2DEC
		"=BIN2DEC(\"10\")":         "2",
		"=BIN2DEC(\"11\")":         "3",
		"=BIN2DEC(\"0000000010\")": "2",
		"=BIN2DEC(\"1111111110\")": "-2",
		"=BIN2DEC(\"110\")":        "6",
		// BIN2HEX
		"=BIN2HEX(\"10\")":         "2",
		"=BIN2HEX(\"0000000001\")": "1",
		"=BIN2HEX(\"10\",10)":      "0000000002",
		"=BIN2HEX(\"1111111110\")": "FFFFFFFFFE",
		"=BIN2HEX(\"11101\")":      "1D",
		// BIN2OCT
		"=BIN2OCT(\"101\")":        "5",
		"=BIN2OCT(\"0000000001\")": "1",
		"=BIN2OCT(\"10\",10)":      "0000000002",
		"=BIN2OCT(\"1111111110\")": "7777777776",
		"=BIN2OCT(\"1110\")":       "16",
		// BITAND
		"=BITAND(13,14)": "12",
		// BITLSHIFT
		"=BITLSHIFT(5,2)": "20",
		"=BITLSHIFT(3,5)": "96",
		// BITOR
		"=BITOR(9,12)": "13",
		// BITRSHIFT
		"=BITRSHIFT(20,2)": "5",
		"=BITRSHIFT(52,4)": "3",
		// BITXOR
		"=BITXOR(5,6)":  "3",
		"=BITXOR(9,12)": "5",
		// COMPLEX
		"=COMPLEX(5,2)":         "5+2i",
		"=COMPLEX(5,-9)":        "5-9i",
		"=COMPLEX(-1,2,\"j\")":  "-1+2j",
		"=COMPLEX(10,-5,\"i\")": "10-5i",
		"=COMPLEX(0,5)":         "5i",
		"=COMPLEX(3,0)":         "3",
		"=COMPLEX(0,-2)":        "-2i",
		"=COMPLEX(0,0)":         "0",
		"=COMPLEX(0,-1,\"j\")":  "-j",
		// DEC2BIN
		"=DEC2BIN(2)":    "10",
		"=DEC2BIN(3)":    "11",
		"=DEC2BIN(2,10)": "0000000010",
		"=DEC2BIN(-2)":   "1111111110",
		"=DEC2BIN(6)":    "110",
		// DEC2HEX
		"=DEC2HEX(10)":    "A",
		"=DEC2HEX(31)":    "1F",
		"=DEC2HEX(16,10)": "0000000010",
		"=DEC2HEX(-16)":   "FFFFFFFFF0",
		"=DEC2HEX(273)":   "111",
		// DEC2OCT
		"=DEC2OCT(8)":    "10",
		"=DEC2OCT(18)":   "22",
		"=DEC2OCT(8,10)": "0000000010",
		"=DEC2OCT(-8)":   "7777777770",
		"=DEC2OCT(237)":  "355",
		// HEX2BIN
		"=HEX2BIN(\"2\")":          "10",
		"=HEX2BIN(\"0000000001\")": "1",
		"=HEX2BIN(\"2\",10)":       "0000000010",
		"=HEX2BIN(\"F0\")":         "11110000",
		"=HEX2BIN(\"1D\")":         "11101",
		// HEX2DEC
		"=HEX2DEC(\"A\")":          "10",
		"=HEX2DEC(\"1F\")":         "31",
		"=HEX2DEC(\"0000000010\")": "16",
		"=HEX2DEC(\"FFFFFFFFF0\")": "-16",
		"=HEX2DEC(\"111\")":        "273",
		"=HEX2DEC(\"\")":           "0",
		// HEX2OCT
		"=HEX2OCT(\"A\")":          "12",
		"=HEX2OCT(\"000000000F\")": "17",
		"=HEX2OCT(\"8\",10)":       "0000000010",
		"=HEX2OCT(\"FFFFFFFFF8\")": "7777777770",
		"=HEX2OCT(\"1F3\")":        "763",
		// IMABS
		"=IMABS(\"2j\")":              "2",
		"=IMABS(\"-1+2i\")":           "2.23606797749979",
		"=IMABS(COMPLEX(-1,2,\"j\"))": "2.23606797749979",
		// IMAGINARY
		"=IMAGINARY(\"5+2i\")": "2",
		"=IMAGINARY(\"2-i\")":  "-1",
		"=IMAGINARY(6)":        "0",
		"=IMAGINARY(\"3i\")":   "3",
		"=IMAGINARY(\"4+i\")":  "1",
		// IMARGUMENT
		"=IMARGUMENT(\"5+2i\")": "0.380506377112365",
		"=IMARGUMENT(\"2-i\")":  "-0.463647609000806",
		"=IMARGUMENT(6)":        "0",
		// IMCONJUGATE
		"=IMCONJUGATE(\"5+2i\")": "5-2i",
		"=IMCONJUGATE(\"2-i\")":  "2+i",
		"=IMCONJUGATE(6)":        "6",
		"=IMCONJUGATE(\"3i\")":   "-3i",
		"=IMCONJUGATE(\"4+i\")":  "4-i",
		// IMCOS
		"=IMCOS(0)":          "1",
		"=IMCOS(0.5)":        "0.877582561890373",
		"=IMCOS(\"3+0.5i\")": "-1.1163412445261518-0.0735369737112366i",
		// IMCOSH
		"=IMCOSH(0.5)":           "1.12762596520638",
		"=IMCOSH(\"3+0.5i\")":    "8.835204606500994+4.802825082743033i",
		"=IMCOSH(\"2-i\")":       "2.0327230070196656-3.0518977991518i",
		"=IMCOSH(COMPLEX(1,-1))": "0.8337300251311491-0.9888977057628651i",
		// IMCOT
		"=IMCOT(0.5)":           "1.830487721712452",
		"=IMCOT(\"3+0.5i\")":    "-0.4793455787473728-2.016092521506228i",
		"=IMCOT(\"2-i\")":       "-0.171383612909185+0.8213297974938518i",
		"=IMCOT(COMPLEX(1,-1))": "0.21762156185440268+0.868014142895925i",
		// IMCSC
		"=IMCSC(\"j\")": "-0.8509181282393216i",
		// IMCSCH
		"=IMCSCH(COMPLEX(1,-1))": "0.30393100162842646+0.6215180171704284i",
		// IMDIV
		"=IMDIV(\"5+2i\",\"1+i\")":          "3.5-1.5i",
		"=IMDIV(\"2+2i\",\"2+i\")":          "1.2+0.4i",
		"=IMDIV(COMPLEX(5,2),COMPLEX(0,1))": "2-5i",
		// IMEXP
		"=IMEXP(0)":             "1",
		"=IMEXP(0.5)":           "1.64872127070013",
		"=IMEXP(\"1-2i\")":      "-1.1312043837568135-2.4717266720048183i",
		"=IMEXP(COMPLEX(1,-1))": "1.4686939399158851-2.2873552871788423i",
		// IMLN
		"=IMLN(0.5)":           "-0.693147180559945",
		"=IMLN(\"3+0.5i\")":    "1.1123117757621668+0.16514867741462683i",
		"=IMLN(\"2-i\")":       "0.8047189562170503-0.4636476090008061i",
		"=IMLN(COMPLEX(1,-1))": "0.3465735902799727-0.7853981633974483i",
		// IMLOG10
		"=IMLOG10(0.5)":           "-0.301029995663981",
		"=IMLOG10(\"3+0.5i\")":    "0.48307086636951624+0.07172315929479262i",
		"=IMLOG10(\"2-i\")":       "0.34948500216800943-0.20135959813668655i",
		"=IMLOG10(COMPLEX(1,-1))": "0.1505149978319906-0.3410940884604603i",
		// IMREAL
		"=IMREAL(\"5+2i\")":     "5",
		"=IMREAL(\"2+2i\")":     "2",
		"=IMREAL(6)":            "6",
		"=IMREAL(\"3i\")":       "0",
		"=IMREAL(COMPLEX(4,1))": "4",
		// IMSEC
		"=IMSEC(0.5)":           "1.139493927324549",
		"=IMSEC(\"3+0.5i\")":    "-0.8919131797403304+0.05875317818173977i",
		"=IMSEC(\"2-i\")":       "-0.4131493442669401-0.687527438655479i",
		"=IMSEC(COMPLEX(1,-1))": "0.49833703055518686-0.5910838417210451i",
		// IMSECH
		"=IMSECH(0.5)":           "0.886818883970074",
		"=IMSECH(\"3+0.5i\")":    "0.08736657796213027-0.047492549490160664i",
		"=IMSECH(\"2-i\")":       "0.1511762982655772+0.22697367539372157i",
		"=IMSECH(COMPLEX(1,-1))": "0.49833703055518686+0.5910838417210451i",
		// IMSIN
		"=IMSIN(0.5)":           "0.479425538604203",
		"=IMSIN(\"3+0.5i\")":    "0.15913058529843999-0.5158804424525267i",
		"=IMSIN(\"2-i\")":       "1.4031192506220405+0.4890562590412937i",
		"=IMSIN(COMPLEX(1,-1))": "1.2984575814159773-0.6349639147847361i",
		// IMSINH
		"=IMSINH(-0)":            "0",
		"=IMSINH(0.5)":           "0.521095305493747",
		"=IMSINH(\"3+0.5i\")":    "8.791512343493714+4.82669427481082i",
		"=IMSINH(\"2-i\")":       "1.9596010414216063-3.165778513216168i",
		"=IMSINH(COMPLEX(1,-1))": "0.6349639147847361-1.2984575814159773i",
		// IMSQRT
		"=IMSQRT(\"i\")":     "0.7071067811865476+0.7071067811865476i",
		"=IMSQRT(\"2-i\")":   "1.455346690225355-0.34356074972251244i",
		"=IMSQRT(\"5+2i\")":  "2.27872385417085+0.4388421169022545i",
		"=IMSQRT(6)":         "2.449489742783178",
		"=IMSQRT(\"-2-4i\")": "1.1117859405028423-1.7989074399478673i",
		// IMSUB
		"=IMSUB(\"5+i\",\"1+4i\")":          "4-3i",
		"=IMSUB(\"9+2i\",6)":                "3+2i",
		"=IMSUB(COMPLEX(5,2),COMPLEX(0,1))": "5+i",
		// IMSUM
		"=IMSUM(\"1-i\",\"5+10i\",2)":       "8+9i",
		"=IMSUM(COMPLEX(5,2),COMPLEX(0,1))": "5+3i",
		// IMTAN
		"=IMTAN(-0)":            "0",
		"=IMTAN(0.5)":           "0.54630248984379",
		"=IMTAN(\"3+0.5i\")":    "-0.11162105077158344+0.46946999342588536i",
		"=IMTAN(\"2-i\")":       "-0.24345820118572523-1.16673625724092i",
		"=IMTAN(COMPLEX(1,-1))": "0.2717525853195117-1.0839233273386948i",
		// OCT2BIN
		"=OCT2BIN(\"5\")":          "101",
		"=OCT2BIN(\"0000000001\")": "1",
		"=OCT2BIN(\"2\",10)":       "0000000010",
		"=OCT2BIN(\"7777777770\")": "1111111000",
		"=OCT2BIN(\"16\")":         "1110",
		// OCT2DEC
		"=OCT2DEC(\"10\")":         "8",
		"=OCT2DEC(\"22\")":         "18",
		"=OCT2DEC(\"0000000010\")": "8",
		"=OCT2DEC(\"7777777770\")": "-8",
		"=OCT2DEC(\"355\")":        "237",
		// OCT2HEX
		"=OCT2HEX(\"10\")":         "8",
		"=OCT2HEX(\"0000000007\")": "7",
		"=OCT2HEX(\"10\",10)":      "0000000008",
		"=OCT2HEX(\"7777777770\")": "FFFFFFFFF8",
		"=OCT2HEX(\"763\")":        "1F3",
		// Math and Trigonometric Functions
		// ABS
		"=ABS(-1)":      "1",
		"=ABS(-6.5)":    "6.5",
		"=ABS(6.5)":     "6.5",
		"=ABS(0)":       "0",
		"=ABS(2-4.5)":   "2.5",
		"=ABS(ABS(-1))": "1",
		// ACOS
		"=ACOS(-1)":     "3.141592653589793",
		"=ACOS(0)":      "1.5707963267949",
		"=ACOS(ABS(0))": "1.5707963267949",
		// ACOSH
		"=ACOSH(1)":        "0",
		"=ACOSH(2.5)":      "1.566799236972411",
		"=ACOSH(5)":        "2.29243166956118",
		"=ACOSH(ACOSH(5))": "1.47138332153668",
		// ACOT
		"=_xlfn.ACOT(1)":             "0.785398163397448",
		"=_xlfn.ACOT(-2)":            "2.677945044588987",
		"=_xlfn.ACOT(0)":             "1.5707963267949",
		"=_xlfn.ACOT(_xlfn.ACOT(0))": "0.566911504941009",
		// ACOTH
		"=_xlfn.ACOTH(-5)":      "-0.202732554054082",
		"=_xlfn.ACOTH(1.1)":     "1.52226121886171",
		"=_xlfn.ACOTH(2)":       "0.549306144334055",
		"=_xlfn.ACOTH(ABS(-2))": "0.549306144334055",
		// ARABIC
		"=_xlfn.ARABIC(\"IV\")":       "4",
		"=_xlfn.ARABIC(\"-IV\")":      "-4",
		"=_xlfn.ARABIC(\"MCXX\")":     "1120",
		"=_xlfn.ARABIC(\"\")":         "0",
		"=_xlfn.ARABIC(\" ll  lc \")": "-50",
		// ASIN
		"=ASIN(-1)":      "-1.5707963267949",
		"=ASIN(0)":       "0",
		"=ASIN(ASIN(0))": "0",
		// ASINH
		"=ASINH(0)":        "0",
		"=ASINH(-0.5)":     "-0.481211825059603",
		"=ASINH(2)":        "1.44363547517881",
		"=ASINH(ASINH(0))": "0",
		// ATAN
		"=ATAN(-1)":      "-0.785398163397448",
		"=ATAN(0)":       "0",
		"=ATAN(1)":       "0.785398163397448",
		"=ATAN(ATAN(0))": "0",
		// ATANH
		"=ATANH(-0.8)":     "-1.09861228866811",
		"=ATANH(0)":        "0",
		"=ATANH(0.5)":      "0.549306144334055",
		"=ATANH(ATANH(0))": "0",
		// ATAN2
		"=ATAN2(1,1)":          "0.785398163397448",
		"=ATAN2(1,-1)":         "-0.785398163397448",
		"=ATAN2(4,0)":          "0",
		"=ATAN2(4,ATAN2(4,0))": "0",
		// BASE
		"=BASE(12,2)":          "1100",
		"=BASE(12,2,8)":        "00001100",
		"=BASE(100000,16)":     "186A0",
		"=BASE(BASE(12,2),16)": "44C",
		// CEILING
		"=CEILING(22.25,0.1)":              "22.3",
		"=CEILING(22.25,0.5)":              "22.5",
		"=CEILING(22.25,1)":                "23",
		"=CEILING(22.25,10)":               "30",
		"=CEILING(22.25,20)":               "40",
		"=CEILING(-22.25,-0.1)":            "-22.3",
		"=CEILING(-22.25,-1)":              "-23",
		"=CEILING(-22.25,-5)":              "-25",
		"=CEILING(22.25)":                  "23",
		"=CEILING(CEILING(22.25,0.1),0.1)": "22.3",
		// _xlfn.CEILING.MATH
		"=_xlfn.CEILING.MATH(15.25,1)":                       "16",
		"=_xlfn.CEILING.MATH(15.25,0.1)":                     "15.3",
		"=_xlfn.CEILING.MATH(15.25,5)":                       "20",
		"=_xlfn.CEILING.MATH(-15.25,1)":                      "-15",
		"=_xlfn.CEILING.MATH(-15.25,1,1)":                    "-15", // should be 16
		"=_xlfn.CEILING.MATH(-15.25,10)":                     "-10",
		"=_xlfn.CEILING.MATH(-15.25)":                        "-15",
		"=_xlfn.CEILING.MATH(-15.25,-5,-1)":                  "-10",
		"=_xlfn.CEILING.MATH(_xlfn.CEILING.MATH(15.25,1),1)": "16",
		// _xlfn.CEILING.PRECISE
		"=_xlfn.CEILING.PRECISE(22.25,0.1)":                          "22.3",
		"=_xlfn.CEILING.PRECISE(22.25,0.5)":                          "22.5",
		"=_xlfn.CEILING.PRECISE(22.25,1)":                            "23",
		"=_xlfn.CEILING.PRECISE(22.25)":                              "23",
		"=_xlfn.CEILING.PRECISE(22.25,10)":                           "30",
		"=_xlfn.CEILING.PRECISE(22.25,0)":                            "0",
		"=_xlfn.CEILING.PRECISE(-22.25,1)":                           "-22",
		"=_xlfn.CEILING.PRECISE(-22.25,-1)":                          "-22",
		"=_xlfn.CEILING.PRECISE(-22.25,5)":                           "-20",
		"=_xlfn.CEILING.PRECISE(_xlfn.CEILING.PRECISE(22.25,0.1),5)": "25",
		// COMBIN
		"=COMBIN(6,1)":           "6",
		"=COMBIN(6,2)":           "15",
		"=COMBIN(6,3)":           "20",
		"=COMBIN(6,4)":           "15",
		"=COMBIN(6,5)":           "6",
		"=COMBIN(6,6)":           "1",
		"=COMBIN(0,0)":           "1",
		"=COMBIN(6,COMBIN(0,0))": "6",
		// _xlfn.COMBINA
		"=_xlfn.COMBINA(6,1)":                  "6",
		"=_xlfn.COMBINA(6,2)":                  "21",
		"=_xlfn.COMBINA(6,3)":                  "56",
		"=_xlfn.COMBINA(6,4)":                  "126",
		"=_xlfn.COMBINA(6,5)":                  "252",
		"=_xlfn.COMBINA(6,6)":                  "462",
		"=_xlfn.COMBINA(0,0)":                  "0",
		"=_xlfn.COMBINA(0,_xlfn.COMBINA(0,0))": "0",
		// COS
		"=COS(0.785398163)": "0.707106781467586",
		"=COS(0)":           "1",
		"=COS(COS(0))":      "0.54030230586814",
		// COSH
		"=COSH(0)":       "1",
		"=COSH(0.5)":     "1.12762596520638",
		"=COSH(-2)":      "3.76219569108363",
		"=COSH(COSH(0))": "1.54308063481524",
		// _xlfn.COT
		"=_xlfn.COT(0.785398163397448)": "1",
		"=_xlfn.COT(_xlfn.COT(0.45))":   "-0.545473116787229",
		// _xlfn.COTH
		"=_xlfn.COTH(-3.14159265358979)": "-1.00374187319732",
		"=_xlfn.COTH(_xlfn.COTH(1))":     "1.15601401811395",
		// _xlfn.CSC
		"=_xlfn.CSC(-6)":              "3.57889954725441",
		"=_xlfn.CSC(1.5707963267949)": "1",
		"=_xlfn.CSC(_xlfn.CSC(1))":    "1.07785184031088",
		// _xlfn.CSCH
		"=_xlfn.CSCH(-3.14159265358979)": "-0.0865895375300472",
		"=_xlfn.CSCH(_xlfn.CSCH(1))":     "1.04451010395518",
		// _xlfn.DECIMAL
		`=_xlfn.DECIMAL("1100",2)`:    "12",
		`=_xlfn.DECIMAL("186A0",16)`:  "100000",
		`=_xlfn.DECIMAL("31L0",32)`:   "100000",
		`=_xlfn.DECIMAL("70122",8)`:   "28754",
		`=_xlfn.DECIMAL("0x70122",8)`: "28754",
		// DEGREES
		"=DEGREES(1)":          "57.29577951308232",
		"=DEGREES(2.5)":        "143.2394487827058",
		"=DEGREES(DEGREES(1))": "3282.806350011744",
		// EVEN
		"=EVEN(23)":   "24",
		"=EVEN(2.22)": "4",
		"=EVEN(0)":    "0",
		"=EVEN(-0.3)": "-2",
		"=EVEN(-11)":  "-12",
		"=EVEN(-4)":   "-4",
		"=EVEN((0))":  "0",
		// EXP
		"=EXP(100)":    "2.6881171418161356E+43",
		"=EXP(0.1)":    "1.10517091807565",
		"=EXP(0)":      "1",
		"=EXP(-5)":     "0.00673794699908547",
		"=EXP(EXP(0))": "2.718281828459045",
		// FACT
		"=FACT(3)":       "6",
		"=FACT(6)":       "720",
		"=FACT(10)":      "3.6288e+06",
		"=FACT(FACT(3))": "720",
		// FACTDOUBLE
		"=FACTDOUBLE(5)":             "15",
		"=FACTDOUBLE(8)":             "384",
		"=FACTDOUBLE(13)":            "135135",
		"=FACTDOUBLE(FACTDOUBLE(1))": "1",
		// FLOOR
		"=FLOOR(26.75,0.1)":        "26.700000000000003",
		"=FLOOR(26.75,0.5)":        "26.5",
		"=FLOOR(26.75,1)":          "26",
		"=FLOOR(26.75,10)":         "20",
		"=FLOOR(26.75,20)":         "20",
		"=FLOOR(-26.75,-0.1)":      "-26.700000000000003",
		"=FLOOR(-26.75,-1)":        "-26",
		"=FLOOR(-26.75,-5)":        "-25",
		"=FLOOR(FLOOR(26.75,1),1)": "26",
		// _xlfn.FLOOR.MATH
		"=_xlfn.FLOOR.MATH(58.55)":                  "58",
		"=_xlfn.FLOOR.MATH(58.55,0.1)":              "58.5",
		"=_xlfn.FLOOR.MATH(58.55,5)":                "55",
		"=_xlfn.FLOOR.MATH(58.55,1,1)":              "58",
		"=_xlfn.FLOOR.MATH(-58.55,1)":               "-59",
		"=_xlfn.FLOOR.MATH(-58.55,1,-1)":            "-58",
		"=_xlfn.FLOOR.MATH(-58.55,1,1)":             "-59", // should be -58
		"=_xlfn.FLOOR.MATH(-58.55,10)":              "-60",
		"=_xlfn.FLOOR.MATH(_xlfn.FLOOR.MATH(1),10)": "0",
		// _xlfn.FLOOR.PRECISE
		"=_xlfn.FLOOR.PRECISE(26.75,0.1)":                     "26.700000000000003",
		"=_xlfn.FLOOR.PRECISE(26.75,0.5)":                     "26.5",
		"=_xlfn.FLOOR.PRECISE(26.75,1)":                       "26",
		"=_xlfn.FLOOR.PRECISE(26.75)":                         "26",
		"=_xlfn.FLOOR.PRECISE(26.75,10)":                      "20",
		"=_xlfn.FLOOR.PRECISE(26.75,0)":                       "0",
		"=_xlfn.FLOOR.PRECISE(-26.75,1)":                      "-27",
		"=_xlfn.FLOOR.PRECISE(-26.75,-1)":                     "-27",
		"=_xlfn.FLOOR.PRECISE(-26.75,-5)":                     "-30",
		"=_xlfn.FLOOR.PRECISE(_xlfn.FLOOR.PRECISE(26.75),-5)": "25",
		// GCD
		"=GCD(0)":        "0",
		"=GCD(1,0)":      "1",
		"=GCD(1,5)":      "1",
		"=GCD(15,10,25)": "5",
		"=GCD(0,8,12)":   "4",
		"=GCD(7,2)":      "1",
		"=GCD(1,GCD(1))": "1",
		// INT
		"=INT(100.9)":  "100",
		"=INT(5.22)":   "5",
		"=INT(5.99)":   "5",
		"=INT(-6.1)":   "-7",
		"=INT(-100.9)": "-101",
		"=INT(INT(0))": "0",
		// ISO.CEILING
		"=ISO.CEILING(22.25)":              "23",
		"=ISO.CEILING(22.25,1)":            "23",
		"=ISO.CEILING(22.25,0.1)":          "22.3",
		"=ISO.CEILING(22.25,10)":           "30",
		"=ISO.CEILING(-22.25,1)":           "-22",
		"=ISO.CEILING(-22.25,0.1)":         "-22.200000000000003",
		"=ISO.CEILING(-22.25,5)":           "-20",
		"=ISO.CEILING(-22.25,0)":           "0",
		"=ISO.CEILING(1,ISO.CEILING(1,0))": "0",
		// LCM
		"=LCM(1,5)":        "5",
		"=LCM(15,10,25)":   "150",
		"=LCM(1,8,12)":     "24",
		"=LCM(7,2)":        "14",
		"=LCM(7)":          "7",
		`=LCM("",1)`:       "1",
		`=LCM(0,0)`:        "0",
		`=LCM(0,LCM(0,0))`: "0",
		// LN
		"=LN(1)":       "0",
		"=LN(100)":     "4.605170185988092",
		"=LN(0.5)":     "-0.693147180559945",
		"=LN(LN(100))": "1.5271796258079",
		// LOG
		"=LOG(64,2)":     "6",
		"=LOG(100)":      "2",
		"=LOG(4,0.5)":    "-2",
		"=LOG(500)":      "2.69897000433602",
		"=LOG(LOG(100))": "0.301029995663981",
		// LOG10
		"=LOG10(100)":        "2",
		"=LOG10(1000)":       "3",
		"=LOG10(0.001)":      "-3",
		"=LOG10(25)":         "1.39794000867204",
		"=LOG10(LOG10(100))": "0.301029995663981",
		// IMLOG2
		"=IMLOG2(\"5+2i\")": "2.4289904975637864+0.5489546632866347i",
		"=IMLOG2(\"2-i\")":  "1.1609640474436813-0.6689021062254881i",
		"=IMLOG2(6)":        "2.584962500721156",
		"=IMLOG2(\"3i\")":   "1.584962500721156+2.266180070913597i",
		"=IMLOG2(\"4+i\")":  "2.04373142062517+0.3534295024167349i",
		// IMPOWER
		"=IMPOWER(\"2-i\",2)":   "3.000000000000001-4i",
		"=IMPOWER(\"2-i\",3)":   "2.0000000000000018-11.000000000000002i",
		"=IMPOWER(9,0.5)":       "3",
		"=IMPOWER(\"2+4i\",-2)": "-0.029999999999999985-0.039999999999999994i",
		// IMPRODUCT
		"=IMPRODUCT(3,6)":                       "18",
		`=IMPRODUCT("",3,SUM(6))`:               "18",
		"=IMPRODUCT(\"1-i\",\"5+10i\",2)":       "30+10i",
		"=IMPRODUCT(COMPLEX(5,2),COMPLEX(0,1))": "-2+5i",
		"=IMPRODUCT(A1:C1)":                     "4",
		// MOD
		"=MOD(6,4)":        "2",
		"=MOD(6,3)":        "0",
		"=MOD(6,2.5)":      "1",
		"=MOD(6,1.333)":    "0.668",
		"=MOD(-10.23,1)":   "0.77",
		"=MOD(MOD(1,1),1)": "0",
		// MROUND
		"=MROUND(333.7,0.5)":     "333.5",
		"=MROUND(333.8,1)":       "334",
		"=MROUND(333.3,2)":       "334",
		"=MROUND(555.3,400)":     "400",
		"=MROUND(555,1000)":      "1000",
		"=MROUND(-555.7,-1)":     "-556",
		"=MROUND(-555.4,-1)":     "-555",
		"=MROUND(-1555,-1000)":   "-2000",
		"=MROUND(MROUND(1,1),1)": "1",
		// MULTINOMIAL
		"=MULTINOMIAL(3,1,2,5)":        "27720",
		`=MULTINOMIAL("",3,1,2,5)`:     "27720",
		"=MULTINOMIAL(MULTINOMIAL(1))": "1",
		// _xlfn.MUNIT
		"=_xlfn.MUNIT(4)": "",
		// ODD
		"=ODD(22)":     "23",
		"=ODD(1.22)":   "3",
		"=ODD(1.22+4)": "7",
		"=ODD(0)":      "1",
		"=ODD(-1.3)":   "-3",
		"=ODD(-10)":    "-11",
		"=ODD(-3)":     "-3",
		"=ODD(ODD(1))": "1",
		// PI
		"=PI()": "3.141592653589793",
		// POWER
		"=POWER(4,2)":          "16",
		"=POWER(4,POWER(1,1))": "4",
		// PRODUCT
		"=PRODUCT(3,6)":            "18",
		`=PRODUCT("",3,6)`:         "18",
		`=PRODUCT(PRODUCT(1),3,6)`: "18",
		// QUOTIENT
		"=QUOTIENT(5,2)":             "2",
		"=QUOTIENT(4.5,3.1)":         "1",
		"=QUOTIENT(-10,3)":           "-3",
		"=QUOTIENT(QUOTIENT(1,2),3)": "0",
		// RADIANS
		"=RADIANS(50)":           "0.872664625997165",
		"=RADIANS(-180)":         "-3.141592653589793",
		"=RADIANS(180)":          "3.141592653589793",
		"=RADIANS(360)":          "6.283185307179586",
		"=RADIANS(RADIANS(360))": "0.109662271123215",
		// ROMAN
		"=ROMAN(499,0)":       "CDXCIX",
		"=ROMAN(1999,0)":      "MCMXCIX",
		"=ROMAN(1999,1)":      "MLMVLIV",
		"=ROMAN(1999,2)":      "MXMIX",
		"=ROMAN(1999,3)":      "MVMIV",
		"=ROMAN(1999,4)":      "MIM",
		"=ROMAN(1999,-1)":     "MCMXCIX",
		"=ROMAN(1999,5)":      "MIM",
		"=ROMAN(1999,ODD(1))": "MLMVLIV",
		// ROUND
		"=ROUND(100.319,1)":       "100.30000000000001",
		"=ROUND(5.28,1)":          "5.300000000000001",
		"=ROUND(5.9999,3)":        "6.000000000000002",
		"=ROUND(99.5,0)":          "100",
		"=ROUND(-6.3,0)":          "-6",
		"=ROUND(-100.5,0)":        "-101",
		"=ROUND(-22.45,1)":        "-22.5",
		"=ROUND(999,-1)":          "1000",
		"=ROUND(991,-1)":          "990",
		"=ROUND(ROUND(100,1),-1)": "100",
		// ROUNDDOWN
		"=ROUNDDOWN(99.999,1)":            "99.9",
		"=ROUNDDOWN(99.999,2)":            "99.99000000000002",
		"=ROUNDDOWN(99.999,0)":            "99",
		"=ROUNDDOWN(99.999,-1)":           "90",
		"=ROUNDDOWN(-99.999,2)":           "-99.99000000000002",
		"=ROUNDDOWN(-99.999,-1)":          "-90",
		"=ROUNDDOWN(ROUNDDOWN(100,1),-1)": "100",
		// ROUNDUP`
		"=ROUNDUP(11.111,1)":          "11.200000000000001",
		"=ROUNDUP(11.111,2)":          "11.120000000000003",
		"=ROUNDUP(11.111,0)":          "12",
		"=ROUNDUP(11.111,-1)":         "20",
		"=ROUNDUP(-11.111,2)":         "-11.120000000000003",
		"=ROUNDUP(-11.111,-1)":        "-20",
		"=ROUNDUP(ROUNDUP(100,1),-1)": "100",
		// SEC
		"=_xlfn.SEC(-3.14159265358979)": "-1",
		"=_xlfn.SEC(0)":                 "1",
		"=_xlfn.SEC(_xlfn.SEC(0))":      "0.54030230586814",
		// SECH
		"=_xlfn.SECH(-3.14159265358979)": "0.0862667383340547",
		"=_xlfn.SECH(0)":                 "1",
		"=_xlfn.SECH(_xlfn.SECH(0))":     "0.648054273663885",
		// SIGN
		"=SIGN(9.5)":        "1",
		"=SIGN(-9.5)":       "-1",
		"=SIGN(0)":          "0",
		"=SIGN(0.00000001)": "1",
		"=SIGN(6-7)":        "-1",
		"=SIGN(SIGN(-1))":   "-1",
		// SIN
		"=SIN(0.785398163)": "0.707106780905509",
		"=SIN(SIN(1))":      "0.745624141665558",
		// SINH
		"=SINH(0)":       "0",
		"=SINH(0.5)":     "0.521095305493747",
		"=SINH(-2)":      "-3.626860407847019",
		"=SINH(SINH(0))": "0",
		// SQRT
		"=SQRT(4)":        "2",
		"=SQRT(SQRT(16))": "2",
		// SQRTPI
		"=SQRTPI(5)":         "3.963327297606011",
		"=SQRTPI(0.2)":       "0.792665459521202",
		"=SQRTPI(100)":       "17.72453850905516",
		"=SQRTPI(0)":         "0",
		"=SQRTPI(SQRTPI(0))": "0",
		// STDEV
		"=STDEV(F2:F9)":         "10724.978287523809",
		"=STDEV(MUNIT(2))":      "0.577350269189626",
		"=STDEV(0,INT(0))":      "0",
		"=STDEV(INT(1),INT(1))": "0",
		// STDEV.S
		"=STDEV.S(F2:F9)": "10724.978287523809",
		// STDEVA
		"=STDEVA(F2:F9)":    "10724.978287523809",
		"=STDEVA(MUNIT(2))": "0.577350269189626",
		"=STDEVA(0,INT(0))": "0",
		// POISSON.DIST
		"=POISSON.DIST(20,25,FALSE)": "0.0519174686084913",
		"=POISSON.DIST(35,40,TRUE)":  "0.242414197690103",
		// POISSON
		"=POISSON(20,25,FALSE)": "0.0519174686084913",
		"=POISSON(35,40,TRUE)":  "0.242414197690103",
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
		`=SUMSQ(1,SUMSQ(1))`:       "2",
		"=SUMSQ(MUNIT(3))":         "0",
		// TAN
		"=TAN(1.047197551)": "1.732050806782486",
		"=TAN(0)":           "0",
		"=TAN(TAN(0))":      "0",
		// TANH
		"=TANH(0)":       "0",
		"=TANH(0.5)":     "0.46211715726001",
		"=TANH(-2)":      "-0.964027580075817",
		"=TANH(TANH(0))": "0",
		// TRUNC
		"=TRUNC(99.999,1)":    "99.9",
		"=TRUNC(99.999,2)":    "99.99",
		"=TRUNC(99.999)":      "99",
		"=TRUNC(99.999,-1)":   "90",
		"=TRUNC(-99.999,2)":   "-99.99",
		"=TRUNC(-99.999,-1)":  "-90",
		"=TRUNC(TRUNC(1),-1)": "0",
		// Statistical Functions
		// AVEDEV
		"=AVEDEV(1,2)":          "0.5",
		"=AVERAGE(A1:A4,B1:B4)": "2.5",
		// AVERAGE
		"=AVERAGE(INT(1))": "1",
		"=AVERAGE(A1)":     "1",
		"=AVERAGE(A1:A2)":  "1.5",
		"=AVERAGE(D2:F9)":  "38014.125",
		// AVERAGEA
		"=AVERAGEA(INT(1))": "1",
		"=AVERAGEA(A1)":     "1",
		"=AVERAGEA(A1:A2)":  "1.5",
		"=AVERAGEA(D2:F9)":  "12671.375",
		// CHIDIST
		"=CHIDIST(0.5,3)": "0.918891411654676",
		"=CHIDIST(8,3)":   "0.0460117056892315",
		// COUNT
		"=COUNT()":                        "0",
		"=COUNT(E1:F2,\"text\",1,INT(2))": "3",
		// COUNTA
		"=COUNTA()":                              "0",
		"=COUNTA(A1:A5,B2:B5,\"text\",1,INT(2))": "8",
		"=COUNTA(COUNTA(1),MUNIT(1))":            "2",
		// COUNTBLANK
		"=COUNTBLANK(MUNIT(1))": "0",
		"=COUNTBLANK(1)":        "0",
		"=COUNTBLANK(B1:C1)":    "1",
		"=COUNTBLANK(C1)":       "1",
		// FISHER
		"=FISHER(-0.9)":   "-1.47221948958322",
		"=FISHER(-0.25)":  "-0.255412811882995",
		"=FISHER(0.8)":    "1.09861228866811",
		"=FISHER(INT(0))": "0",
		// FISHERINV
		"=FISHERINV(-0.2)":   "-0.197375320224904",
		"=FISHERINV(INT(0))": "0",
		"=FISHERINV(2.8)":    "0.992631520201128",
		// GAMMA
		"=GAMMA(0.1)":    "9.513507698668732",
		"=GAMMA(INT(1))": "1",
		"=GAMMA(1.5)":    "0.886226925452758",
		"=GAMMA(5.5)":    "52.34277778455352",
		// GAMMALN
		"=GAMMALN(4.5)":    "2.45373657084244",
		"=GAMMALN(INT(1))": "0",
		// HARMEAN
		"=HARMEAN(2.5,3,0.5,1,3)":               "1.22950819672131",
		"=HARMEAN(\"2.5\",3,0.5,1,INT(3),\"\")": "1.22950819672131",
		// KURT
		"=KURT(F1:F9)":           "-1.03350350255137",
		"=KURT(F1,F2:F9)":        "-1.03350350255137",
		"=KURT(INT(1),MUNIT(2))": "-3.33333333333334",
		// NORM.DIST
		"=NORM.DIST(0.8,1,0.3,TRUE)": "0.252492537546923",
		"=NORM.DIST(50,40,20,FALSE)": "0.017603266338215",
		// NORMDIST
		"=NORMDIST(0.8,1,0.3,TRUE)": "0.252492537546923",
		"=NORMDIST(50,40,20,FALSE)": "0.017603266338215",
		// NORM.INV
		"=NORM.INV(0.6,5,2)": "5.506694205719997",
		// NORMINV
		"=NORMINV(0.6,5,2)":     "5.506694205719997",
		"=NORMINV(0.99,40,1.5)": "43.489521811582044",
		"=NORMINV(0.02,40,1.5)": "36.91937663649545",
		// NORM.S.DIST
		"=NORM.S.DIST(0.8,TRUE)": "0.788144601416603",
		// NORMSDIST
		"=NORMSDIST(1.333333)": "0.908788725604095",
		"=NORMSDIST(0)":        "0.5",
		// NORM.S.INV
		"=NORM.S.INV(0.25)": "-0.674489750223423",
		// NORMSINV
		"=NORMSINV(0.25)": "-0.674489750223423",
		// LARGE
		"=LARGE(A1:A5,1)": "3",
		"=LARGE(A1:B5,2)": "4",
		"=LARGE(A1,1)":    "1",
		"=LARGE(A1:F2,1)": "36693",
		// MAX
		"=MAX(1)":          "1",
		"=MAX(TRUE())":     "1",
		"=MAX(0.5,TRUE())": "1",
		"=MAX(FALSE())":    "0",
		"=MAX(MUNIT(2))":   "1",
		"=MAX(INT(1))":     "1",
		// MAXA
		"=MAXA(1)":          "1",
		"=MAXA(TRUE())":     "1",
		"=MAXA(0.5,TRUE())": "1",
		"=MAXA(FALSE())":    "0",
		"=MAXA(MUNIT(2))":   "1",
		"=MAXA(INT(1))":     "1",
		"=MAXA(A1:B4,MUNIT(1),INT(0),1,E1:F2,\"\")": "36693",
		// MEDIAN
		"=MEDIAN(A1:A5,12)":               "2",
		"=MEDIAN(A1:A5)":                  "1.5",
		"=MEDIAN(A1:A5,MEDIAN(A1:A5,12))": "2",
		// MIN
		"=MIN(1)":           "1",
		"=MIN(TRUE())":      "1",
		"=MIN(0.5,FALSE())": "0",
		"=MIN(FALSE())":     "0",
		"=MIN(MUNIT(2))":    "0",
		"=MIN(INT(1))":      "1",
		// MINA
		"=MINA(1)":           "1",
		"=MINA(TRUE())":      "1",
		"=MINA(0.5,FALSE())": "0",
		"=MINA(FALSE())":     "0",
		"=MINA(MUNIT(2))":    "0",
		"=MINA(INT(1))":      "1",
		"=MINA(A1:B4,MUNIT(1),INT(0),1,E1:F2,\"\")": "0",
		// PERCENTILE.INC
		"=PERCENTILE.INC(A1:A4,0.2)": "0.6",
		// PERCENTILE
		"=PERCENTILE(A1:A4,0.2)": "0.6",
		"=PERCENTILE(0,0)":       "0",
		// PERMUT
		"=PERMUT(6,6)":  "720",
		"=PERMUT(7,6)":  "5040",
		"=PERMUT(10,6)": "151200",
		// PERMUTATIONA
		"=PERMUTATIONA(6,6)": "46656",
		"=PERMUTATIONA(7,6)": "117649",
		// QUARTILE
		"=QUARTILE(A1:A4,2)": "1.5",
		// QUARTILE.INC
		"=QUARTILE.INC(A1:A4,0)": "0",
		// SKEW
		"=SKEW(1,2,3,4,3)": "-0.404796008910937",
		"=SKEW(A1:B2)":     "0",
		"=SKEW(A1:D3)":     "0",
		// SMALL
		"=SMALL(A1:A5,1)": "0",
		"=SMALL(A1:B5,2)": "1",
		"=SMALL(A1,1)":    "1",
		"=SMALL(A1:F2,1)": "1",
		// VARP
		"=VARP(A1:A5)": "1.25",
		// VAR.P
		"=VAR.P(A1:A5)": "1.25",
		// WEIBULL
		"=WEIBULL(1,3,1,FALSE)":  "1.103638323514327",
		"=WEIBULL(2,5,1.5,TRUE)": "0.985212776817482",
		// WEIBULL.DIST
		"=WEIBULL.DIST(1,3,1,FALSE)":  "1.103638323514327",
		"=WEIBULL.DIST(2,5,1.5,TRUE)": "0.985212776817482",
		// Information Functions
		// ISBLANK
		"=ISBLANK(A1)": "FALSE",
		"=ISBLANK(A5)": "TRUE",
		// ISERR
		"=ISERR(A1)":           "FALSE",
		"=ISERR(NA())":         "FALSE",
		"=ISERR(POWER(0,-1)))": "TRUE",
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
		"=ISNONTEXT(NA())":       "TRUE",
		// ISNUMBER
		"=ISNUMBER(A1)": "TRUE",
		"=ISNUMBER(D1)": "FALSE",
		// ISODD
		"=ISODD(A1)": "TRUE",
		"=ISODD(A2)": "FALSE",
		// ISTEXT
		"=ISTEXT(D1)": "TRUE",
		"=ISTEXT(A1)": "FALSE",
		// N
		"=N(10)":     "10",
		"=N(\"10\")": "10",
		"=N(\"x\")":  "0",
		"=N(TRUE)":   "1",
		"=N(FALSE)":  "0",
		// SHEET
		"=SHEET()": "1",
		// T
		"=T(\"text\")": "text",
		"=T(N(10))":    "",
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
		// FALSE
		"=FALSE()": "FALSE",
		// IFERROR
		"=IFERROR(1/2,0)":       "0.5",
		"=IFERROR(ISERROR(),0)": "0",
		"=IFERROR(1/0,0)":       "0",
		// IFNA
		"=IFNA(1,\"not found\")":    "1",
		"=IFNA(NA(),\"not found\")": "not found",
		// IFS
		"=IFS(4>1,5/4,4<-1,-5/4,TRUE,0)":     "1.25",
		"=IFS(-2>1,5/-2,-2<-1,-5/-2,TRUE,0)": "2.5",
		"=IFS(0>1,5/0,0<-1,-5/0,TRUE,0)":     "0",
		// NOT
		"=NOT(FALSE())":     "TRUE",
		"=NOT(\"false\")":   "TRUE",
		"=NOT(\"true\")":    "FALSE",
		"=NOT(ISBLANK(B1))": "TRUE",
		// OR
		"=OR(1)":       "TRUE",
		"=OR(0)":       "FALSE",
		"=OR(1=2,2=2)": "TRUE",
		"=OR(1=2,2=3)": "FALSE",
		// TRUE
		"=TRUE()": "TRUE",
		// XOR
		"=XOR(1>0,2>0)":                       "FALSE",
		"=XOR(1>0,0>1)":                       "TRUE",
		"=XOR(1>0,0>1,INT(0),INT(1),A1:A4,2)": "FALSE",
		// Date and Time Functions
		// DATE
		"=DATE(2020,10,21)": "2020-10-21 00:00:00 +0000 UTC",
		"=DATE(1900,1,1)":   "1899-12-31 00:00:00 +0000 UTC",
		// DATEDIF
		"=DATEDIF(43101,43101,\"D\")":  "0",
		"=DATEDIF(43101,43891,\"d\")":  "790",
		"=DATEDIF(43101,43891,\"Y\")":  "2",
		"=DATEDIF(42156,44242,\"y\")":  "5",
		"=DATEDIF(43101,43891,\"M\")":  "26",
		"=DATEDIF(42171,44242,\"m\")":  "67",
		"=DATEDIF(42156,44454,\"MD\")": "14",
		"=DATEDIF(42171,44242,\"md\")": "30",
		"=DATEDIF(43101,43891,\"YM\")": "2",
		"=DATEDIF(42171,44242,\"ym\")": "7",
		"=DATEDIF(43101,43891,\"YD\")": "59",
		"=DATEDIF(36526,73110,\"YD\")": "60",
		"=DATEDIF(42171,44242,\"yd\")": "244",
		// DATEVALUE
		"=DATEVALUE(\"01/01/16\")":   "42370",
		"=DATEVALUE(\"01/01/2016\")": "42370",
		"=DATEVALUE(\"01/01/29\")":   "47119",
		"=DATEVALUE(\"01/01/30\")":   "10959",
		// DAY
		"=DAY(0)":                                "0",
		"=DAY(INT(7))":                           "7",
		"=DAY(\"35\")":                           "4",
		"=DAY(42171)":                            "16",
		"=DAY(\"2-28-1900\")":                    "28",
		"=DAY(\"31-May-2015\")":                  "31",
		"=DAY(\"01/03/2019 12:14:16\")":          "3",
		"=DAY(\"January 25, 2020 01 AM\")":       "25",
		"=DAY(\"January 25, 2020 01:03 AM\")":    "25",
		"=DAY(\"January 25, 2020 12:00:00 AM\")": "25",
		"=DAY(\"1900-1-1\")":                     "1",
		"=DAY(\"12-1-1900\")":                    "1",
		"=DAY(\"3-January-1900\")":               "3",
		"=DAY(\"3-February-2000\")":              "3",
		"=DAY(\"3-February-2008\")":              "3",
		"=DAY(\"01/25/20\")":                     "25",
		"=DAY(\"01/25/31\")":                     "25",
		// DAYS
		"=DAYS(2,1)":                           "1",
		"=DAYS(INT(2),INT(1))":                 "1",
		"=DAYS(\"02/02/2015\",\"01/01/2015\")": "32",
		// MONTH
		"=MONTH(42171)":           "6",
		"=MONTH(\"31-May-2015\")": "5",
		// YEAR
		"=YEAR(15)":              "1900",
		"=YEAR(\"15\")":          "1900",
		"=YEAR(2048)":            "1905",
		"=YEAR(42171)":           "2015",
		"=YEAR(\"29-May-2015\")": "2015",
		"=YEAR(\"05/03/1984\")":  "1984",
		// YEARFRAC
		"=YEARFRAC(42005,42005)":                      "0",
		"=YEARFRAC(42005,42094)":                      "0.25",
		"=YEARFRAC(42005,42094,0)":                    "0.25",
		"=YEARFRAC(42005,42094,1)":                    "0.243835616438356",
		"=YEARFRAC(42005,42094,2)":                    "0.247222222222222",
		"=YEARFRAC(42005,42094,3)":                    "0.243835616438356",
		"=YEARFRAC(42005,42094,4)":                    "0.247222222222222",
		"=YEARFRAC(\"01/01/2015\",\"03/31/2015\")":    "0.25",
		"=YEARFRAC(\"01/01/2015\",\"03/31/2015\",0)":  "0.25",
		"=YEARFRAC(\"01/01/2015\",\"03/31/2015\",1)":  "0.243835616438356",
		"=YEARFRAC(\"01/01/2015\",\"03/31/2015\",2)":  "0.247222222222222",
		"=YEARFRAC(\"01/01/2015\",\"03/31/2015\",3)":  "0.243835616438356",
		"=YEARFRAC(\"01/01/2015\",\"03/31/2015\",4)":  "0.247222222222222",
		"=YEARFRAC(\"01/01/2015\",42094)":             "0.25",
		"=YEARFRAC(42005,\"03/31/2015\",0)":           "0.25",
		"=YEARFRAC(\"01/31/2015\",\"03/31/2015\")":    "0.166666666666667",
		"=YEARFRAC(\"01/30/2015\",\"03/31/2015\")":    "0.166666666666667",
		"=YEARFRAC(\"02/29/2000\", \"02/29/2008\")":   "8",
		"=YEARFRAC(\"02/29/2000\", \"02/29/2008\",1)": "7.998175182481752",
		"=YEARFRAC(\"02/29/2000\", \"01/29/2001\",1)": "0.915300546448087",
		"=YEARFRAC(\"02/29/2000\", \"03/29/2000\",1)": "0.0792349726775956",
		"=YEARFRAC(\"01/31/2000\", \"03/29/2000\",4)": "0.163888888888889",
		// TIME
		"=TIME(5,44,32)":             "0.239259259259259",
		"=TIME(\"5\",\"44\",\"32\")": "0.239259259259259",
		"=TIME(0,0,73)":              "0.000844907407407407",
		// Text Functions
		// CHAR
		"=CHAR(65)": "A",
		"=CHAR(97)": "a",
		"=CHAR(63)": "?",
		"=CHAR(51)": "3",
		// CLEAN
		"=CLEAN(\"\u0009clean text\")": "clean text",
		"=CLEAN(0)":                    "0",
		// CODE
		"=CODE(\"Alpha\")": "65",
		"=CODE(\"alpha\")": "97",
		"=CODE(\"?\")":     "63",
		"=CODE(\"3\")":     "51",
		"=CODE(\"\")":      "0",
		// CONCAT
		"=CONCAT(TRUE(),1,FALSE(),\"0\",INT(2))": "TRUE1FALSE02",
		// CONCATENATE
		"=CONCATENATE(TRUE(),1,FALSE(),\"0\",INT(2))": "TRUE1FALSE02",
		// EXACT
		"=EXACT(1,\"1\")":     "TRUE",
		"=EXACT(1,1)":         "TRUE",
		"=EXACT(\"A\",\"a\")": "FALSE",
		// FIXED
		"=FIXED(5123.591)":         "5,123.591",
		"=FIXED(5123.591,1)":       "5,123.6",
		"=FIXED(5123.591,0)":       "5,124",
		"=FIXED(5123.591,-1)":      "5,120",
		"=FIXED(5123.591,-2)":      "5,100",
		"=FIXED(5123.591,-3,TRUE)": "5000",
		"=FIXED(5123.591,-5)":      "0",
		"=FIXED(-77262.23973,-5)":  "-100,000",
		// FIND
		"=FIND(\"T\",\"Original Text\")":   "10",
		"=FIND(\"t\",\"Original Text\")":   "13",
		"=FIND(\"i\",\"Original Text\")":   "3",
		"=FIND(\"i\",\"Original Text\",4)": "5",
		"=FIND(\"\",\"Original Text\")":    "1",
		"=FIND(\"\",\"Original Text\",2)":  "2",
		// FINDB
		"=FINDB(\"T\",\"Original Text\")":   "10",
		"=FINDB(\"t\",\"Original Text\")":   "13",
		"=FINDB(\"i\",\"Original Text\")":   "3",
		"=FINDB(\"i\",\"Original Text\",4)": "5",
		"=FINDB(\"\",\"Original Text\")":    "1",
		"=FINDB(\"\",\"Original Text\",2)":  "2",
		// LEFT
		"=LEFT(\"Original Text\")":    "O",
		"=LEFT(\"Original Text\",4)":  "Orig",
		"=LEFT(\"Original Text\",0)":  "",
		"=LEFT(\"Original Text\",13)": "Original Text",
		"=LEFT(\"Original Text\",20)": "Original Text",
		// LEFTB
		"=LEFTB(\"Original Text\")":    "O",
		"=LEFTB(\"Original Text\",4)":  "Orig",
		"=LEFTB(\"Original Text\",0)":  "",
		"=LEFTB(\"Original Text\",13)": "Original Text",
		"=LEFTB(\"Original Text\",20)": "Original Text",
		// LEN
		"=LEN(\"\")": "0",
		"=LEN(D1)":   "5",
		// LENB
		"=LENB(\"\")": "0",
		"=LENB(D1)":   "5",
		// LOWER
		"=LOWER(\"test\")":     "test",
		"=LOWER(\"TEST\")":     "test",
		"=LOWER(\"Test\")":     "test",
		"=LOWER(\"TEST 123\")": "test 123",
		// MID
		"=MID(\"Original Text\",7,1)": "a",
		"=MID(\"Original Text\",4,7)": "ginal T",
		"=MID(\"255 years\",3,1)":     "5",
		"=MID(\"text\",3,6)":          "xt",
		"=MID(\"text\",6,0)":          "",
		// MIDB
		"=MIDB(\"Original Text\",7,1)": "a",
		"=MIDB(\"Original Text\",4,7)": "ginal T",
		"=MIDB(\"255 years\",3,1)":     "5",
		"=MIDB(\"text\",3,6)":          "xt",
		"=MIDB(\"text\",6,0)":          "",
		// PROPER
		"=PROPER(\"this is a test sentence\")": "This Is A Test Sentence",
		"=PROPER(\"THIS IS A TEST SENTENCE\")": "This Is A Test Sentence",
		"=PROPER(\"123tEST teXT\")":            "123Test Text",
		"=PROPER(\"Mr. SMITH's address\")":     "Mr. Smith'S Address",
		// REPLACE
		"=REPLACE(\"test string\",7,3,\"X\")":          "test sXng",
		"=REPLACE(\"second test string\",8,4,\"XXX\")": "second XXX string",
		"=REPLACE(\"text\",5,0,\" and char\")":         "text and char",
		"=REPLACE(\"text\",1,20,\"char and \")":        "char and ",
		// REPLACEB
		"=REPLACEB(\"test string\",7,3,\"X\")":          "test sXng",
		"=REPLACEB(\"second test string\",8,4,\"XXX\")": "second XXX string",
		"=REPLACEB(\"text\",5,0,\" and char\")":         "text and char",
		"=REPLACEB(\"text\",1,20,\"char and \")":        "char and ",
		// REPT
		"=REPT(\"*\",0)":  "",
		"=REPT(\"*\",1)":  "*",
		"=REPT(\"**\",2)": "****",
		// RIGHT
		"=RIGHT(\"Original Text\")":    "t",
		"=RIGHT(\"Original Text\",4)":  "Text",
		"=RIGHT(\"Original Text\",0)":  "",
		"=RIGHT(\"Original Text\",13)": "Original Text",
		"=RIGHT(\"Original Text\",20)": "Original Text",
		// RIGHTB
		"=RIGHTB(\"Original Text\")":    "t",
		"=RIGHTB(\"Original Text\",4)":  "Text",
		"=RIGHTB(\"Original Text\",0)":  "",
		"=RIGHTB(\"Original Text\",13)": "Original Text",
		"=RIGHTB(\"Original Text\",20)": "Original Text",
		// SUBSTITUTE
		"=SUBSTITUTE(\"abab\",\"a\",\"X\")":                      "XbXb",
		"=SUBSTITUTE(\"abab\",\"a\",\"X\",2)":                    "abXb",
		"=SUBSTITUTE(\"abab\",\"x\",\"X\",2)":                    "abab",
		"=SUBSTITUTE(\"John is 5 years old\",\"John\",\"Jack\")": "Jack is 5 years old",
		"=SUBSTITUTE(\"John is 5 years old\",\"5\",\"6\")":       "John is 6 years old",
		// TRIM
		"=TRIM(\" trim text \")": "trim text",
		"=TRIM(0)":               "0",
		// UNICHAR
		"=UNICHAR(65)": "A",
		"=UNICHAR(97)": "a",
		"=UNICHAR(63)": "?",
		"=UNICHAR(51)": "3",
		// UNICODE
		"=UNICODE(\"Alpha\")": "65",
		"=UNICODE(\"alpha\")": "97",
		"=UNICODE(\"?\")":     "63",
		"=UNICODE(\"3\")":     "51",
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
		`=IF(1=1, "equal", "notequal")`:         "equal",
		`=IF(1<>1, "equal", "notequal")`:        "notequal",
		`=IF("A"="A", "equal", "notequal")`:     "equal",
		`=IF("A"<>"A", "equal", "notequal")`:    "notequal",
		`=IF(FALSE,0,ROUND(4/2,0))`:             "2",
		`=IF(TRUE,ROUND(4/2,0),0)`:              "2",
		// Excel Lookup and Reference Functions
		// ADDRESS
		"=ADDRESS(1,1,1,TRUE)":            "$A$1",
		"=ADDRESS(1,1,1,FALSE)":           "R1C1",
		"=ADDRESS(1,1,2,TRUE)":            "A$1",
		"=ADDRESS(1,1,2,FALSE)":           "R1C[1]",
		"=ADDRESS(1,1,3,TRUE)":            "$A1",
		"=ADDRESS(1,1,3,FALSE)":           "R[1]C1",
		"=ADDRESS(1,1,4,TRUE)":            "A1",
		"=ADDRESS(1,1,4,FALSE)":           "R[1]C[1]",
		"=ADDRESS(1,1,4,TRUE,\"\")":       "A1",
		"=ADDRESS(1,1,4,TRUE,\"Sheet1\")": "Sheet1!A1",
		// CHOOSE
		"=CHOOSE(4,\"red\",\"blue\",\"green\",\"brown\")": "brown",
		"=CHOOSE(1,\"red\",\"blue\",\"green\",\"brown\")": "red",
		"=SUM(CHOOSE(A2,A1,B1:B2,A1:A3,A1:A4))":           "9",
		// COLUMN
		"=COLUMN()":                "3",
		"=COLUMN(Sheet1!A1)":       "1",
		"=COLUMN(Sheet1!A1:B1:C1)": "1",
		"=COLUMN(Sheet1!F1:G1)":    "6",
		"=COLUMN(H1)":              "8",
		// COLUMNS
		"=COLUMNS(B1)":                   "1",
		"=COLUMNS(1:1)":                  "16384",
		"=COLUMNS(Sheet1!1:1)":           "16384",
		"=COLUMNS(B1:E5)":                "4",
		"=COLUMNS(Sheet1!E5:H7:B1)":      "7",
		"=COLUMNS(E5:H7:B1:C1:Z1:C1:B1)": "25",
		"=COLUMNS(E5:B1)":                "4",
		"=COLUMNS(EM38:HZ81)":            "92",
		// HLOOKUP
		"=HLOOKUP(D2,D2:D8,1,FALSE)":          "Jan",
		"=HLOOKUP(F3,F3:F8,3,FALSE)":          "34440",
		"=HLOOKUP(INT(F3),F3:F8,3,FALSE)":     "34440",
		"=HLOOKUP(MUNIT(1),MUNIT(1),1,FALSE)": "1",
		// VLOOKUP
		"=VLOOKUP(D2,D:D,1,FALSE)":            "Jan",
		"=VLOOKUP(D2,D1:D10,1)":               "Jan",
		"=VLOOKUP(D2,D1:D11,1)":               "Feb",
		"=VLOOKUP(D2,D1:D10,1,FALSE)":         "Jan",
		"=VLOOKUP(INT(36693),F2:F2,1,FALSE)":  "36693",
		"=VLOOKUP(INT(F2),F3:F9,1)":           "32080",
		"=VLOOKUP(INT(F2),F3:F9,1,TRUE)":      "32080",
		"=VLOOKUP(MUNIT(3),MUNIT(3),1)":       "0",
		"=VLOOKUP(A1,A3:B5,1)":                "0",
		"=VLOOKUP(MUNIT(1),MUNIT(1),1,FALSE)": "1",
		// LOOKUP
		"=LOOKUP(F8,F8:F9,F8:F9)":      "32080",
		"=LOOKUP(F8,F8:F9,D8:D9)":      "Feb",
		"=LOOKUP(E3,E2:E5,F2:F5)":      "22100",
		"=LOOKUP(E3,E2:F5)":            "22100",
		"=LOOKUP(F3+1,F3:F4,F3:F4)":    "22100",
		"=LOOKUP(F4+1,F3:F4,F3:F4)":    "53321",
		"=LOOKUP(1,MUNIT(1))":          "1",
		"=LOOKUP(1,MUNIT(1),MUNIT(1))": "1",
		// ROW
		"=ROW()":                "1",
		"=ROW(Sheet1!A1)":       "1",
		"=ROW(Sheet1!A1:B2:C3)": "1",
		"=ROW(Sheet1!F5:G6)":    "5",
		"=ROW(A8)":              "8",
		// ROWS
		"=ROWS(B1)":                    "1",
		"=ROWS(B:B)":                   "1048576",
		"=ROWS(Sheet1!B:B)":            "1048576",
		"=ROWS(B1:E5)":                 "5",
		"=ROWS(Sheet1!E5:H7:B1)":       "7",
		"=ROWS(E5:H8:B2:C3:Z26:C3:B2)": "25",
		"=ROWS(E5:B1)":                 "5",
		"=ROWS(EM38:HZ81)":             "44",
		// Web Functions
		// ENCODEURL
		"=ENCODEURL(\"https://xuri.me/excelize/en/?q=Save As\")": "https%3A%2F%2Fxuri.me%2Fexcelize%2Fen%2F%3Fq%3DSave%20As",
		// Financial Functions
		// ACCRINTM
		"=ACCRINTM(\"01/01/2012\",\"12/31/2012\",8%,10000)":   "800",
		"=ACCRINTM(\"01/01/2012\",\"12/31/2012\",8%,10000,3)": "800",
		// AMORDEGRC
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",20,1,20%)":    "42",
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",20,1,20%,4)":  "42",
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",20,1,40%,4)":  "42",
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",20,1,25%,4)":  "41",
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",109,1,25%,4)": "54",
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",110,2,25%,4)": "0",
		// AMORLINC
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",20,1,20%,4)":  "30",
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",20,1,0%,4)":   "0",
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",20,20,15%,4)": "0",
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",20,6,15%,4)":  "0.6875",
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",20,0,15%,4)":  "16.8125",
		// CUMIPMT
		"=CUMIPMT(0.05/12,60,50000,1,12,0)":  "-2294.97753732664",
		"=CUMIPMT(0.05/12,60,50000,13,24,0)": "-1833.1000665738893",
		// CUMPRINC
		"=CUMPRINC(0.05/12,60,50000,1,12,0)":  "-9027.762649079885",
		"=CUMPRINC(0.05/12,60,50000,13,24,0)": "-9489.640119832635",
		// DB
		"=DB(0,1000,5,1)":       "0",
		"=DB(10000,1000,5,1)":   "3690",
		"=DB(10000,1000,5,2)":   "2328.39",
		"=DB(10000,1000,5,1,6)": "1845",
		"=DB(10000,1000,5,6,6)": "238.52712458788187",
		// DDB
		"=DDB(0,1000,5,1)":     "0",
		"=DDB(10000,1000,5,1)": "4000",
		"=DDB(10000,1000,5,2)": "2400",
		"=DDB(10000,1000,5,3)": "1440",
		"=DDB(10000,1000,5,4)": "864",
		"=DDB(10000,1000,5,5)": "296",
		// DISC
		"=DISC(\"04/01/2016\",\"03/31/2021\",95,100)": "0.01",
		// DOLLARDE
		"=DOLLARDE(1.01,16)": "1.0625",
		// DOLLARFR
		"=DOLLARFR(1.0625,16)": "1.01",
		// EFFECT
		"=EFFECT(0.1,4)":   "0.103812890625",
		"=EFFECT(0.025,2)": "0.02515625",
		// FV
		"=FV(0.05/12,60,-1000)":   "68006.08284084337",
		"=FV(0.1/4,16,-2000,0,1)": "39729.46089416617",
		"=FV(0,16,-2000)":         "32000",
		// FVSCHEDULE
		"=FVSCHEDULE(10000,A1:A5)": "240000",
		"=FVSCHEDULE(10000,0.5)":   "15000",
		// INTRATE
		"=INTRATE(\"04/01/2005\",\"03/31/2010\",1000,2125)": "0.225",
		// IPMT
		"=IPMT(0.05/12,2,60,50000)":   "-205.26988187971995",
		"=IPMT(0.035/4,2,8,0,5000,1)": "5.257455237829077",
		// ISPMT
		"=ISPMT(0.05/12,1,60,50000)": "-204.8611111111111",
		"=ISPMT(0.05/12,2,60,50000)": "-201.38888888888886",
		"=ISPMT(0.05/12,2,1,50000)":  "208.33333333333334",
		// NOMINAL
		"=NOMINAL(0.025,12)": "0.0247180352381129",
		// NPER
		"=NPER(0.04,-6000,50000)":           "10.338035071507665",
		"=NPER(0,-6000,50000)":              "8.333333333333334",
		"=NPER(0.06/4,-2000,60000,30000,1)": "52.794773709274764",
		// NPV
		"=NPV(0.02,-5000,\"\",800)": "-4133.025759323337",
		// PDURATION
		"=PDURATION(0.04,10000,15000)": "10.33803507150765",
		// PMT
		"=PMT(0,8,0,5000,1)":       "-625",
		"=PMT(0.035/4,8,0,5000,1)": "-600.8520271804658",
		// PPMT
		"=PPMT(0.05/12,2,60,50000)":   "-738.2918003208238",
		"=PPMT(0.035/4,2,8,0,5000,1)": "-606.1094824182949",
		// PRICEDISC
		"=PRICEDISC(\"04/01/2017\",\"03/31/2021\",2.5%,100)":   "90",
		"=PRICEDISC(\"04/01/2017\",\"03/31/2021\",2.5%,100,3)": "90",
		// RRI
		"=RRI(10,10000,15000)": "0.0413797439924106",
		// SLN
		"=SLN(10000,1000,5)": "1800",
		// SYD
		"=SYD(10000,1000,5,1)": "3000",
		"=SYD(10000,1000,5,2)": "2400",
	}
	for formula, expected := range mathCalc {
		f := prepareCalcData(cellData)
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	mathCalcError := map[string]string{
		"=1/0":       "#DIV/0!",
		"1^\"text\"": "strconv.ParseFloat: parsing \"text\": invalid syntax",
		"\"text\"^1": "strconv.ParseFloat: parsing \"text\": invalid syntax",
		"1+\"text\"": "strconv.ParseFloat: parsing \"text\": invalid syntax",
		"\"text\"+1": "strconv.ParseFloat: parsing \"text\": invalid syntax",
		"1-\"text\"": "strconv.ParseFloat: parsing \"text\": invalid syntax",
		"\"text\"-1": "strconv.ParseFloat: parsing \"text\": invalid syntax",
		"1*\"text\"": "strconv.ParseFloat: parsing \"text\": invalid syntax",
		"\"text\"*1": "strconv.ParseFloat: parsing \"text\": invalid syntax",
		"1/\"text\"": "strconv.ParseFloat: parsing \"text\": invalid syntax",
		"\"text\"/1": "strconv.ParseFloat: parsing \"text\": invalid syntax",
		// Engineering Functions
		// BESSELI
		"=BESSELI()":       "BESSELI requires 2 numeric arguments",
		"=BESSELI(\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=BESSELI(0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// BESSELJ
		"=BESSELJ()":       "BESSELJ requires 2 numeric arguments",
		"=BESSELJ(\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=BESSELJ(0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// BESSELK
		"=BESSELK()":       "BESSELK requires 2 numeric arguments",
		"=BESSELK(\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=BESSELK(0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=BESSELK(-1,0)":   "#NUM!",
		"=BESSELK(1,-1)":   "#NUM!",
		// BESSELY
		"=BESSELY()":       "BESSELY requires 2 numeric arguments",
		"=BESSELY(\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=BESSELY(0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=BESSELY(-1,0)":   "#NUM!",
		"=BESSELY(1,-1)":   "#NUM!",
		// BIN2DEC
		"=BIN2DEC()":     "BIN2DEC requires 1 numeric argument",
		"=BIN2DEC(\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// BIN2HEX
		"=BIN2HEX()":               "BIN2HEX requires at least 1 argument",
		"=BIN2HEX(1,1,1)":          "BIN2HEX allows at most 2 arguments",
		"=BIN2HEX(\"\",1)":         "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=BIN2HEX(1,\"\")":         "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=BIN2HEX(12345678901,10)": "#NUM!",
		"=BIN2HEX(1,-1)":           "#NUM!",
		"=BIN2HEX(31,1)":           "#NUM!",
		// BIN2OCT
		"=BIN2OCT()":                 "BIN2OCT requires at least 1 argument",
		"=BIN2OCT(1,1,1)":            "BIN2OCT allows at most 2 arguments",
		"=BIN2OCT(\"\",1)":           "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=BIN2OCT(1,\"\")":           "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=BIN2OCT(-12345678901 ,10)": "#NUM!",
		"=BIN2OCT(1,-1)":             "#NUM!",
		"=BIN2OCT(8,1)":              "#NUM!",
		// BITAND
		"=BITAND()":        "BITAND requires 2 numeric arguments",
		"=BITAND(-1,2)":    "#NUM!",
		"=BITAND(2^48,2)":  "#NUM!",
		"=BITAND(1,-1)":    "#NUM!",
		"=BITAND(\"\",-1)": "#NUM!",
		"=BITAND(1,\"\")":  "#NUM!",
		"=BITAND(1,2^48)":  "#NUM!",
		// BITLSHIFT
		"=BITLSHIFT()":        "BITLSHIFT requires 2 numeric arguments",
		"=BITLSHIFT(-1,2)":    "#NUM!",
		"=BITLSHIFT(2^48,2)":  "#NUM!",
		"=BITLSHIFT(1,-1)":    "#NUM!",
		"=BITLSHIFT(\"\",-1)": "#NUM!",
		"=BITLSHIFT(1,\"\")":  "#NUM!",
		"=BITLSHIFT(1,2^48)":  "#NUM!",
		// BITOR
		"=BITOR()":        "BITOR requires 2 numeric arguments",
		"=BITOR(-1,2)":    "#NUM!",
		"=BITOR(2^48,2)":  "#NUM!",
		"=BITOR(1,-1)":    "#NUM!",
		"=BITOR(\"\",-1)": "#NUM!",
		"=BITOR(1,\"\")":  "#NUM!",
		"=BITOR(1,2^48)":  "#NUM!",
		// BITRSHIFT
		"=BITRSHIFT()":        "BITRSHIFT requires 2 numeric arguments",
		"=BITRSHIFT(-1,2)":    "#NUM!",
		"=BITRSHIFT(2^48,2)":  "#NUM!",
		"=BITRSHIFT(1,-1)":    "#NUM!",
		"=BITRSHIFT(\"\",-1)": "#NUM!",
		"=BITRSHIFT(1,\"\")":  "#NUM!",
		"=BITRSHIFT(1,2^48)":  "#NUM!",
		// BITXOR
		"=BITXOR()":        "BITXOR requires 2 numeric arguments",
		"=BITXOR(-1,2)":    "#NUM!",
		"=BITXOR(2^48,2)":  "#NUM!",
		"=BITXOR(1,-1)":    "#NUM!",
		"=BITXOR(\"\",-1)": "#NUM!",
		"=BITXOR(1,\"\")":  "#NUM!",
		"=BITXOR(1,2^48)":  "#NUM!",
		// COMPLEX
		"=COMPLEX()":              "COMPLEX requires at least 2 arguments",
		"=COMPLEX(10,-5,\"\")":    "#VALUE!",
		"=COMPLEX(\"\",0)":        "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=COMPLEX(0,\"\")":        "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=COMPLEX(10,-5,\"i\",0)": "COMPLEX allows at most 3 arguments",
		// DEC2BIN
		"=DEC2BIN()":        "DEC2BIN requires at least 1 argument",
		"=DEC2BIN(1,1,1)":   "DEC2BIN allows at most 2 arguments",
		"=DEC2BIN(\"\",1)":  "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=DEC2BIN(1,\"\")":  "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=DEC2BIN(-513,10)": "#NUM!",
		"=DEC2BIN(1,-1)":    "#NUM!",
		"=DEC2BIN(2,1)":     "#NUM!",
		// DEC2HEX
		"=DEC2HEX()":                 "DEC2HEX requires at least 1 argument",
		"=DEC2HEX(1,1,1)":            "DEC2HEX allows at most 2 arguments",
		"=DEC2HEX(\"\",1)":           "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=DEC2HEX(1,\"\")":           "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=DEC2HEX(-549755813888,10)": "#NUM!",
		"=DEC2HEX(1,-1)":             "#NUM!",
		"=DEC2HEX(31,1)":             "#NUM!",
		// DEC2OCT
		"=DEC2OCT()":               "DEC2OCT requires at least 1 argument",
		"=DEC2OCT(1,1,1)":          "DEC2OCT allows at most 2 arguments",
		"=DEC2OCT(\"\",1)":         "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=DEC2OCT(1,\"\")":         "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=DEC2OCT(-536870912 ,10)": "#NUM!",
		"=DEC2OCT(1,-1)":           "#NUM!",
		"=DEC2OCT(8,1)":            "#NUM!",
		// HEX2BIN
		"=HEX2BIN()":        "HEX2BIN requires at least 1 argument",
		"=HEX2BIN(1,1,1)":   "HEX2BIN allows at most 2 arguments",
		"=HEX2BIN(\"X\",1)": "strconv.ParseInt: parsing \"X\": invalid syntax",
		"=HEX2BIN(1,\"\")":  "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=HEX2BIN(-513,10)": "strconv.ParseInt: parsing \"-\": invalid syntax",
		"=HEX2BIN(1,-1)":    "#NUM!",
		"=HEX2BIN(2,1)":     "#NUM!",
		// HEX2DEC
		"=HEX2DEC()":      "HEX2DEC requires 1 numeric argument",
		"=HEX2DEC(\"X\")": "strconv.ParseInt: parsing \"X\": invalid syntax",
		// HEX2OCT
		"=HEX2OCT()":        "HEX2OCT requires at least 1 argument",
		"=HEX2OCT(1,1,1)":   "HEX2OCT allows at most 2 arguments",
		"=HEX2OCT(\"X\",1)": "strconv.ParseInt: parsing \"X\": invalid syntax",
		"=HEX2OCT(1,\"\")":  "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=HEX2OCT(-513,10)": "strconv.ParseInt: parsing \"-\": invalid syntax",
		"=HEX2OCT(1,-1)":    "#NUM!",
		// IMABS
		"=IMABS()":     "IMABS requires 1 argument",
		"=IMABS(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		// IMAGINARY
		"=IMAGINARY()":     "IMAGINARY requires 1 argument",
		"=IMAGINARY(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		// IMARGUMENT
		"=IMARGUMENT()":     "IMARGUMENT requires 1 argument",
		"=IMARGUMENT(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		// IMCONJUGATE
		"=IMCONJUGATE()":     "IMCONJUGATE requires 1 argument",
		"=IMCONJUGATE(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		// IMCOS
		"=IMCOS()":     "IMCOS requires 1 argument",
		"=IMCOS(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		// IMCOSH
		"=IMCOSH()":     "IMCOSH requires 1 argument",
		"=IMCOSH(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		// IMCOT
		"=IMCOT()":     "IMCOT requires 1 argument",
		"=IMCOT(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		// IMCSC
		"=IMCSC()":     "IMCSC requires 1 argument",
		"=IMCSC(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		"=IMCSC(0)":    "#NUM!",
		// IMCSCH
		"=IMCSCH()":     "IMCSCH requires 1 argument",
		"=IMCSCH(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		"=IMCSCH(0)":    "#NUM!",
		// IMDIV
		"=IMDIV()":       "IMDIV requires 2 arguments",
		"=IMDIV(0,\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		"=IMDIV(\"\",0)": "strconv.ParseComplex: parsing \"\": invalid syntax",
		"=IMDIV(1,0)":    "#NUM!",
		// IMEXP
		"=IMEXP()":     "IMEXP requires 1 argument",
		"=IMEXP(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		// IMLN
		"=IMLN()":     "IMLN requires 1 argument",
		"=IMLN(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		"=IMLN(0)":    "#NUM!",
		// IMLOG10
		"=IMLOG10()":     "IMLOG10 requires 1 argument",
		"=IMLOG10(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		"=IMLOG10(0)":    "#NUM!",
		// IMLOG2
		"=IMLOG2()":     "IMLOG2 requires 1 argument",
		"=IMLOG2(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		"=IMLOG2(0)":    "#NUM!",
		// IMPOWER
		"=IMPOWER()":       "IMPOWER requires 2 arguments",
		"=IMPOWER(0,\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		"=IMPOWER(\"\",0)": "strconv.ParseComplex: parsing \"\": invalid syntax",
		"=IMPOWER(0,0)":    "#NUM!",
		"=IMPOWER(0,-1)":   "#NUM!",
		// IMPRODUCT
		"=IMPRODUCT(\"x\")": "strconv.ParseComplex: parsing \"x\": invalid syntax",
		"=IMPRODUCT(A1:D1)": "strconv.ParseComplex: parsing \"Month\": invalid syntax",
		// IMREAL
		"=IMREAL()":     "IMREAL requires 1 argument",
		"=IMREAL(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		// IMSEC
		"=IMSEC()":     "IMSEC requires 1 argument",
		"=IMSEC(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		// IMSECH
		"=IMSECH()":     "IMSECH requires 1 argument",
		"=IMSECH(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		// IMSIN
		"=IMSIN()":     "IMSIN requires 1 argument",
		"=IMSIN(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		// IMSINH
		"=IMSINH()":     "IMSINH requires 1 argument",
		"=IMSINH(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		// IMSQRT
		"=IMSQRT()":     "IMSQRT requires 1 argument",
		"=IMSQRT(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		// IMSUB
		"=IMSUB()":       "IMSUB requires 2 arguments",
		"=IMSUB(0,\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		"=IMSUB(\"\",0)": "strconv.ParseComplex: parsing \"\": invalid syntax",
		// IMSUM
		"=IMSUM()":     "IMSUM requires at least 1 argument",
		"=IMSUM(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		// IMTAN
		"=IMTAN()":     "IMTAN requires 1 argument",
		"=IMTAN(\"\")": "strconv.ParseComplex: parsing \"\": invalid syntax",
		// OCT2BIN
		"=OCT2BIN()":               "OCT2BIN requires at least 1 argument",
		"=OCT2BIN(1,1,1)":          "OCT2BIN allows at most 2 arguments",
		"=OCT2BIN(\"\",1)":         "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=OCT2BIN(1,\"\")":         "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=OCT2BIN(-536870912 ,10)": "#NUM!",
		"=OCT2BIN(1,-1)":           "#NUM!",
		// OCT2DEC
		"=OCT2DEC()":     "OCT2DEC requires 1 numeric argument",
		"=OCT2DEC(\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// OCT2HEX
		"=OCT2HEX()":               "OCT2HEX requires at least 1 argument",
		"=OCT2HEX(1,1,1)":          "OCT2HEX allows at most 2 arguments",
		"=OCT2HEX(\"\",1)":         "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=OCT2HEX(1,\"\")":         "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=OCT2HEX(-536870912 ,10)": "#NUM!",
		"=OCT2HEX(1,-1)":           "#NUM!",
		// Math and Trigonometric Functions
		// ABS
		"=ABS()":    "ABS requires 1 numeric argument",
		`=ABS("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		"=ABS(~)":   `invalid column name "~"`,
		// ACOS
		"=ACOS()":        "ACOS requires 1 numeric argument",
		`=ACOS("X")`:     "strconv.ParseFloat: parsing \"X\": invalid syntax",
		"=ACOS(ACOS(0))": "#NUM!",
		// ACOSH
		"=ACOSH()":    "ACOSH requires 1 numeric argument",
		`=ACOSH("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// _xlfn.ACOT
		"=_xlfn.ACOT()":    "ACOT requires 1 numeric argument",
		`=_xlfn.ACOT("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// _xlfn.ACOTH
		"=_xlfn.ACOTH()":               "ACOTH requires 1 numeric argument",
		`=_xlfn.ACOTH("X")`:            "strconv.ParseFloat: parsing \"X\": invalid syntax",
		"=_xlfn.ACOTH(_xlfn.ACOTH(2))": "#NUM!",
		// _xlfn.ARABIC
		"=_xlfn.ARABIC()": "ARABIC requires 1 numeric argument",
		"=_xlfn.ARABIC(\"" + strings.Repeat("I", 256) + "\")": "#VALUE!",
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
		`=BASE(1,"X")`:   "strconv.ParseFloat: parsing \"X\": invalid syntax",
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
		`=_xlfn.CEILING.PRECISE(1,"X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
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
		"=GCD(\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
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
		// MDETERM
		"MDETERM()": "MDETERM requires at least 1 argument",
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
		"=_xlfn.MUNIT()":    "MUNIT requires 1 numeric argument",
		`=_xlfn.MUNIT("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		"=_xlfn.MUNIT(-1)":  "",
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
		`=RANDBETWEEN("X",1)`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		`=RANDBETWEEN(1,"X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
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
		`=SQRT("")`:  "strconv.ParseFloat: parsing \"\": invalid syntax",
		`=SQRT("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		"=SQRT(-1)":  "#NUM!",
		// SQRTPI
		"=SQRTPI()":    "SQRTPI requires 1 numeric argument",
		`=SQRTPI("X")`: "strconv.ParseFloat: parsing \"X\": invalid syntax",
		// STDEV
		"=STDEV()":      "STDEV requires at least 1 argument",
		"=STDEV(E2:E9)": "#DIV/0!",
		// STDEV.S
		"=STDEV.S()": "STDEV.S requires at least 1 argument",
		// STDEVA
		"=STDEVA()":      "STDEVA requires at least 1 argument",
		"=STDEVA(E2:E9)": "#DIV/0!",
		// POISSON.DIST
		"=POISSON.DIST()": "POISSON.DIST requires 3 arguments",
		// POISSON
		"=POISSON()":             "POISSON requires 3 arguments",
		"=POISSON(\"\",0,FALSE)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=POISSON(0,\"\",FALSE)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=POISSON(0,0,\"\")":     "strconv.ParseBool: parsing \"\": invalid syntax",
		"=POISSON(0,-1,TRUE)":    "#N/A",
		// SUM
		"=SUM((":   ErrInvalidFormula.Error(),
		"=SUM(-)":  ErrInvalidFormula.Error(),
		"=SUM(1+)": ErrInvalidFormula.Error(),
		"=SUM(1-)": ErrInvalidFormula.Error(),
		"=SUM(1*)": ErrInvalidFormula.Error(),
		"=SUM(1/)": ErrInvalidFormula.Error(),
		// SUMIF
		"=SUMIF()": "SUMIF requires at least 2 argument",
		// SUMSQ
		`=SUMSQ("X")`:   "strconv.ParseFloat: parsing \"X\": invalid syntax",
		"=SUMSQ(C1:D2)": "strconv.ParseFloat: parsing \"Month\": invalid syntax",
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
		// AVEDEV
		"=AVEDEV()":       "AVEDEV requires at least 1 argument",
		"=AVEDEV(\"\")":   "#VALUE!",
		"=AVEDEV(1,\"\")": "#VALUE!",
		// AVERAGE
		"=AVERAGE(H1)": "AVERAGE divide by zero",
		// AVERAGE
		"=AVERAGEA(H1)": "AVERAGEA divide by zero",
		// CHIDIST
		"=CHIDIST()":         "CHIDIST requires 2 numeric arguments",
		"=CHIDIST(\"\",3)":   "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=CHIDIST(0.5,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// COUNTBLANK
		"=COUNTBLANK()":    "COUNTBLANK requires 1 argument",
		"=COUNTBLANK(1,2)": "COUNTBLANK requires 1 argument",
		// FISHER
		"=FISHER()":         "FISHER requires 1 numeric argument",
		"=FISHER(2)":        "#N/A",
		"=FISHER(INT(-2)))": "#N/A",
		"=FISHER(F1)":       "FISHER requires 1 numeric argument",
		// FISHERINV
		"=FISHERINV()":   "FISHERINV requires 1 numeric argument",
		"=FISHERINV(F1)": "FISHERINV requires 1 numeric argument",
		// GAMMA
		"=GAMMA()":       "GAMMA requires 1 numeric argument",
		"=GAMMA(F1)":     "GAMMA requires 1 numeric argument",
		"=GAMMA(0)":      "#N/A",
		"=GAMMA(INT(0))": "#N/A",
		// GAMMALN
		"=GAMMALN()":       "GAMMALN requires 1 numeric argument",
		"=GAMMALN(F1)":     "GAMMALN requires 1 numeric argument",
		"=GAMMALN(0)":      "#N/A",
		"=GAMMALN(INT(0))": "#N/A",
		// HARMEAN
		"=HARMEAN()":   "HARMEAN requires at least 1 argument",
		"=HARMEAN(-1)": "#N/A",
		"=HARMEAN(0)":  "#N/A",
		// KURT
		"=KURT()":          "KURT requires at least 1 argument",
		"=KURT(F1,INT(1))": "#DIV/0!",
		// NORM.DIST
		"=NORM.DIST()": "NORM.DIST requires 4 arguments",
		// NORMDIST
		"=NORMDIST()":               "NORMDIST requires 4 arguments",
		"=NORMDIST(\"\",0,0,FALSE)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=NORMDIST(0,\"\",0,FALSE)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=NORMDIST(0,0,\"\",FALSE)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=NORMDIST(0,0,0,\"\")":     "strconv.ParseBool: parsing \"\": invalid syntax",
		"=NORMDIST(0,0,-1,TRUE)":    "#N/A",
		// NORM.INV
		"=NORM.INV()": "NORM.INV requires 3 arguments",
		// NORMINV
		"=NORMINV()":         "NORMINV requires 3 arguments",
		"=NORMINV(\"\",0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=NORMINV(0,\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=NORMINV(0,0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=NORMINV(0,0,-1)":   "#N/A",
		"=NORMINV(-1,0,0)":   "#N/A",
		"=NORMINV(0,0,0)":    "#NUM!",
		// NORM.S.DIST
		"=NORM.S.DIST()": "NORM.S.DIST requires 2 numeric arguments",
		// NORMSDIST
		"=NORMSDIST()": "NORMSDIST requires 1 numeric argument",
		// NORM.S.INV
		"=NORM.S.INV()": "NORM.S.INV requires 1 numeric argument",
		// NORMSINV
		"=NORMSINV()": "NORMSINV requires 1 numeric argument",
		// LARGE
		"=LARGE()":           "LARGE requires 2 arguments",
		"=LARGE(A1:A5,0)":    "k should be > 0",
		"=LARGE(A1:A5,6)":    "k should be <= length of array",
		"=LARGE(A1:A5,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// MAX
		"=MAX()":     "MAX requires at least 1 argument",
		"=MAX(NA())": "#N/A",
		// MAXA
		"=MAXA()":     "MAXA requires at least 1 argument",
		"=MAXA(NA())": "#N/A",
		// MEDIAN
		"=MEDIAN()":      "MEDIAN requires at least 1 argument",
		"=MEDIAN(\"\")":  "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=MEDIAN(D1:D2)": "strconv.ParseFloat: parsing \"Month\": invalid syntax",
		// MIN
		"=MIN()":     "MIN requires at least 1 argument",
		"=MIN(NA())": "#N/A",
		// MINA
		"=MINA()":     "MINA requires at least 1 argument",
		"=MINA(NA())": "#N/A",
		// PERCENTILE.INC
		"=PERCENTILE.INC()": "PERCENTILE.INC requires 2 arguments",
		// PERCENTILE
		"=PERCENTILE()":       "PERCENTILE requires 2 arguments",
		"=PERCENTILE(0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=PERCENTILE(0,-1)":   "#N/A",
		"=PERCENTILE(NA(),1)": "#N/A",
		// PERMUT
		"=PERMUT()":       "PERMUT requires 2 numeric arguments",
		"=PERMUT(\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=PERMUT(0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=PERMUT(6,8)":    "#N/A",
		// PERMUTATIONA
		"=PERMUTATIONA()":       "PERMUTATIONA requires 2 numeric arguments",
		"=PERMUTATIONA(\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=PERMUTATIONA(0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=PERMUTATIONA(-1,0)":   "#N/A",
		"=PERMUTATIONA(0,-1)":   "#N/A",
		// QUARTILE
		"=QUARTILE()":           "QUARTILE requires 2 arguments",
		"=QUARTILE(A1:A4,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=QUARTILE(A1:A4,-1)":   "#NUM!",
		"=QUARTILE(A1:A4,5)":    "#NUM!",
		// QUARTILE.INC
		"=QUARTILE.INC()": "QUARTILE.INC requires 2 arguments",
		// SKEW
		"=SKEW()":     "SKEW requires at least 1 argument",
		"=SKEW(\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=SKEW(0)":    "#DIV/0!",
		// SMALL
		"=SMALL()":           "SMALL requires 2 arguments",
		"=SMALL(A1:A5,0)":    "k should be > 0",
		"=SMALL(A1:A5,6)":    "k should be <= length of array",
		"=SMALL(A1:A5,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// VARP
		"=VARP()":     "VARP requires at least 1 argument",
		"=VARP(\"\")": "#DIV/0!",
		// VAR.P
		"=VAR.P()":     "VAR.P requires at least 1 argument",
		"=VAR.P(\"\")": "#DIV/0!",
		// WEIBULL
		"=WEIBULL()":               "WEIBULL requires 4 arguments",
		"=WEIBULL(\"\",1,1,FALSE)": "#VALUE!",
		"=WEIBULL(1,0,1,FALSE)":    "#N/A",
		"=WEIBULL(1,1,-1,FALSE)":   "#N/A",
		// WEIBULL.DIST
		"=WEIBULL.DIST()":               "WEIBULL.DIST requires 4 arguments",
		"=WEIBULL.DIST(\"\",1,1,FALSE)": "#VALUE!",
		"=WEIBULL.DIST(1,0,1,FALSE)":    "#N/A",
		"=WEIBULL.DIST(1,1,-1,FALSE)":   "#N/A",
		// Z.TEST
		"=Z.TEST(A1)":        "Z.TEST requires at least 2 arguments",
		"=Z.TEST(A1,0,0,0)":  "Z.TEST accepts at most 3 arguments",
		"=Z.TEST(H1,0)":      "#N/A",
		"=Z.TEST(A1,\"\")":   "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=Z.TEST(A1,1)":      "#DIV/0!",
		"=Z.TEST(A1,1,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// ZTEST
		"=ZTEST(A1)":        "ZTEST requires at least 2 arguments",
		"=ZTEST(A1,0,0,0)":  "ZTEST accepts at most 3 arguments",
		"=ZTEST(H1,0)":      "#N/A",
		"=ZTEST(A1,\"\")":   "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=ZTEST(A1,1)":      "#DIV/0!",
		"=ZTEST(A1,1,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
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
		// ISTEXT
		"=ISTEXT()": "ISTEXT requires 1 argument",
		// N
		"=N()":     "N requires 1 argument",
		"=N(NA())": "#N/A",
		// NA
		"=NA()":  "#N/A",
		"=NA(1)": "NA accepts no arguments",
		// SHEET
		"=SHEET(1)": "SHEET accepts no arguments",
		// T
		"=T()":     "T requires 1 argument",
		"=T(NA())": "#N/A",
		// Logical Functions
		// AND
		`=AND("text")`: "strconv.ParseFloat: parsing \"text\": invalid syntax",
		`=AND(A1:B1)`:  "#VALUE!",
		"=AND()":       "AND requires at least 1 argument",
		"=AND(1" + strings.Repeat(",1", 30) + ")": "AND accepts at most 30 arguments",
		// FALSE
		"=FALSE(A1)": "FALSE takes no arguments",
		// IFERROR
		"=IFERROR()": "IFERROR requires 2 arguments",
		// IFNA
		"=IFNA()": "IFNA requires 2 arguments",
		// IFS
		"=IFS()": "IFS requires at least 2 arguments",
		// NOT
		"=NOT()":      "NOT requires 1 argument",
		"=NOT(NOT())": "NOT requires 1 argument",
		"=NOT(\"\")":  "NOT expects 1 boolean or numeric argument",
		// OR
		`=OR("text")`:                            "strconv.ParseFloat: parsing \"text\": invalid syntax",
		`=OR(A1:B1)`:                             "#VALUE!",
		"=OR()":                                  "OR requires at least 1 argument",
		"=OR(1" + strings.Repeat(",1", 30) + ")": "OR accepts at most 30 arguments",
		// TRUE
		"=TRUE(A1)": "TRUE takes no arguments",
		// XOR
		"=XOR()":              "XOR requires at least 1 argument",
		"=XOR(\"text\")":      "#VALUE!",
		"=XOR(XOR(\"text\"))": "#VALUE!",
		// Date and Time Functions
		// DATE
		"=DATE()":               "DATE requires 3 number arguments",
		`=DATE("text",10,21)`:   "DATE requires 3 number arguments",
		`=DATE(2020,"text",21)`: "DATE requires 3 number arguments",
		`=DATE(2020,10,"text")`: "DATE requires 3 number arguments",
		// DATEDIF
		"=DATEDIF()":                  "DATEDIF requires 3 number arguments",
		"=DATEDIF(\"\",\"\",\"\")":    "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=DATEDIF(43891,43101,\"Y\")": "start_date > end_date",
		"=DATEDIF(43101,43891,\"x\")": "DATEDIF has invalid unit",
		// DATEVALUE
		"=DATEVALUE()":             "DATEVALUE requires 1 argument",
		"=DATEVALUE(\"01/01\")":    "#VALUE!", // valid in Excel, which uses years by the system date
		"=DATEVALUE(\"1900-0-0\")": "#VALUE!",
		// DAY
		"=DAY()":         "DAY requires exactly 1 argument",
		"=DAY(-1)":       "DAY only accepts positive argument",
		"=DAY(0,0)":      "DAY requires exactly 1 argument",
		"=DAY(\"text\")": "#VALUE!",
		"=DAY(\"January 25, 2020 9223372036854775808 AM\")":                   "#VALUE!",
		"=DAY(\"January 25, 2020 9223372036854775808:00 AM\")":                "#VALUE!",
		"=DAY(\"January 25, 2020 00:9223372036854775808 AM\")":                "#VALUE!",
		"=DAY(\"January 25, 2020 9223372036854775808:00.0 AM\")":              "#VALUE!",
		"=DAY(\"January 25, 2020 0:1" + strings.Repeat("0", 309) + ".0 AM\")": "#VALUE!",
		"=DAY(\"January 25, 2020 9223372036854775808:00:00 AM\")":             "#VALUE!",
		"=DAY(\"January 25, 2020 0:9223372036854775808:0 AM\")":               "#VALUE!",
		"=DAY(\"January 25, 2020 0:0:1" + strings.Repeat("0", 309) + " AM\")": "#VALUE!",
		"=DAY(\"January 25, 2020 0:61:0 AM\")":                                "#VALUE!",
		"=DAY(\"January 25, 2020 0:00:60 AM\")":                               "#VALUE!",
		"=DAY(\"January 25, 2020 24:00:00\")":                                 "#VALUE!",
		"=DAY(\"January 25, 2020 00:00:10001\")":                              "#VALUE!",
		"=DAY(\"9223372036854775808/25/2020\")":                               "#VALUE!",
		"=DAY(\"01/9223372036854775808/2020\")":                               "#VALUE!",
		"=DAY(\"01/25/9223372036854775808\")":                                 "#VALUE!",
		"=DAY(\"01/25/10000\")":                                               "#VALUE!",
		"=DAY(\"01/25/100\")":                                                 "#VALUE!",
		"=DAY(\"January 9223372036854775808, 2020\")":                         "#VALUE!",
		"=DAY(\"January 25, 9223372036854775808\")":                           "#VALUE!",
		"=DAY(\"January 25, 10000\")":                                         "#VALUE!",
		"=DAY(\"January 25, 100\")":                                           "#VALUE!",
		"=DAY(\"9223372036854775808-25-2020\")":                               "#VALUE!",
		"=DAY(\"01-9223372036854775808-2020\")":                               "#VALUE!",
		"=DAY(\"01-25-9223372036854775808\")":                                 "#VALUE!",
		"=DAY(\"1900-0-0\")":                                                  "#VALUE!",
		"=DAY(\"14-25-1900\")":                                                "#VALUE!",
		"=DAY(\"3-January-9223372036854775808\")":                             "#VALUE!",
		"=DAY(\"9223372036854775808-January-1900\")":                          "#VALUE!",
		"=DAY(\"0-January-1900\")":                                            "#VALUE!",
		// DAYS
		"=DAYS()":       "DAYS requires 2 arguments",
		"=DAYS(\"\",0)": "#VALUE!",
		"=DAYS(0,\"\")": "#VALUE!",
		"=DAYS(NA(),0)": "#VALUE!",
		"=DAYS(0,NA())": "#VALUE!",
		// MONTH
		"=MONTH()":                    "MONTH requires exactly 1 argument",
		"=MONTH(0,0)":                 "MONTH requires exactly 1 argument",
		"=MONTH(-1)":                  "MONTH only accepts positive argument",
		"=MONTH(\"text\")":            "#VALUE!",
		"=MONTH(\"January 25, 100\")": "#VALUE!",
		// YEAR
		"=YEAR()":                    "YEAR requires exactly 1 argument",
		"=YEAR(0,0)":                 "YEAR requires exactly 1 argument",
		"=YEAR(-1)":                  "YEAR only accepts positive argument",
		"=YEAR(\"text\")":            "#VALUE!",
		"=YEAR(\"January 25, 100\")": "#VALUE!",
		// YEARFRAC
		"=YEARFRAC()":                 "YEARFRAC requires 3 or 4 arguments",
		"=YEARFRAC(42005,42094,5)":    "invalid basis",
		"=YEARFRAC(\"\",42094,5)":     "#VALUE!",
		"=YEARFRAC(42005,\"\",5)":     "#VALUE!",
		"=YEARFRAC(42005,42094,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// NOW
		"=NOW(A1)": "NOW accepts no arguments",
		// TIME
		"=TIME()":         "TIME requires 3 number arguments",
		"=TIME(\"\",0,0)": "TIME requires 3 number arguments",
		"=TIME(0,0,-1)":   "#NUM!",
		// TODAY
		"=TODAY(A1)": "TODAY accepts no arguments",
		// Text Functions
		// CHAR
		"=CHAR()":     "CHAR requires 1 argument",
		"=CHAR(-1)":   "#VALUE!",
		"=CHAR(256)":  "#VALUE!",
		"=CHAR(\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// CLEAN
		"=CLEAN()":    "CLEAN requires 1 argument",
		"=CLEAN(1,2)": "CLEAN requires 1 argument",
		// CODE
		"=CODE()":    "CODE requires 1 argument",
		"=CODE(1,2)": "CODE requires 1 argument",
		// CONCAT
		"=CONCAT(MUNIT(2))": "CONCAT requires arguments to be strings",
		// CONCATENATE
		"=CONCATENATE(MUNIT(2))": "CONCATENATE requires arguments to be strings",
		// EXACT
		"=EXACT()":      "EXACT requires 2 arguments",
		"=EXACT(1,2,3)": "EXACT requires 2 arguments",
		// FIXED
		"=FIXED()":         "FIXED requires at least 1 argument",
		"=FIXED(0,1,2,3)":  "FIXED allows at most 3 arguments",
		"=FIXED(\"\")":     "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=FIXED(0,\"\")":   "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=FIXED(0,0,\"\")": "strconv.ParseBool: parsing \"\": invalid syntax",
		// FIND
		"=FIND()":                 "FIND requires at least 2 arguments",
		"=FIND(1,2,3,4)":          "FIND allows at most 3 arguments",
		"=FIND(\"x\",\"\")":       "#VALUE!",
		"=FIND(\"x\",\"x\",-1)":   "#VALUE!",
		"=FIND(\"x\",\"x\",\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// FINDB
		"=FINDB()":                 "FINDB requires at least 2 arguments",
		"=FINDB(1,2,3,4)":          "FINDB allows at most 3 arguments",
		"=FINDB(\"x\",\"\")":       "#VALUE!",
		"=FINDB(\"x\",\"x\",-1)":   "#VALUE!",
		"=FINDB(\"x\",\"x\",\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// LEFT
		"=LEFT()":          "LEFT requires at least 1 argument",
		"=LEFT(\"\",2,3)":  "LEFT allows at most 2 arguments",
		"=LEFT(\"\",\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=LEFT(\"\",-1)":   "#VALUE!",
		// LEFTB
		"=LEFTB()":          "LEFTB requires at least 1 argument",
		"=LEFTB(\"\",2,3)":  "LEFTB allows at most 2 arguments",
		"=LEFTB(\"\",\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=LEFTB(\"\",-1)":   "#VALUE!",
		// LEN
		"=LEN()": "LEN requires 1 string argument",
		// LENB
		"=LENB()": "LENB requires 1 string argument",
		// LOWER
		"=LOWER()":    "LOWER requires 1 argument",
		"=LOWER(1,2)": "LOWER requires 1 argument",
		// MID
		"=MID()":            "MID requires 3 arguments",
		"=MID(\"\",-1,1)":   "#VALUE!",
		"=MID(\"\",\"\",1)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=MID(\"\",1,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// MIDB
		"=MIDB()":            "MIDB requires 3 arguments",
		"=MIDB(\"\",-1,1)":   "#VALUE!",
		"=MIDB(\"\",\"\",1)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=MIDB(\"\",1,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// PROPER
		"=PROPER()":    "PROPER requires 1 argument",
		"=PROPER(1,2)": "PROPER requires 1 argument",
		// REPLACE
		"=REPLACE()":                           "REPLACE requires 4 arguments",
		"=REPLACE(\"text\",0,4,\"string\")":    "#VALUE!",
		"=REPLACE(\"text\",\"\",0,\"string\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=REPLACE(\"text\",1,\"\",\"string\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// REPLACEB
		"=REPLACEB()":                           "REPLACEB requires 4 arguments",
		"=REPLACEB(\"text\",0,4,\"string\")":    "#VALUE!",
		"=REPLACEB(\"text\",\"\",0,\"string\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=REPLACEB(\"text\",1,\"\",\"string\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// REPT
		"=REPT()":            "REPT requires 2 arguments",
		"=REPT(INT(0),2)":    "REPT requires first argument to be a string",
		"=REPT(\"*\",\"*\")": "REPT requires second argument to be a number",
		"=REPT(\"*\",-1)":    "REPT requires second argument to be >= 0",
		// RIGHT
		"=RIGHT()":          "RIGHT requires at least 1 argument",
		"=RIGHT(\"\",2,3)":  "RIGHT allows at most 2 arguments",
		"=RIGHT(\"\",\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=RIGHT(\"\",-1)":   "#VALUE!",
		// RIGHTB
		"=RIGHTB()":          "RIGHTB requires at least 1 argument",
		"=RIGHTB(\"\",2,3)":  "RIGHTB allows at most 2 arguments",
		"=RIGHTB(\"\",\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=RIGHTB(\"\",-1)":   "#VALUE!",
		// SUBSTITUTE
		"=SUBSTITUTE()":                    "SUBSTITUTE requires 3 or 4 arguments",
		"=SUBSTITUTE(\"\",\"\",\"\",\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=SUBSTITUTE(\"\",\"\",\"\",0)":    "instance_num should be > 0",
		// TRIM
		"=TRIM()":    "TRIM requires 1 argument",
		"=TRIM(1,2)": "TRIM requires 1 argument",
		// UNICHAR
		"=UNICHAR()":      "UNICHAR requires 1 argument",
		"=UNICHAR(\"\")":  "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=UNICHAR(55296)": "#VALUE!",
		"=UNICHAR(0)":     "#VALUE!",
		// UNICODE
		"=UNICODE()":     "UNICODE requires 1 argument",
		"=UNICODE(\"\")": "#VALUE!",
		// UPPER
		"=UPPER()":    "UPPER requires 1 argument",
		"=UPPER(1,2)": "UPPER requires 1 argument",
		// Conditional Functions
		// IF
		"=IF()":        "IF requires at least 1 argument",
		"=IF(0,1,2,3)": "IF accepts at most 3 arguments",
		"=IF(D1,1,2)":  "strconv.ParseBool: parsing \"Month\": invalid syntax",
		// Excel Lookup and Reference Functions
		// ADDRESS
		"=ADDRESS()":                        "ADDRESS requires at least 2 arguments",
		"=ADDRESS(1,1,1,TRUE,\"Sheet1\",0)": "ADDRESS requires at most 5 arguments",
		"=ADDRESS(\"\",1,1,TRUE)":           "#VALUE!",
		"=ADDRESS(1,\"\",1,TRUE)":           "#VALUE!",
		"=ADDRESS(1,1,\"\",TRUE)":           "#VALUE!",
		"=ADDRESS(1,1,1,\"\")":              "#VALUE!",
		"=ADDRESS(1,1,0,TRUE)":              "#NUM!",
		"=ADDRESS(1,16385,2,TRUE)":          "#VALUE!",
		"=ADDRESS(1,16385,3,TRUE)":          "#VALUE!",
		"=ADDRESS(1048576,1,1,TRUE)":        "#VALUE!",
		// CHOOSE
		"=CHOOSE()":                "CHOOSE requires 2 arguments",
		"=CHOOSE(\"index_num\",0)": "CHOOSE requires first argument of type number",
		"=CHOOSE(2,0)":             "index_num should be <= to the number of values",
		// COLUMN
		"=COLUMN(1,2)":          "COLUMN requires at most 1 argument",
		"=COLUMN(\"\")":         "invalid reference",
		"=COLUMN(Sheet1)":       "invalid column name \"Sheet1\"",
		"=COLUMN(Sheet1!A1!B1)": "invalid column name \"Sheet1\"",
		// COLUMNS
		"=COLUMNS()":              "COLUMNS requires 1 argument",
		"=COLUMNS(1)":             "invalid reference",
		"=COLUMNS(\"\")":          "invalid reference",
		"=COLUMNS(Sheet1)":        "invalid column name \"Sheet1\"",
		"=COLUMNS(Sheet1!A1!B1)":  "invalid column name \"Sheet1\"",
		"=COLUMNS(Sheet1!Sheet1)": "invalid column name \"Sheet1\"",
		// HLOOKUP
		"=HLOOKUP()":                     "HLOOKUP requires at least 3 arguments",
		"=HLOOKUP(D2,D1,1,FALSE)":        "HLOOKUP requires second argument of table array",
		"=HLOOKUP(D2,D:D,FALSE,FALSE)":   "HLOOKUP requires numeric row argument",
		"=HLOOKUP(D2,D:D,1,FALSE,FALSE)": "HLOOKUP requires at most 4 arguments",
		"=HLOOKUP(D2,D:D,1,2)":           "strconv.ParseBool: parsing \"2\": invalid syntax",
		"=HLOOKUP(D2,D10:D10,1,FALSE)":   "HLOOKUP no result found",
		"=HLOOKUP(D2,D2:D3,4,FALSE)":     "HLOOKUP has invalid row index",
		"=HLOOKUP(D2,C:C,1,FALSE)":       "HLOOKUP no result found",
		"=HLOOKUP(ISNUMBER(1),F3:F9,1)":  "HLOOKUP no result found",
		"=HLOOKUP(INT(1),E2:E9,1)":       "HLOOKUP no result found",
		"=HLOOKUP(MUNIT(2),MUNIT(3),1)":  "HLOOKUP no result found",
		"=HLOOKUP(A1:B2,B2:B3,1)":        "HLOOKUP no result found",
		// MATCH
		"=MATCH()":              "MATCH requires 1 or 2 arguments",
		"=MATCH(0,A1:A1,0,0)":   "MATCH requires 1 or 2 arguments",
		"=MATCH(0,A1:A1,\"x\")": "MATCH requires numeric match_type argument",
		"=MATCH(0,A1)":          "MATCH arguments lookup_array should be one-dimensional array",
		"=MATCH(0,A1:B1)":       "MATCH arguments lookup_array should be one-dimensional array",
		// TRANSPOSE
		"=TRANSPOSE()": "TRANSPOSE requires 1 argument",
		// VLOOKUP
		"=VLOOKUP()":                     "VLOOKUP requires at least 3 arguments",
		"=VLOOKUP(D2,D1,1,FALSE)":        "VLOOKUP requires second argument of table array",
		"=VLOOKUP(D2,D:D,FALSE,FALSE)":   "VLOOKUP requires numeric col argument",
		"=VLOOKUP(D2,D:D,1,FALSE,FALSE)": "VLOOKUP requires at most 4 arguments",
		"=VLOOKUP(D2,D:D,1,2)":           "strconv.ParseBool: parsing \"2\": invalid syntax",
		"=VLOOKUP(D2,D10:D10,1,FALSE)":   "VLOOKUP no result found",
		"=VLOOKUP(D2,D:D,2,FALSE)":       "VLOOKUP has invalid column index",
		"=VLOOKUP(D2,C:C,1,FALSE)":       "VLOOKUP no result found",
		"=VLOOKUP(ISNUMBER(1),F3:F9,1)":  "VLOOKUP no result found",
		"=VLOOKUP(INT(1),E2:E9,1)":       "VLOOKUP no result found",
		"=VLOOKUP(MUNIT(2),MUNIT(3),1)":  "VLOOKUP no result found",
		"=VLOOKUP(1,G1:H2,1,FALSE)":      "VLOOKUP no result found",
		// LOOKUP
		"=LOOKUP()":                     "LOOKUP requires at least 2 arguments",
		"=LOOKUP(D2,D1,D2)":             "LOOKUP requires second argument of table array",
		"=LOOKUP(D2,D1,D2,FALSE)":       "LOOKUP requires at most 3 arguments",
		"=LOOKUP(1,MUNIT(0))":           "LOOKUP requires not empty range as second argument",
		"=LOOKUP(D1,MUNIT(1),MUNIT(1))": "LOOKUP no result found",
		// ROW
		"=ROW(1,2)":          "ROW requires at most 1 argument",
		"=ROW(\"\")":         "invalid reference",
		"=ROW(Sheet1)":       "invalid column name \"Sheet1\"",
		"=ROW(Sheet1!A1!B1)": "invalid column name \"Sheet1\"",
		// ROWS
		"=ROWS()":              "ROWS requires 1 argument",
		"=ROWS(1)":             "invalid reference",
		"=ROWS(\"\")":          "invalid reference",
		"=ROWS(Sheet1)":        "invalid column name \"Sheet1\"",
		"=ROWS(Sheet1!A1!B1)":  "invalid column name \"Sheet1\"",
		"=ROWS(Sheet1!Sheet1)": "invalid column name \"Sheet1\"",
		// Web Functions
		// ENCODEURL
		"=ENCODEURL()": "ENCODEURL requires 1 argument",
		// Financial Functions
		// ACCRINTM
		"=ACCRINTM()": "ACCRINTM requires 4 or 5 arguments",
		"=ACCRINTM(\"\",\"01/01/2012\",8%,10000)":                "#VALUE!",
		"=ACCRINTM(\"01/01/2012\",\"\",8%,10000)":                "#VALUE!",
		"=ACCRINTM(\"12/31/2012\",\"01/01/2012\",8%,10000)":      "#NUM!",
		"=ACCRINTM(\"01/01/2012\",\"12/31/2012\",\"\",10000)":    "#NUM!",
		"=ACCRINTM(\"01/01/2012\",\"12/31/2012\",8%,\"\",10000)": "#NUM!",
		"=ACCRINTM(\"01/01/2012\",\"12/31/2012\",8%,-1,10000)":   "#NUM!",
		"=ACCRINTM(\"01/01/2012\",\"12/31/2012\",8%,10000,\"\")": "#NUM!",
		"=ACCRINTM(\"01/01/2012\",\"12/31/2012\",8%,10000,5)":    "invalid basis",
		// AMORDEGRC
		"=AMORDEGRC()": "AMORDEGRC requires 6 or 7 arguments",
		"=AMORDEGRC(\"\",\"01/01/2015\",\"09/30/2015\",20,1,20%)":     "AMORDEGRC requires cost to be number argument",
		"=AMORDEGRC(-1,\"01/01/2015\",\"09/30/2015\",20,1,20%)":       "AMORDEGRC requires cost >= 0",
		"=AMORDEGRC(150,\"\",\"09/30/2015\",20,1,20%)":                "#VALUE!",
		"=AMORDEGRC(150,\"01/01/2015\",\"\",20,1,20%)":                "#VALUE!",
		"=AMORDEGRC(150,\"09/30/2015\",\"01/01/2015\",20,1,20%)":      "#NUM!",
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",\"\",1,20%)":    "#NUM!",
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",-1,1,20%)":      "#NUM!",
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",20,\"\",20%)":   "#NUM!",
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",20,-1,20%)":     "#NUM!",
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",20,1,\"\")":     "#NUM!",
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",20,1,-1)":       "#NUM!",
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",20,1,20%,\"\")": "#NUM!",
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",20,1,50%)":      "AMORDEGRC requires rate to be < 0.5",
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",20,1,20%,5)":    "invalid basis",
		// AMORLINC
		"=AMORLINC()": "AMORLINC requires 6 or 7 arguments",
		"=AMORLINC(\"\",\"01/01/2015\",\"09/30/2015\",20,1,20%)":     "AMORLINC requires cost to be number argument",
		"=AMORLINC(-1,\"01/01/2015\",\"09/30/2015\",20,1,20%)":       "AMORLINC requires cost >= 0",
		"=AMORLINC(150,\"\",\"09/30/2015\",20,1,20%)":                "#VALUE!",
		"=AMORLINC(150,\"01/01/2015\",\"\",20,1,20%)":                "#VALUE!",
		"=AMORLINC(150,\"09/30/2015\",\"01/01/2015\",20,1,20%)":      "#NUM!",
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",\"\",1,20%)":    "#NUM!",
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",-1,1,20%)":      "#NUM!",
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",20,\"\",20%)":   "#NUM!",
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",20,-1,20%)":     "#NUM!",
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",20,1,\"\")":     "#NUM!",
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",20,1,-1)":       "#NUM!",
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",20,1,20%,\"\")": "#NUM!",
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",20,1,20%,5)":    "invalid basis",
		// CUMIPMT
		"=CUMIPMT()":               "CUMIPMT requires 6 arguments",
		"=CUMIPMT(0,0,0,0,0,2)":    "#N/A",
		"=CUMIPMT(0,0,0,-1,0,0)":   "#N/A",
		"=CUMIPMT(0,0,0,1,0,0)":    "#N/A",
		"=CUMIPMT(\"\",0,0,0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=CUMIPMT(0,\"\",0,0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=CUMIPMT(0,0,\"\",0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=CUMIPMT(0,0,0,\"\",0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=CUMIPMT(0,0,0,0,\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=CUMIPMT(0,0,0,0,0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// CUMPRINC
		"=CUMPRINC()":               "CUMPRINC requires 6 arguments",
		"=CUMPRINC(0,0,0,0,0,2)":    "#N/A",
		"=CUMPRINC(0,0,0,-1,0,0)":   "#N/A",
		"=CUMPRINC(0,0,0,1,0,0)":    "#N/A",
		"=CUMPRINC(\"\",0,0,0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=CUMPRINC(0,\"\",0,0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=CUMPRINC(0,0,\"\",0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=CUMPRINC(0,0,0,\"\",0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=CUMPRINC(0,0,0,0,\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=CUMPRINC(0,0,0,0,0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// DB
		"=DB()":             "DB requires at least 4 arguments",
		"=DB(0,0,0,0,0,0)":  "DB allows at most 5 arguments",
		"=DB(-1,0,0,0)":     "#N/A",
		"=DB(\"\",0,0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=DB(0,\"\",0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=DB(0,0,\"\",0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=DB(0,0,0,\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=DB(0,0,0,0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// DDB
		"=DDB()":             "DDB requires at least 4 arguments",
		"=DDB(0,0,0,0,0,0)":  "DDB allows at most 5 arguments",
		"=DDB(-1,0,0,0)":     "#N/A",
		"=DDB(\"\",0,0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=DDB(0,\"\",0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=DDB(0,0,\"\",0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=DDB(0,0,0,\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=DDB(0,0,0,0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// DISC
		"=DISC()":                                          "DISC requires 4 or 5 arguments",
		"=DISC(\"\",\"03/31/2021\",95,100)":                "#VALUE!",
		"=DISC(\"04/01/2016\",\"\",95,100)":                "#VALUE!",
		"=DISC(\"04/01/2016\",\"03/31/2021\",\"\",100)":    "#VALUE!",
		"=DISC(\"04/01/2016\",\"03/31/2021\",95,\"\")":     "#VALUE!",
		"=DISC(\"04/01/2016\",\"03/31/2021\",95,100,\"\")": "#NUM!",
		"=DISC(\"03/31/2021\",\"04/01/2016\",95,100)":      "DISC requires maturity > settlement",
		"=DISC(\"04/01/2016\",\"03/31/2021\",0,100)":       "DISC requires pr > 0",
		"=DISC(\"04/01/2016\",\"03/31/2021\",95,0)":        "DISC requires redemption > 0",
		"=DISC(\"04/01/2016\",\"03/31/2021\",95,100,5)":    "invalid basis",
		// DOLLARDE
		"=DOLLARDE()":       "DOLLARDE requires 2 arguments",
		"=DOLLARDE(\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=DOLLARDE(0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=DOLLARDE(0,-1)":   "#NUM!",
		"=DOLLARDE(0,0)":    "#DIV/0!",
		// DOLLARFR
		"=DOLLARFR()":       "DOLLARFR requires 2 arguments",
		"=DOLLARFR(\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=DOLLARFR(0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=DOLLARFR(0,-1)":   "#NUM!",
		"=DOLLARFR(0,0)":    "#DIV/0!",
		// EFFECT
		"=EFFECT()":       "EFFECT requires 2 arguments",
		"=EFFECT(\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=EFFECT(0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=EFFECT(0,0)":    "#NUM!",
		"=EFFECT(1,0)":    "#NUM!",
		// FV
		"=FV()":              "FV requires at least 3 arguments",
		"=FV(0,0,0,0,0,0,0)": "FV allows at most 5 arguments",
		"=FV(0,0,0,0,2)":     "#N/A",
		"=FV(\"\",0,0,0,0)":  "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=FV(0,\"\",0,0,0)":  "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=FV(0,0,\"\",0,0)":  "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=FV(0,0,0,\"\",0)":  "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=FV(0,0,0,0,\"\")":  "strconv.ParseFloat: parsing \"\": invalid syntax",
		// FVSCHEDULE
		"=FVSCHEDULE()":        "FVSCHEDULE requires 2 arguments",
		"=FVSCHEDULE(\"\",0)":  "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=FVSCHEDULE(0,\"x\")": "strconv.ParseFloat: parsing \"x\": invalid syntax",
		// INTRATE
		"=INTRATE()":                                          "INTRATE requires 4 or 5 arguments",
		"=INTRATE(\"\",\"03/31/2021\",95,100)":                "#VALUE!",
		"=INTRATE(\"04/01/2016\",\"\",95,100)":                "#VALUE!",
		"=INTRATE(\"04/01/2016\",\"03/31/2021\",\"\",100)":    "#VALUE!",
		"=INTRATE(\"04/01/2016\",\"03/31/2021\",95,\"\")":     "#VALUE!",
		"=INTRATE(\"04/01/2016\",\"03/31/2021\",95,100,\"\")": "#NUM!",
		"=INTRATE(\"03/31/2021\",\"04/01/2016\",95,100)":      "INTRATE requires maturity > settlement",
		"=INTRATE(\"04/01/2016\",\"03/31/2021\",0,100)":       "INTRATE requires investment > 0",
		"=INTRATE(\"04/01/2016\",\"03/31/2021\",95,0)":        "INTRATE requires redemption > 0",
		"=INTRATE(\"04/01/2016\",\"03/31/2021\",95,100,5)":    "invalid basis",
		// IPMT
		"=IPMT()":               "IPMT requires at least 4 arguments",
		"=IPMT(0,0,0,0,0,0,0)":  "IPMT allows at most 6 arguments",
		"=IPMT(0,0,0,0,0,2)":    "#N/A",
		"=IPMT(0,-1,0,0,0,0)":   "#N/A",
		"=IPMT(0,1,0,0,0,0)":    "#N/A",
		"=IPMT(\"\",0,0,0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=IPMT(0,\"\",0,0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=IPMT(0,0,\"\",0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=IPMT(0,0,0,\"\",0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=IPMT(0,0,0,0,\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=IPMT(0,0,0,0,0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// ISPMT
		"=ISPMT()":           "ISPMT requires 4 arguments",
		"=ISPMT(\"\",0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=ISPMT(0,\"\",0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=ISPMT(0,0,\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=ISPMT(0,0,0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// NOMINAL
		"=NOMINAL()":       "NOMINAL requires 2 arguments",
		"=NOMINAL(\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=NOMINAL(0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=NOMINAL(0,0)":    "#NUM!",
		"=NOMINAL(1,0)":    "#NUM!",
		// NPER
		"=NPER()":             "NPER requires at least 3 arguments",
		"=NPER(0,0,0,0,0,0)":  "NPER allows at most 5 arguments",
		"=NPER(0,0,0)":        "#NUM!",
		"=NPER(0,0,0,0,2)":    "#N/A",
		"=NPER(\"\",0,0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=NPER(0,\"\",0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=NPER(0,0,\"\",0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=NPER(0,0,0,\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=NPER(0,0,0,0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// NPV
		"=NPV()":       "NPV requires at least 2 arguments",
		"=NPV(\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// PDURATION
		"=PDURATION()":         "PDURATION requires 3 arguments",
		"=PDURATION(\"\",0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=PDURATION(0,\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=PDURATION(0,0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=PDURATION(0,0,0)":    "#NUM!",
		// PMT
		"=PMT()":             "PMT requires at least 3 arguments",
		"=PMT(0,0,0,0,0,0)":  "PMT allows at most 5 arguments",
		"=PMT(0,0,0,0,2)":    "#N/A",
		"=PMT(\"\",0,0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=PMT(0,\"\",0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=PMT(0,0,\"\",0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=PMT(0,0,0,\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=PMT(0,0,0,0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// PPMT
		"=PPMT()":               "PPMT requires at least 4 arguments",
		"=PPMT(0,0,0,0,0,0,0)":  "PPMT allows at most 6 arguments",
		"=PPMT(0,0,0,0,0,2)":    "#N/A",
		"=PPMT(0,-1,0,0,0,0)":   "#N/A",
		"=PPMT(0,1,0,0,0,0)":    "#N/A",
		"=PPMT(\"\",0,0,0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=PPMT(0,\"\",0,0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=PPMT(0,0,\"\",0,0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=PPMT(0,0,0,\"\",0,0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=PPMT(0,0,0,0,\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=PPMT(0,0,0,0,0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		// PRICEDISC
		"=PRICEDISC()":                                          "PRICEDISC requires 4 or 5 arguments",
		"=PRICEDISC(\"\",\"03/31/2021\",95,100)":                "#VALUE!",
		"=PRICEDISC(\"04/01/2016\",\"\",95,100)":                "#VALUE!",
		"=PRICEDISC(\"04/01/2016\",\"03/31/2021\",\"\",100)":    "#VALUE!",
		"=PRICEDISC(\"04/01/2016\",\"03/31/2021\",95,\"\")":     "#VALUE!",
		"=PRICEDISC(\"04/01/2016\",\"03/31/2021\",95,100,\"\")": "#NUM!",
		"=PRICEDISC(\"03/31/2021\",\"04/01/2016\",95,100)":      "PRICEDISC requires maturity > settlement",
		"=PRICEDISC(\"04/01/2016\",\"03/31/2021\",0,100)":       "PRICEDISC requires discount > 0",
		"=PRICEDISC(\"04/01/2016\",\"03/31/2021\",95,0)":        "PRICEDISC requires redemption > 0",
		"=PRICEDISC(\"04/01/2016\",\"03/31/2021\",95,100,5)":    "invalid basis",
		// RRI
		"=RRI()":               "RRI requires 3 arguments",
		"=RRI(\"\",\"\",\"\")": "#NUM!",
		"=RRI(0,10000,15000)":  "RRI requires nper argument to be > 0",
		"=RRI(10,0,15000)":     "RRI requires pv argument to be > 0",
		"=RRI(10,10000,-1)":    "RRI requires fv argument to be >= 0",
		// SLN
		"=SLN()":               "SLN requires 3 arguments",
		"=SLN(\"\",\"\",\"\")": "#NUM!",
		"=SLN(10000,1000,0)":   "SLN requires life argument to be > 0",
		// SYD
		"=SYD()":                    "SYD requires 4 arguments",
		"=SYD(\"\",\"\",\"\",\"\")": "#NUM!",
		"=SYD(10000,1000,0,1)":      "SYD requires life argument to be > 0",
		"=SYD(10000,1000,5,0)":      "SYD requires per argument to be > 0",
		"=SYD(10000,1000,1,5)":      "#NUM!",
	}
	for formula, expected := range mathCalcError {
		f := prepareCalcData(cellData)
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
		// IMPRODUCT
		"=IMPRODUCT(Sheet1!A1:Sheet1!A1:A2,A2)": "4",
		// SUM
		"=A1/A3":                          "0.333333333333333",
		"=SUM(A1:A2)":                     "3",
		"=SUM(Sheet1!A1,A2)":              "3",
		"=(-2-SUM(-4+A2))*5":              "0",
		"=SUM(Sheet1!A1:Sheet1!A1:A2,A2)": "5",
		"=SUM(A1,A2,A3)*SUM(2,3)":         "30",
		"=1+SUM(SUM(A1+A2/A3)*(2-3),2)":   "1.33333333333333",
		"=A1/A2/SUM(A1:A2:B1)":            "0.0416666666666667",
		"=A1/A2/SUM(A1:A2:B1)*A3":         "0.125",
		"=SUM(B1:D1)":                     "4",
		"=SUM(\"X\")":                     "0",
	}
	for formula, expected := range referenceCalc {
		f := prepareCalcData(cellData)
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
		f := prepareCalcData(cellData)
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.EqualError(t, err, expected)
		assert.Equal(t, "", result, formula)
	}

	volatileFuncs := []string{
		"=NOW()",
		"=RAND()",
		"=RANDBETWEEN(1,2)",
		"=TODAY()",
	}
	for _, formula := range volatileFuncs {
		f := prepareCalcData(cellData)
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		_, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err)
	}

	// Test get calculated cell value on not formula cell.
	f := prepareCalcData(cellData)
	result, err := f.CalcCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "", result)
	// Test get calculated cell value on not exists worksheet.
	f = prepareCalcData(cellData)
	_, err = f.CalcCellValue("SheetN", "A1")
	assert.EqualError(t, err, "sheet SheetN is not exist")
	// Test get calculated cell value with not support formula.
	f = prepareCalcData(cellData)
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

func TestCalcWithDefinedName(t *testing.T) {
	cellData := [][]interface{}{
		{"A1 value", "B1 value", nil},
	}
	f := prepareCalcData(cellData)
	assert.NoError(t, f.SetDefinedName(&DefinedName{Name: "defined_name1", RefersTo: "Sheet1!A1", Scope: "Workbook"}))
	assert.NoError(t, f.SetDefinedName(&DefinedName{Name: "defined_name1", RefersTo: "Sheet1!B1", Scope: "Sheet1"}))

	assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "=defined_name1"))
	result, err := f.CalcCellValue("Sheet1", "C1")
	assert.NoError(t, err)
	// DefinedName with scope WorkSheet takes precedence over DefinedName with scope Workbook, so we should get B1 value
	assert.Equal(t, "B1 value", result, "=defined_name1")

	assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "=CONCATENATE(\"<\",defined_name1,\">\")"))
	result, err = f.CalcCellValue("Sheet1", "C1")
	assert.NoError(t, err)
	assert.Equal(t, "<B1 value>", result, "=defined_name1")

}

func TestCalcISBLANK(t *testing.T) {
	argsList := list.New()
	argsList.PushBack(formulaArg{
		Type: ArgUnknown,
	})
	fn := formulaFuncs{}
	result := fn.ISBLANK(argsList)
	assert.Equal(t, result.String, "TRUE")
	assert.Empty(t, result.Error)
}

func TestCalcAND(t *testing.T) {
	argsList := list.New()
	argsList.PushBack(formulaArg{
		Type: ArgUnknown,
	})
	fn := formulaFuncs{}
	result := fn.AND(argsList)
	assert.Equal(t, result.String, "")
	assert.Empty(t, result.Error)
}

func TestCalcOR(t *testing.T) {
	argsList := list.New()
	argsList.PushBack(formulaArg{
		Type: ArgUnknown,
	})
	fn := formulaFuncs{}
	result := fn.OR(argsList)
	assert.Equal(t, result.String, "FALSE")
	assert.Empty(t, result.Error)
}

func TestCalcDet(t *testing.T) {
	assert.Equal(t, det([][]float64{
		{1, 2, 3, 4},
		{2, 3, 4, 5},
		{3, 4, 5, 6},
		{4, 5, 6, 7},
	}), float64(0))
}

func TestCalcToBool(t *testing.T) {
	b := newBoolFormulaArg(true).ToBool()
	assert.Equal(t, b.Boolean, true)
	assert.Equal(t, b.Number, 1.0)
}

func TestCalcToList(t *testing.T) {
	assert.Equal(t, []formulaArg(nil), newEmptyFormulaArg().ToList())
	formulaList := []formulaArg{newEmptyFormulaArg()}
	assert.Equal(t, formulaList, newListFormulaArg(formulaList).ToList())
}

func TestCalcCompareFormulaArg(t *testing.T) {
	assert.Equal(t, compareFormulaArg(newEmptyFormulaArg(), newEmptyFormulaArg(), false, false), criteriaEq)
	lhs := newListFormulaArg([]formulaArg{newEmptyFormulaArg()})
	rhs := newListFormulaArg([]formulaArg{newEmptyFormulaArg(), newEmptyFormulaArg()})
	assert.Equal(t, compareFormulaArg(lhs, rhs, false, false), criteriaL)
	assert.Equal(t, compareFormulaArg(rhs, lhs, false, false), criteriaG)

	lhs = newListFormulaArg([]formulaArg{newBoolFormulaArg(true)})
	rhs = newListFormulaArg([]formulaArg{newBoolFormulaArg(true)})
	assert.Equal(t, compareFormulaArg(lhs, rhs, false, false), criteriaEq)

	assert.Equal(t, compareFormulaArg(formulaArg{Type: ArgUnknown}, formulaArg{Type: ArgUnknown}, false, false), criteriaErr)
}

func TestCalcMatchPattern(t *testing.T) {
	assert.True(t, matchPattern("", ""))
	assert.True(t, matchPattern("file/*", "file/abc/bcd/def"))
	assert.True(t, matchPattern("*", ""))
	assert.False(t, matchPattern("file/?", "file/abc/bcd/def"))
}

func TestCalcTRANSPOSE(t *testing.T) {
	cellData := [][]interface{}{
		{"a", "d"},
		{"b", "e"},
		{"c", "f"},
	}
	formula := "=TRANSPOSE(A1:A3)"
	f := prepareCalcData(cellData)
	formulaType, ref := STCellFormulaTypeArray, "D1:F2"
	assert.NoError(t, f.SetCellFormula("Sheet1", "D1", formula,
		FormulaOpts{Ref: &ref, Type: &formulaType}))
	_, err := f.CalcCellValue("Sheet1", "D1")
	assert.NoError(t, err, formula)
}

func TestCalcVLOOKUP(t *testing.T) {
	cellData := [][]interface{}{
		{nil, nil, nil, nil, nil, nil},
		{nil, "Score", "Grade", nil, nil, nil},
		{nil, 0, "F", nil, "Score", 85},
		{nil, 60, "D", nil, "Grade"},
		{nil, 70, "C", nil, nil, nil},
		{nil, 80, "b", nil, nil, nil},
		{nil, 90, "A", nil, nil, nil},
		{nil, 85, "B", nil, nil, nil},
		{nil, nil, nil, nil, nil, nil},
	}
	f := prepareCalcData(cellData)
	calc := map[string]string{
		"=VLOOKUP(F3,B3:C8,2)":       "b",
		"=VLOOKUP(F3,B3:C8,2,TRUE)":  "b",
		"=VLOOKUP(F3,B3:C8,2,FALSE)": "B",
	}
	for formula, expected := range calc {
		assert.NoError(t, f.SetCellFormula("Sheet1", "F4", formula))
		result, err := f.CalcCellValue("Sheet1", "F4")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	calcError := map[string]string{
		"=VLOOKUP(INT(1),C3:C3,1,FALSE)": "VLOOKUP no result found",
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "F4", formula))
		result, err := f.CalcCellValue("Sheet1", "F4")
		assert.EqualError(t, err, expected, formula)
		assert.Equal(t, "", result, formula)
	}
}

func TestCalcBoolean(t *testing.T) {
	cellData := [][]interface{}{
		{0.5, "TRUE", -0.5, "FALSE"},
	}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=AVERAGEA(A1:C1)":  "0.333333333333333",
		"=MAX(0.5,B1)":      "0.5",
		"=MAX(A1:B1)":       "0.5",
		"=MAXA(A1:B1)":      "1",
		"=MAXA(0.5,B1)":     "1",
		"=MIN(-0.5,D1)":     "-0.5",
		"=MIN(C1:D1)":       "-0.5",
		"=MINA(C1:D1)":      "-0.5",
		"=MINA(-0.5,D1)":    "-0.5",
		"=STDEV(A1:C1)":     "0.707106781186548",
		"=STDEV(A1,B1,C1)":  "0.707106781186548",
		"=STDEVA(A1:C1,B1)": "0.707106781186548",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "B10", formula))
		result, err := f.CalcCellValue("Sheet1", "B10")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
}

func TestCalcHLOOKUP(t *testing.T) {
	cellData := [][]interface{}{
		{"Example Result Table"},
		{nil, "A", "B", "C", "E", "F"},
		{"Math", .58, .9, .67, .76, .8},
		{"French", .61, .71, .59, .59, .76},
		{"Physics", .75, .45, .39, .52, .69},
		{"Biology", .39, .55, .77, .61, .45},
		{},
		{"Individual Student Score"},
		{"Student:", "Biology Score:"},
		{"E"},
	}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=HLOOKUP(A10,A2:F6,5,FALSE)":  "0.61",
		"=HLOOKUP(D3,D3:D3,1,TRUE)":    "0.67",
		"=HLOOKUP(F3,D3:F3,1,TRUE)":    "0.8",
		"=HLOOKUP(A5,A2:F2,1,TRUE)":    "F",
		"=HLOOKUP(\"D\",A2:F2,1,TRUE)": "C",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "B10", formula))
		result, err := f.CalcCellValue("Sheet1", "B10")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	calcError := map[string]string{
		"=HLOOKUP(INT(1),A3:A3,1,FALSE)": "HLOOKUP no result found",
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "B10", formula))
		result, err := f.CalcCellValue("Sheet1", "B10")
		assert.EqualError(t, err, expected, formula)
		assert.Equal(t, "", result, formula)
	}
}

func TestCalcIRR(t *testing.T) {
	cellData := [][]interface{}{{-1}, {0.2}, {0.24}, {0.288}, {0.3456}, {0.4147}}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=IRR(A1:A4)":      "-0.136189509034157",
		"=IRR(A1:A6)":      "0.130575760006905",
		"=IRR(A1:A4,-0.1)": "-0.136189514994621",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", formula))
		result, err := f.CalcCellValue("Sheet1", "B1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	calcError := map[string]string{
		"=IRR()":       "IRR requires at least 1 argument",
		"=IRR(0,0,0)":  "IRR allows at most 2 arguments",
		"=IRR(0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=IRR(A2:A3)":  "#NUM!",
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", formula))
		result, err := f.CalcCellValue("Sheet1", "B1")
		assert.EqualError(t, err, expected, formula)
		assert.Equal(t, "", result, formula)
	}
}

func TestCalcMIRR(t *testing.T) {
	cellData := [][]interface{}{{-100}, {18}, {22.5}, {28}, {35.5}, {45}}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=MIRR(A1:A6,0.055,0.05)": "0.1000268752662",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", formula))
		result, err := f.CalcCellValue("Sheet1", "B1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	calcError := map[string]string{
		"=MIRR()":             "MIRR requires 3 arguments",
		"=MIRR(A1:A5,\"\",0)": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=MIRR(A1:A5,0,\"\")": "strconv.ParseFloat: parsing \"\": invalid syntax",
		"=MIRR(B1:B5,0,0)":    "#DIV/0!",
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", formula))
		result, err := f.CalcCellValue("Sheet1", "B1")
		assert.EqualError(t, err, expected, formula)
		assert.Equal(t, "", result, formula)
	}
}

func TestCalcMATCH(t *testing.T) {
	f := NewFile()
	for cell, row := range map[string][]interface{}{
		"A1": {"cccc", 7, 4, 16},
		"A2": {"dddd", 2, 6, 11},
		"A3": {"aaaa", 4, 7, 10},
		"A4": {"bbbb", 1, 10, 7},
		"A5": {"eeee", 8, 11, 6},
		"A6": {nil, 11, 16, 4},
	} {
		assert.NoError(t, f.SetSheetRow("Sheet1", cell, &row))
	}
	formulaList := map[string]string{
		"=MATCH(\"aaaa\",A1:A6,0)": "3",
		"=MATCH(\"*b\",A1:A5,0)":   "4",
		"=MATCH(\"?eee\",A1:A5,0)": "5",
		"=MATCH(\"?*?e\",A1:A5,0)": "5",
		"=MATCH(\"aaaa\",A1:A6,1)": "3",
		"=MATCH(10,B1:B6)":         "5",
		"=MATCH(8,C1:C6,1)":        "3",
		"=MATCH(6,B1:B6,-1)":       "1",
		"=MATCH(10,D1:D6,-1)":      "3",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "E1", formula))
		result, err := f.CalcCellValue("Sheet1", "E1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	calcError := map[string]string{
		"=MATCH(3,C1:C6,1)":  "#N/A",
		"=MATCH(5,C1:C6,-1)": "#N/A",
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "E1", formula))
		result, err := f.CalcCellValue("Sheet1", "E1")
		assert.EqualError(t, err, expected, formula)
		assert.Equal(t, "", result, formula)
	}
	assert.Equal(t, newErrorFormulaArg(formulaErrorNA, formulaErrorNA), calcMatch(2, nil, []formulaArg{}))
}

func TestCalcZTEST(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetSheetRow("Sheet1", "A1", &[]int{4, 5, 2, 5, 8, 9, 3, 2, 3, 8, 9, 5}))
	formulaList := map[string]string{
		"=Z.TEST(A1:L1,5)":   "0.371103278558538",
		"=Z.TEST(A1:L1,6)":   "0.838129187019751",
		"=Z.TEST(A1:L1,5,1)": "0.193238115385616",
		"=ZTEST(A1:L1,5)":    "0.371103278558538",
		"=ZTEST(A1:L1,6)":    "0.838129187019751",
		"=ZTEST(A1:L1,5,1)":  "0.193238115385616",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "M1", formula))
		result, err := f.CalcCellValue("Sheet1", "M1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
}

func TestStrToDate(t *testing.T) {
	_, _, _, _, err := strToDate("")
	assert.Equal(t, formulaErrorVALUE, err.Error)
}

func TestGetYearDays(t *testing.T) {
	for _, data := range [][]int{{2021, 0, 360}, {2000, 1, 366}, {2021, 1, 365}, {2000, 3, 365}} {
		assert.Equal(t, data[2], getYearDays(data[0], data[1]))
	}
}
