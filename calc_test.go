package excelize

import (
	"container/list"
	"math"
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
		"=2^3":                   "8",
		"=1=1":                   "TRUE",
		"=1=2":                   "FALSE",
		"=1<2":                   "TRUE",
		"=3<2":                   "FALSE",
		"=1<\"-1\"":              "TRUE",
		"=\"-1\"<1":              "FALSE",
		"=\"-1\"<\"-2\"":         "TRUE",
		"=2<=3":                  "TRUE",
		"=2<=1":                  "FALSE",
		"=1<=\"-1\"":             "TRUE",
		"=\"-1\"<=1":             "FALSE",
		"=\"-1\"<=\"-2\"":        "TRUE",
		"=2>1":                   "TRUE",
		"=2>3":                   "FALSE",
		"=1>\"-1\"":              "FALSE",
		"=\"-1\">-1":             "TRUE",
		"=\"-1\">\"-2\"":         "FALSE",
		"=2>=1":                  "TRUE",
		"=2>=3":                  "FALSE",
		"=1>=\"-1\"":             "FALSE",
		"=\"-1\">=-1":            "TRUE",
		"=\"-1\">=\"-2\"":        "FALSE",
		"=-----1+1":              "0",
		"=------1+1":             "2",
		"=---1---1":              "-2",
		"=---1----1":             "0",
		"=1&2":                   "12",
		"=15%":                   "0.15",
		"=1+20%":                 "1.2",
		"={1}+2":                 "3",
		"=1+{2}":                 "3",
		"={1}+{2}":               "3",
		"=A1+(B1-C1)":            "5",
		"=A1+(C1-B1)":            "-3",
		"=A1&B1&C1":              "14",
		"=B1+C1":                 "4",
		"=C1+B1":                 "4",
		"=C1+C1":                 "0",
		"=\"A\"=\"A\"":           "TRUE",
		"=\"A\"<>\"A\"":          "FALSE",
		"=TRUE()&FALSE()":        "TRUEFALSE",
		"=TRUE()&FALSE()<>FALSE": "TRUE",
		"=TRUE()&\"1\"":          "TRUE1",
		"=TRUE<>FALSE()":         "TRUE",
		"=TRUE<>1&\"x\"":         "TRUE",
		// Engineering Functions
		// BESSELI
		"=BESSELI(4.5,1)":    "15.3892227537359",
		"=BESSELI(32,1)":     "5502845511211.25",
		"=BESSELI({32},1)":   "5502845511211.25",
		"=BESSELI(32,{1})":   "5502845511211.25",
		"=BESSELI({32},{1})": "5502845511211.25",
		// BESSELJ
		"=BESSELJ(1.9,2)":     "0.329925727692387",
		"=BESSELJ({1.9},2)":   "0.329925727692387",
		"=BESSELJ(1.9,{2})":   "0.329925727692387",
		"=BESSELJ({1.9},{2})": "0.329925727692387",
		// BESSELK
		"=BESSELK(0.05,0)":  "3.11423403428966",
		"=BESSELK(0.05,1)":  "19.9096743272486",
		"=BESSELK(0.05,2)":  "799.501207124235",
		"=BESSELK(3,2)":     "0.0615104585619118",
		"=BESSELK({3},2)":   "0.0615104585619118",
		"=BESSELK(3,{2})":   "0.0615104585619118",
		"=BESSELK({3},{2})": "0.0615104585619118",
		// BESSELY
		"=BESSELY(0.05,0)":  "-1.97931100684153",
		"=BESSELY(0.05,1)":  "-12.789855163794",
		"=BESSELY(0.05,2)":  "-509.61489554492",
		"=BESSELY(9,2)":     "-0.229082087487741",
		"=BESSELY({9},2)":   "-0.229082087487741",
		"=BESSELY(9,{2})":   "-0.229082087487741",
		"=BESSELY({9},{2})": "-0.229082087487741",
		// BIN2DEC
		"=BIN2DEC(\"10\")":         "2",
		"=BIN2DEC(\"11\")":         "3",
		"=BIN2DEC(\"0000000010\")": "2",
		"=BIN2DEC(\"1111111110\")": "-2",
		"=BIN2DEC(\"110\")":        "6",
		"=BIN2DEC({\"110\"})":      "6",
		// BIN2HEX
		"=BIN2HEX(\"10\")":         "2",
		"=BIN2HEX(\"0000000001\")": "1",
		"=BIN2HEX(\"10\",10)":      "0000000002",
		"=BIN2HEX(\"1111111110\")": "FFFFFFFFFE",
		"=BIN2HEX(\"11101\")":      "1D",
		"=BIN2HEX({\"11101\"})":    "1D",
		// BIN2OCT
		"=BIN2OCT(\"101\")":        "5",
		"=BIN2OCT(\"0000000001\")": "1",
		"=BIN2OCT(\"10\",10)":      "0000000002",
		"=BIN2OCT(\"1111111110\")": "7777777776",
		"=BIN2OCT(\"1110\")":       "16",
		"=BIN2OCT({\"1110\"})":     "16",
		// BITAND
		"=BITAND(13,14)":     "12",
		"=BITAND({13},14)":   "12",
		"=BITAND(13,{14})":   "12",
		"=BITAND({13},{14})": "12",
		// BITLSHIFT
		"=BITLSHIFT(5,2)":     "20",
		"=BITLSHIFT({3},5)":   "96",
		"=BITLSHIFT(3,5)":     "96",
		"=BITLSHIFT(3,{5})":   "96",
		"=BITLSHIFT({3},{5})": "96",
		// BITOR
		"=BITOR(9,12)":     "13",
		"=BITOR({9},12)":   "13",
		"=BITOR(9,{12})":   "13",
		"=BITOR({9},{12})": "13",
		// BITRSHIFT
		"=BITRSHIFT(20,2)":     "5",
		"=BITRSHIFT(52,4)":     "3",
		"=BITRSHIFT({52},4)":   "3",
		"=BITRSHIFT(52,{4})":   "3",
		"=BITRSHIFT({52},{4})": "3",
		// BITXOR
		"=BITXOR(5,6)":      "3",
		"=BITXOR(9,12)":     "5",
		"=BITXOR({9},12)":   "5",
		"=BITXOR(9,{12})":   "5",
		"=BITXOR({9},{12})": "5",
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
		// CONVERT
		"=CONVERT(20.2,\"m\",\"yd\")":                    "22.0909886264217",
		"=CONVERT(20.2,\"cm\",\"yd\")":                   "0.220909886264217",
		"=CONVERT(0.2,\"gal\",\"tsp\")":                  "153.6",
		"=CONVERT(5,\"gal\",\"l\")":                      "18.92705892",
		"=CONVERT(0.02,\"Gm\",\"m\")":                    "20000000",
		"=CONVERT(0,\"C\",\"F\")":                        "32",
		"=CONVERT(1,\"ly^2\",\"ly^2\")":                  "1",
		"=CONVERT(0.00194255938572296,\"sg\",\"ozm\")":   "1",
		"=CONVERT(5,\"kg\",\"kg\")":                      "5",
		"=CONVERT(4.5359237E-01,\"kg\",\"lbm\")":         "1",
		"=CONVERT(0.2,\"kg\",\"hg\")":                    "2",
		"=CONVERT(12.345000000000001,\"km\",\"m\")":      "12345",
		"=CONVERT(12345,\"m\",\"km\")":                   "12.345",
		"=CONVERT(0.621371192237334,\"mi\",\"km\")":      "1",
		"=CONVERT(1.23450000000000E+05,\"ang\",\"um\")":  "12.345",
		"=CONVERT(1.23450000000000E+02,\"kang\",\"um\")": "12.345",
		"=CONVERT(1000,\"dal\",\"hl\")":                  "100",
		"=CONVERT(1,\"yd\",\"ft\")":                      "2.99999999999999",
		"=CONVERT(20,\"C\",\"F\")":                       "68",
		"=CONVERT(68,\"F\",\"C\")":                       "20",
		"=CONVERT(293.15,\"K\",\"F\")":                   "68",
		"=CONVERT(68,\"F\",\"K\")":                       "293.15",
		"=CONVERT(-273.15,\"C\",\"K\")":                  "0",
		"=CONVERT(-459.67,\"F\",\"K\")":                  "0",
		"=CONVERT(295.65,\"K\",\"C\")":                   "22.5",
		"=CONVERT(22.5,\"C\",\"K\")":                     "295.65",
		"=CONVERT(1667.85,\"C\",\"K\")":                  "1941",
		"=CONVERT(3034.13,\"F\",\"K\")":                  "1941",
		"=CONVERT(3493.8,\"Rank\",\"K\")":                "1941",
		"=CONVERT(1334.28,\"Reau\",\"K\")":               "1941",
		"=CONVERT(1941,\"K\",\"Rank\")":                  "3493.8",
		"=CONVERT(1941,\"K\",\"Reau\")":                  "1334.28",
		"=CONVERT(123.45,\"K\",\"kel\")":                 "123.45",
		"=CONVERT(123.45,\"C\",\"cel\")":                 "123.45",
		"=CONVERT(123.45,\"F\",\"fah\")":                 "123.45",
		"=CONVERT(16,\"bit\",\"byte\")":                  "2",
		"=CONVERT(1,\"kbyte\",\"byte\")":                 "1000",
		"=CONVERT(1,\"kibyte\",\"byte\")":                "1024",
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
		// DELTA
		"=DELTA(5,4)":       "0",
		"=DELTA(1.00001,1)": "0",
		"=DELTA(1.23,1.23)": "1",
		"=DELTA(1)":         "0",
		"=DELTA(0)":         "1",
		// ERF
		"=ERF(1.5)":   "0.966105146475311",
		"=ERF(0,1.5)": "0.966105146475311",
		"=ERF(1,2)":   "0.152621472069238",
		// ERF.PRECISE
		"=ERF.PRECISE(-1)":  "-0.842700792949715",
		"=ERF.PRECISE(1.5)": "0.966105146475311",
		// ERFC
		"=ERFC(0)":   "1",
		"=ERFC(0.5)": "0.479500122186953",
		"=ERFC(-1)":  "1.84270079294971",
		// ERFC.PRECISE
		"=ERFC.PRECISE(0)":   "1",
		"=ERFC.PRECISE(0.5)": "0.479500122186953",
		"=ERFC.PRECISE(-1)":  "1.84270079294971",
		// GESTEP
		"=GESTEP(1.2,0.001)":  "1",
		"=GESTEP(0.05,0.05)":  "1",
		"=GESTEP(-0.00001,0)": "0",
		"=GESTEP(-0.00001)":   "0",
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
		"=HEX2OCT({\"1F3\"})":      "763",
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
		"=IMCOS(\"3+0.5i\")": "-1.11634124452615-0.0735369737112366i",
		// IMCOSH
		"=IMCOSH(0.5)":           "1.12762596520638",
		"=IMCOSH(\"3+0.5i\")":    "8.83520460650099+4.80282508274303i",
		"=IMCOSH(\"2-i\")":       "2.03272300701967-3.0518977991518i",
		"=IMCOSH(COMPLEX(1,-1))": "0.833730025131149-0.988897705762865i",
		// IMCOT
		"=IMCOT(0.5)":           "1.83048772171245",
		"=IMCOT(\"3+0.5i\")":    "-0.479345578747373-2.01609252150623i",
		"=IMCOT(\"2-i\")":       "-0.171383612909185+0.821329797493852i",
		"=IMCOT(COMPLEX(1,-1))": "0.217621561854403+0.868014142895925i",
		// IMCSC
		"=IMCSC(\"j\")": "-0.850918128239322j",
		// IMCSCH
		"=IMCSCH(COMPLEX(1,-1))": "0.303931001628426+0.621518017170428i",
		// IMDIV
		"=IMDIV(\"5+2i\",\"1+i\")":          "3.5-1.5i",
		"=IMDIV(\"2+2i\",\"2+i\")":          "1.2+0.4i",
		"=IMDIV(COMPLEX(5,2),COMPLEX(0,1))": "2-5i",
		// IMEXP
		"=IMEXP(0)":             "1",
		"=IMEXP(0.5)":           "1.64872127070013",
		"=IMEXP(\"1-2i\")":      "-1.13120438375681-2.47172667200482i",
		"=IMEXP(COMPLEX(1,-1))": "1.46869393991589-2.28735528717884i",
		// IMLN
		"=IMLN(0.5)":           "-0.693147180559945",
		"=IMLN(\"3+0.5i\")":    "1.11231177576217+0.165148677414627i",
		"=IMLN(\"2-i\")":       "0.80471895621705-0.463647609000806i",
		"=IMLN(COMPLEX(1,-1))": "0.346573590279973-0.785398163397448i",
		// IMLOG10
		"=IMLOG10(0.5)":           "-0.301029995663981",
		"=IMLOG10(\"3+0.5i\")":    "0.483070866369516+0.0717231592947926i",
		"=IMLOG10(\"2-i\")":       "0.349485002168009-0.201359598136687i",
		"=IMLOG10(COMPLEX(1,-1))": "0.150514997831991-0.34109408846046i",
		// IMREAL
		"=IMREAL(\"5+2i\")":     "5",
		"=IMREAL(\"2+2i\")":     "2",
		"=IMREAL(6)":            "6",
		"=IMREAL(\"3i\")":       "0",
		"=IMREAL(COMPLEX(4,1))": "4",
		// IMSEC
		"=IMSEC(0.5)":           "1.13949392732455",
		"=IMSEC(\"3+0.5i\")":    "-0.89191317974033+0.0587531781817398i",
		"=IMSEC(\"2-i\")":       "-0.41314934426694-0.687527438655479i",
		"=IMSEC(COMPLEX(1,-1))": "0.498337030555187-0.591083841721045i",
		// IMSECH
		"=IMSECH(0.5)":           "0.886818883970074",
		"=IMSECH(\"3+0.5i\")":    "0.0873665779621303-0.0474925494901607i",
		"=IMSECH(\"2-i\")":       "0.151176298265577+0.226973675393722i",
		"=IMSECH(COMPLEX(1,-1))": "0.498337030555187+0.591083841721045i",
		// IMSIN
		"=IMSIN(0.5)":           "0.479425538604203",
		"=IMSIN(\"3+0.5i\")":    "0.15913058529844-0.515880442452527i",
		"=IMSIN(\"2-i\")":       "1.40311925062204+0.489056259041294i",
		"=IMSIN(COMPLEX(1,-1))": "1.29845758141598-0.634963914784736i",
		// IMSINH
		"=IMSINH(-0)":            "0",
		"=IMSINH(0.5)":           "0.521095305493747",
		"=IMSINH(\"3+0.5i\")":    "8.79151234349371+4.82669427481082i",
		"=IMSINH(\"2-i\")":       "1.95960104142161-3.16577851321617i",
		"=IMSINH(COMPLEX(1,-1))": "0.634963914784736-1.29845758141598i",
		// IMSQRT
		"=IMSQRT(\"i\")":     "0.707106781186548+0.707106781186548i",
		"=IMSQRT(\"2-i\")":   "1.45534669022535-0.343560749722512i",
		"=IMSQRT(\"5+2i\")":  "2.27872385417085+0.438842116902254i",
		"=IMSQRT(6)":         "2.44948974278318",
		"=IMSQRT(\"-2-4i\")": "1.11178594050284-1.79890743994787i",
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
		"=IMTAN(\"3+0.5i\")":    "-0.111621050771583+0.469469993425885i",
		"=IMTAN(\"2-i\")":       "-0.243458201185725-1.16673625724092i",
		"=IMTAN(COMPLEX(1,-1))": "0.271752585319512-1.08392332733869i",
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
		"=ACOS(-1)":     "3.14159265358979",
		"=ACOS(0)":      "1.5707963267949",
		"=ACOS(ABS(0))": "1.5707963267949",
		// ACOSH
		"=ACOSH(1)":        "0",
		"=ACOSH(2.5)":      "1.56679923697241",
		"=ACOSH(5)":        "2.29243166956118",
		"=ACOSH(ACOSH(5))": "1.47138332153668",
		// _xlfn.ACOT
		"=_xlfn.ACOT(1)":             "0.785398163397448",
		"=_xlfn.ACOT(-2)":            "2.67794504458899",
		"=_xlfn.ACOT(0)":             "1.5707963267949",
		"=_xlfn.ACOT(_xlfn.ACOT(0))": "0.566911504941009",
		// _xlfn.ACOTH
		"=_xlfn.ACOTH(-5)":      "-0.202732554054082",
		"=_xlfn.ACOTH(1.1)":     "1.52226121886171",
		"=_xlfn.ACOTH(2)":       "0.549306144334055",
		"=_xlfn.ACOTH(ABS(-2))": "0.549306144334055",
		// _xlfn.AGGREGATE
		"=_xlfn.AGGREGATE(1,0,A1:A6)":    "1.5",
		"=_xlfn.AGGREGATE(2,0,A1:A6)":    "4",
		"=_xlfn.AGGREGATE(3,0,A1:A6)":    "4",
		"=_xlfn.AGGREGATE(4,0,A1:A6)":    "3",
		"=_xlfn.AGGREGATE(5,0,A1:A6)":    "0",
		"=_xlfn.AGGREGATE(6,0,A1:A6)":    "0",
		"=_xlfn.AGGREGATE(7,0,A1:A6)":    "1.29099444873581",
		"=_xlfn.AGGREGATE(8,0,A1:A6)":    "1.11803398874989",
		"=_xlfn.AGGREGATE(9,0,A1:A6)":    "6",
		"=_xlfn.AGGREGATE(10,0,A1:A6)":   "1.66666666666667",
		"=_xlfn.AGGREGATE(11,0,A1:A6)":   "1.25",
		"=_xlfn.AGGREGATE(12,0,A1:A6)":   "1.5",
		"=_xlfn.AGGREGATE(14,0,A1:A6,1)": "3",
		"=_xlfn.AGGREGATE(15,0,A1:A6,1)": "0",
		"=_xlfn.AGGREGATE(16,0,A1:A6,1)": "3",
		"=_xlfn.AGGREGATE(17,0,A1:A6,1)": "0.75",
		"=_xlfn.AGGREGATE(19,0,A1:A6,1)": "0.25",
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
		"=-COS(0)":          "-1",
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
		"=_xlfn.DECIMAL(\"1100\",2)":    "12",
		"=_xlfn.DECIMAL(\"186A0\",16)":  "100000",
		"=_xlfn.DECIMAL(\"31L0\",32)":   "100000",
		"=_xlfn.DECIMAL(\"70122\",8)":   "28754",
		"=_xlfn.DECIMAL(\"0x70122\",8)": "28754",
		// DEGREES
		"=DEGREES(1)":          "57.2957795130823",
		"=DEGREES(2.5)":        "143.239448782706",
		"=DEGREES(DEGREES(1))": "3282.80635001174",
		// EVEN
		"=EVEN(23)":   "24",
		"=EVEN(2.22)": "4",
		"=EVEN(0)":    "0",
		"=EVEN(-0.3)": "-2",
		"=EVEN(-11)":  "-12",
		"=EVEN(-4)":   "-4",
		"=EVEN((0))":  "0",
		// EXP
		"=EXP(100)":    "2.68811714181614E+43",
		"=EXP(0.1)":    "1.10517091807565",
		"=EXP(0)":      "1",
		"=EXP(-5)":     "0.00673794699908547",
		"=EXP(EXP(0))": "2.71828182845905",
		// FACT
		"=FACT(3)":       "6",
		"=FACT(6)":       "720",
		"=FACT(10)":      "3628800",
		"=FACT(FACT(3))": "720",
		// FACTDOUBLE
		"=FACTDOUBLE(5)":             "15",
		"=FACTDOUBLE(8)":             "384",
		"=FACTDOUBLE(13)":            "135135",
		"=FACTDOUBLE(FACTDOUBLE(1))": "1",
		// FLOOR
		"=FLOOR(26.75,0.1)":        "26.7",
		"=FLOOR(26.75,0.5)":        "26.5",
		"=FLOOR(26.75,1)":          "26",
		"=FLOOR(26.75,10)":         "20",
		"=FLOOR(26.75,20)":         "20",
		"=FLOOR(-26.75,-0.1)":      "-26.7",
		"=FLOOR(-26.75,-1)":        "-26",
		"=FLOOR(-26.75,-5)":        "-25",
		"=FLOOR(-2.05,2)":          "-4",
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
		"=_xlfn.FLOOR.PRECISE(26.75,0.1)":                     "26.7",
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
		"=GCD(\"0\",1)":  "1",
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
		"=ISO.CEILING(-22.25,0.1)":         "-22.2",
		"=ISO.CEILING(-22.25,5)":           "-20",
		"=ISO.CEILING(-22.25,0)":           "0",
		"=ISO.CEILING(1,ISO.CEILING(1,0))": "0",
		// LCM
		"=LCM(1,5)":        "5",
		"=LCM(15,10,25)":   "150",
		"=LCM(1,8,12)":     "24",
		"=LCM(7,2)":        "14",
		"=LCM(7)":          "7",
		"=LCM(\"\",1)":     "1",
		"=LCM(0,0)":        "0",
		"=LCM(0,LCM(0,0))": "0",
		// LN
		"=LN(1)":       "0",
		"=LN(100)":     "4.60517018598809",
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
		"=IMLOG2(\"5+2i\")": "2.42899049756379+0.548954663286635i",
		"=IMLOG2(\"2-i\")":  "1.16096404744368-0.668902106225488i",
		"=IMLOG2(6)":        "2.58496250072116",
		"=IMLOG2(\"3i\")":   "1.58496250072116+2.2661800709136i",
		"=IMLOG2(\"4+i\")":  "2.04373142062517+0.353429502416735i",
		// IMPOWER
		"=IMPOWER(\"2-i\",2)":   "3-4i",
		"=IMPOWER(\"2-i\",3)":   "2-11i",
		"=IMPOWER(9,0.5)":       "3",
		"=IMPOWER(\"2+4i\",-2)": "-0.03-0.04i",
		// IMPRODUCT
		"=IMPRODUCT(3,6)":                       "18",
		"=IMPRODUCT(\"\",3,SUM(6))":             "18",
		"=IMPRODUCT(\"1-i\",\"5+10i\",2)":       "30+10i",
		"=IMPRODUCT(COMPLEX(5,2),COMPLEX(0,1))": "-2+5i",
		"=IMPRODUCT(A1:C1)":                     "4",
		// MINVERSE
		"=MINVERSE(A1:B2)": "-0",
		// MMULT
		"=MMULT(0,0)":         "0",
		"=MMULT(2,4)":         "8",
		"=MMULT(A4:A4,A4:A4)": "0",
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
		"=MULTINOMIAL(\"\",3,1,2,5)":   "27720",
		"=MULTINOMIAL(MULTINOMIAL(1))": "1",
		// _xlfn.MUNIT
		"=_xlfn.MUNIT(4)": "1",
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
		"=PI()": "3.14159265358979",
		// POWER
		"=POWER(4,2)":          "16",
		"=POWER(4,POWER(1,1))": "4",
		// PRODUCT
		"=PRODUCT(3,6)":            "18",
		"=PRODUCT(\"3\",\"6\")":    "18",
		"=PRODUCT(PRODUCT(1),3,6)": "18",
		"=PRODUCT(C1:C2)":          "1",
		// QUOTIENT
		"=QUOTIENT(5,2)":             "2",
		"=QUOTIENT(4.5,3.1)":         "1",
		"=QUOTIENT(-10,3)":           "-3",
		"=QUOTIENT(QUOTIENT(1,2),3)": "0",
		// RADIANS
		"=RADIANS(50)":           "0.872664625997165",
		"=RADIANS(-180)":         "-3.14159265358979",
		"=RADIANS(180)":          "3.14159265358979",
		"=RADIANS(360)":          "6.28318530717959",
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
		"=ROUND(100.319,1)":       "100.3",
		"=ROUND(5.28,1)":          "5.3",
		"=ROUND(5.9999,3)":        "6",
		"=ROUND(99.5,0)":          "100",
		"=ROUND(-6.3,0)":          "-6",
		"=ROUND(-100.5,0)":        "-101",
		"=ROUND(-22.45,1)":        "-22.5",
		"=ROUND(999,-1)":          "1000",
		"=ROUND(991,-1)":          "990",
		"=ROUND(ROUND(100,1),-1)": "100",
		// ROUNDDOWN
		"=ROUNDDOWN(99.999,1)":            "99.9",
		"=ROUNDDOWN(99.999,2)":            "99.99",
		"=ROUNDDOWN(99.999,0)":            "99",
		"=ROUNDDOWN(99.999,-1)":           "90",
		"=ROUNDDOWN(-99.999,2)":           "-99.99",
		"=ROUNDDOWN(-99.999,-1)":          "-90",
		"=ROUNDDOWN(ROUNDDOWN(100,1),-1)": "100",
		// ROUNDUP
		"=ROUNDUP(11.111,1)":          "11.2",
		"=ROUNDUP(11.111,2)":          "11.12",
		"=ROUNDUP(11.111,0)":          "12",
		"=ROUNDUP(11.111,-1)":         "20",
		"=ROUNDUP(-11.111,2)":         "-11.12",
		"=ROUNDUP(-11.111,-1)":        "-20",
		"=ROUNDUP(ROUNDUP(100,1),-1)": "100",
		// SEARCH
		"=SEARCH(\"s\",F1)":           "1",
		"=SEARCH(\"s\",F1,2)":         "5",
		"=SEARCH(\"e\",F1)":           "4",
		"=SEARCH(\"e*\",F1)":          "4",
		"=SEARCH(\"?e\",F1)":          "3",
		"=SEARCH(\"??e\",F1)":         "2",
		"=SEARCH(6,F2)":               "2",
		"=SEARCH(\"?\",\"你好world\")":  "1",
		"=SEARCH(\"?l\",\"你好world\")": "5",
		"=SEARCH(\"?+\",\"你好 1+2\")":  "4",
		"=SEARCH(\" ?+\",\"你好 1+2\")": "3",
		// SEARCHB
		"=SEARCHB(\"s\",F1)":           "1",
		"=SEARCHB(\"s\",F1,2)":         "5",
		"=SEARCHB(\"e\",F1)":           "4",
		"=SEARCHB(\"e*\",F1)":          "4",
		"=SEARCHB(\"?e\",F1)":          "3",
		"=SEARCHB(\"??e\",F1)":         "2",
		"=SEARCHB(6,F2)":               "2",
		"=SEARCHB(\"?\",\"你好world\")":  "5",
		"=SEARCHB(\"?l\",\"你好world\")": "7",
		"=SEARCHB(\"?+\",\"你好 1+2\")":  "6",
		"=SEARCHB(\" ?+\",\"你好 1+2\")": "5",
		// SEC
		"=_xlfn.SEC(-3.14159265358979)": "-1",
		"=_xlfn.SEC(0)":                 "1",
		"=_xlfn.SEC(_xlfn.SEC(0))":      "0.54030230586814",
		// SECH
		"=_xlfn.SECH(-3.14159265358979)": "0.0862667383340547",
		"=_xlfn.SECH(0)":                 "1",
		"=_xlfn.SECH(_xlfn.SECH(0))":     "0.648054273663885",
		// SERIESSUM
		"=SERIESSUM(1,2,3,A1:A4)": "6",
		"=SERIESSUM(1,2,3,A1:B5)": "15",
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
		"=SINH(-2)":      "-3.62686040784702",
		"=SINH(SINH(0))": "0",
		// SQRT
		"=SQRT(4)":        "2",
		"=SQRT(SQRT(16))": "2",
		// SQRTPI
		"=SQRTPI(5)":         "3.96332729760601",
		"=SQRTPI(0.2)":       "0.792665459521202",
		"=SQRTPI(100)":       "17.7245385090552",
		"=SQRTPI(0)":         "0",
		"=SQRTPI(SQRTPI(0))": "0",
		// STDEV
		"=STDEV(F2:F9)":         "10724.9782875238",
		"=STDEV(MUNIT(2))":      "0.577350269189626",
		"=STDEV(0,INT(0))":      "0",
		"=STDEV(INT(1),INT(1))": "0",
		// STDEV.S
		"=STDEV.S(F2:F9)": "10724.9782875238",
		// STDEVA
		"=STDEVA(F2:F9)":    "10724.9782875238",
		"=STDEVA(MUNIT(2))": "0.577350269189626",
		"=STDEVA(0,INT(0))": "0",
		// POISSON.DIST
		"=POISSON.DIST(20,25,FALSE)": "0.0519174686084913",
		"=POISSON.DIST(35,40,TRUE)":  "0.242414197690103",
		// POISSON
		"=POISSON(20,25,FALSE)": "0.0519174686084913",
		"=POISSON(35,40,TRUE)":  "0.242414197690103",
		// SUBTOTAL
		"=SUBTOTAL(1,A1:A6)":         "1.5",
		"=SUBTOTAL(2,A1:A6)":         "4",
		"=SUBTOTAL(3,A1:A6)":         "4",
		"=SUBTOTAL(4,A1:A6)":         "3",
		"=SUBTOTAL(5,A1:A6)":         "0",
		"=SUBTOTAL(6,A1:A6)":         "0",
		"=SUBTOTAL(7,A1:A6)":         "1.29099444873581",
		"=SUBTOTAL(8,A1:A6)":         "1.11803398874989",
		"=SUBTOTAL(9,A1:A6)":         "6",
		"=SUBTOTAL(10,A1:A6)":        "1.66666666666667",
		"=SUBTOTAL(11,A1:A6)":        "1.25",
		"=SUBTOTAL(101,A1:A6)":       "1.5",
		"=SUBTOTAL(102,A1:A6)":       "4",
		"=SUBTOTAL(103,A1:A6)":       "4",
		"=SUBTOTAL(104,A1:A6)":       "3",
		"=SUBTOTAL(105,A1:A6)":       "0",
		"=SUBTOTAL(106,A1:A6)":       "0",
		"=SUBTOTAL(107,A1:A6)":       "1.29099444873581",
		"=SUBTOTAL(108,A1:A6)":       "1.11803398874989",
		"=SUBTOTAL(109,A1:A6)":       "6",
		"=SUBTOTAL(109,A1:A6,A1:A6)": "12",
		"=SUBTOTAL(110,A1:A6)":       "1.66666666666667",
		"=SUBTOTAL(111,A1:A6)":       "1.25",
		"=SUBTOTAL(111,A1:A6,A1:A6)": "1.25",
		// SUM
		"=SUM(1,2)":                           "3",
		"=SUM(\"1\",\"2\")":                   "3",
		"=SUM(\"\",1,2)":                      "3",
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
		"=1+SUM(SUM(1,2*3),4)*4/3+5+(4+2)*3":  "38.6666666666667",
		"=SUM(1+ROW())":                       "2",
		"=SUM((SUM(2))+1)":                    "3",
		"=IF(2<0, 1, (4))":                    "4",
		"=IF(2>0, (1), 4)":                    "1",
		"=IF(2>0, (A1)*2.5, 4)":               "2.5",
		"=SUM({1,2,3,4,\"\"})":                "10",
		// SUMIF
		"=SUMIF(F1:F5, \"\")":             "0",
		"=SUMIF(A1:A5, \"3\")":            "3",
		"=SUMIF(F1:F5, \"=36693\")":       "36693",
		"=SUMIF(F1:F5, \"<100\")":         "0",
		"=SUMIF(F1:F5, \"<=36693\")":      "93233",
		"=SUMIF(F1:F5, \">100\")":         "146554",
		"=SUMIF(F1:F5, \">=100\")":        "146554",
		"=SUMIF(F1:F5, \">=text\")":       "0",
		"=SUMIF(F1:F5, \"*Jan\",F2:F5)":   "0",
		"=SUMIF(D3:D7,\"Jan\",F2:F5)":     "112114",
		"=SUMIF(D2:D9,\"Feb\",F2:F9)":     "157559",
		"=SUMIF(E2:E9,\"North 1\",F2:F9)": "66582",
		"=SUMIF(E2:E9,\"North*\",F2:F9)":  "138772",
		"=SUMIF(D1:D3,\"Month\",D1:D3)":   "0",
		// SUMPRODUCT
		"=SUMPRODUCT(A1,B1)":             "4",
		"=SUMPRODUCT(A1:A2,B1:B2)":       "14",
		"=SUMPRODUCT(A1:A3,B1:B3)":       "14",
		"=SUMPRODUCT(A1:B3)":             "15",
		"=SUMPRODUCT(A1:A3,B1:B3,B2:B4)": "20",
		// SUMSQ
		"=SUMSQ(A1:A4)":              "14",
		"=SUMSQ(A1,B1,A2,B2,6)":      "82",
		"=SUMSQ(\"\",A1,B1,A2,B2,6)": "82",
		"=SUMSQ(1,SUMSQ(1))":         "2",
		"=SUMSQ(\"1\",SUMSQ(1))":     "2",
		"=SUMSQ(MUNIT(3))":           "3",
		// SUMX2MY2
		"=SUMX2MY2(A1:A4,B1:B4)": "-36",
		// SUMX2PY2
		"=SUMX2PY2(A1:A4,B1:B4)": "46",
		// SUMXMY2
		"=SUMXMY2(A1:A4,B1:B4)": "18",
		// TAN
		"=TAN(1.047197551)": "1.73205080678249",
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
		"=AVERAGEA(\"1\")":  "1",
		"=AVERAGEA(A1:A2)":  "1.5",
		"=AVERAGEA(D2:F9)":  "12671.375",
		// BETA.DIST
		"=BETA.DIST(0.4,4,5,TRUE,0,1)":  "0.4059136",
		"=BETA.DIST(0.6,4,5,FALSE,0,1)": "1.548288",
		// BETADIST
		"=BETADIST(0.4,4,5)":         "0.4059136",
		"=BETADIST(0.4,4,5,0,1)":     "0.4059136",
		"=BETADIST(0.4,4,5,0.4,1)":   "0",
		"=BETADIST(1,2,2,1,3)":       "0",
		"=BETADIST(0.4,4,5,0.2,0.4)": "1",
		"=BETADIST(0.4,4,1)":         "0.0256",
		"=BETADIST(0.4,1,5)":         "0.92224",
		"=BETADIST(3,4,6,2,4)":       "0.74609375",
		"=BETADIST(0.4,2,100)":       "1",
		"=BETADIST(0.75,3,4)":        "0.96240234375",
		"=BETADIST(0.2,0.7,4)":       "0.71794309318323",
		"=BETADIST(0.01,3,4)":        "1.955359E-05",
		"=BETADIST(0.75,130,140)":    "1",
		// BETAINV
		"=BETAINV(0.2,4,5,0,1)": "0.303225844664082",
		// BETA.INV
		"=BETA.INV(0.2,4,5,0,1)": "0.303225844664082",
		// BINOMDIST
		"=BINOMDIST(10,100,0.5,FALSE)": "1.36554263874631E-17",
		"=BINOMDIST(50,100,0.5,FALSE)": "0.0795892373871787",
		"=BINOMDIST(65,100,0.5,FALSE)": "0.000863855665741652",
		"=BINOMDIST(10,100,0.5,TRUE)":  "1.53164508771899E-17",
		"=BINOMDIST(50,100,0.5,TRUE)":  "0.539794618693589",
		"=BINOMDIST(65,100,0.5,TRUE)":  "0.999105034804256",
		// BINOM.DIST
		"=BINOM.DIST(10,100,0.5,FALSE)": "1.36554263874631E-17",
		"=BINOM.DIST(50,100,0.5,FALSE)": "0.0795892373871787",
		"=BINOM.DIST(65,100,0.5,FALSE)": "0.000863855665741652",
		"=BINOM.DIST(10,100,0.5,TRUE)":  "1.53164508771899E-17",
		"=BINOM.DIST(50,100,0.5,TRUE)":  "0.539794618693589",
		"=BINOM.DIST(65,100,0.5,TRUE)":  "0.999105034804256",
		// BINOM.DIST.RANGE
		"=BINOM.DIST.RANGE(100,0.5,0,40)":   "0.0284439668204904",
		"=BINOM.DIST.RANGE(100,0.5,45,55)":  "0.728746975926165",
		"=BINOM.DIST.RANGE(100,0.5,50,100)": "0.539794618693589",
		"=BINOM.DIST.RANGE(100,0.5,50)":     "0.0795892373871787",
		// BINOM.INV
		"=BINOM.INV(0,0.5,0.75)":   "0",
		"=BINOM.INV(0.1,0.1,0.75)": "0",
		"=BINOM.INV(0.6,0.4,0.75)": "0",
		"=BINOM.INV(2,0.4,0.75)":   "1",
		"=BINOM.INV(100,0.5,20%)":  "46",
		"=BINOM.INV(100,0.5,50%)":  "50",
		"=BINOM.INV(100,0.5,90%)":  "56",
		// CHIDIST
		"=CHIDIST(0.5,3)": "0.918891411654676",
		"=CHIDIST(8,3)":   "0.0460117056892314",
		"=CHIDIST(40,4)":  "4.32842260712097E-08",
		"=CHIDIST(42,4)":  "1.66816329414062E-08",
		// CHIINV
		"=CHIINV(0.5,1)":  "0.454936423119572",
		"=CHIINV(0.75,1)": "0.101531044267622",
		"=CHIINV(0.1,2)":  "4.60517018598809",
		"=CHIINV(0.8,2)":  "0.446287102628419",
		// CHISQ.DIST
		"=CHISQ.DIST(0,2,TRUE)":        "0",
		"=CHISQ.DIST(4,1,TRUE)":        "0.954499736103642",
		"=CHISQ.DIST(1180,1180,FALSE)": "0.00821093706387967",
		"=CHISQ.DIST(2,1,FALSE)":       "0.103776874355149",
		"=CHISQ.DIST(3,2,FALSE)":       "0.111565080074215",
		"=CHISQ.DIST(2,3,FALSE)":       "0.207553748710297",
		"=CHISQ.DIST(1425,1,FALSE)":    "3.88315098887099E-312",
		"=CHISQ.DIST(3,2,TRUE)":        "0.77686983985157",
		// CHISQ.DIST.RT
		"=CHISQ.DIST.RT(0.5,3)": "0.918891411654676",
		"=CHISQ.DIST.RT(8,3)":   "0.0460117056892314",
		"=CHISQ.DIST.RT(40,4)":  "4.32842260712097E-08",
		"=CHISQ.DIST.RT(42,4)":  "1.66816329414062E-08",
		// CHISQ.INV
		"=CHISQ.INV(0,2)":    "0",
		"=CHISQ.INV(0.75,1)": "1.32330369693147",
		"=CHISQ.INV(0.1,2)":  "0.210721031315653",
		"=CHISQ.INV(0.8,2)":  "3.2188758248682",
		"=CHISQ.INV(0.25,3)": "1.21253290304567",
		// CHISQ.INV.RT
		"=CHISQ.INV.RT(0.75,1)": "0.101531044267622",
		"=CHISQ.INV.RT(0.1,2)":  "4.60517018598809",
		"=CHISQ.INV.RT(0.8,2)":  "0.446287102628419",
		// CONFIDENCE
		"=CONFIDENCE(0.05,0.07,100)": "0.0137197479028414",
		// CONFIDENCE.NORM
		"=CONFIDENCE.NORM(0.05,0.07,100)": "0.0137197479028414",
		// CONFIDENCE.T
		"=CONFIDENCE.T(0.05,0.07,100)": "0.0138895186611049",
		// CORREL
		"=CORREL(A1:A5,B1:B5)": "1",
		// COUNT
		"=COUNT()":                              "0",
		"=COUNT(E1:F2,\"text\",1,INT(2),\"0\")": "4",
		// COUNTA
		"=COUNTA()":                              "0",
		"=COUNTA(A1:A5,B2:B5,\"text\",1,INT(2))": "8",
		"=COUNTA(COUNTA(1),MUNIT(1))":            "2",
		"=COUNTA(D1:D2)":                         "2",
		// COUNTBLANK
		"=COUNTBLANK(MUNIT(1))": "0",
		"=COUNTBLANK(1)":        "0",
		"=COUNTBLANK(B1:C1)":    "1",
		"=COUNTBLANK(C1)":       "0",
		// COUNTIF
		"=COUNTIF(D1:D9,\"Jan\")":     "4",
		"=COUNTIF(D1:D9,\"<>Jan\")":   "5",
		"=COUNTIF(A1:F9,\">=50000\")": "2",
		"=COUNTIF(A1:F9,TRUE)":        "0",
		// COUNTIFS
		"=COUNTIFS(A1:A9,2,D1:D9,\"Jan\")":          "1",
		"=COUNTIFS(F1:F9,\">20000\",D1:D9,\"Jan\")": "4",
		"=COUNTIFS(F1:F9,\">60000\",D1:D9,\"Jan\")": "0",
		// CRITBINOM
		"=CRITBINOM(0,0.5,0.75)":   "0",
		"=CRITBINOM(0.1,0.1,0.75)": "0",
		"=CRITBINOM(0.6,0.4,0.75)": "0",
		"=CRITBINOM(2,0.4,0.75)":   "1",
		"=CRITBINOM(100,0.5,20%)":  "46",
		"=CRITBINOM(100,0.5,50%)":  "50",
		"=CRITBINOM(100,0.5,90%)":  "56",
		// DEVSQ
		"=DEVSQ(1,3,5,2,9,7)": "47.5",
		"=DEVSQ(A1:D2)":       "10",
		// FISHER
		"=FISHER(-0.9)":    "-1.47221948958322",
		"=FISHER(-0.25)":   "-0.255412811882995",
		"=FISHER(0.8)":     "1.09861228866811",
		"=FISHER(\"0.8\")": "1.09861228866811",
		"=FISHER(INT(0))":  "0",
		// FISHERINV
		"=FISHERINV(-0.2)":   "-0.197375320224904",
		"=FISHERINV(INT(0))": "0",
		"=FISHERINV(\"0\")":  "0",
		"=FISHERINV(2.8)":    "0.992631520201128",
		// FORECAST
		"=FORECAST(7,A1:A7,B1:B7)": "4",
		// FORECAST.LINEAR
		"=FORECAST.LINEAR(7,A1:A7,B1:B7)": "4",
		// FREQUENCY
		"=SUM(FREQUENCY(A2,B2))":       "1",
		"=SUM(FREQUENCY(A1:A5,B1:B2))": "4",
		// GAMMA
		"=GAMMA(0.1)":     "9.51350769866873",
		"=GAMMA(INT(1))":  "1",
		"=GAMMA(1.5)":     "0.886226925452758",
		"=GAMMA(5.5)":     "52.3427777845535",
		"=GAMMA(\"5.5\")": "52.3427777845535",
		// GAMMA.DIST
		"=GAMMA.DIST(6,3,2,FALSE)": "0.112020903827694",
		"=GAMMA.DIST(6,3,2,TRUE)":  "0.576809918873156",
		// GAMMADIST
		"=GAMMADIST(6,3,2,FALSE)": "0.112020903827694",
		"=GAMMADIST(6,3,2,TRUE)":  "0.576809918873156",
		// GAMMA.INV
		"=GAMMA.INV(0.5,3,2)":   "5.34812062744712",
		"=GAMMA.INV(0.5,0.5,1)": "0.227468211559786",
		// GAMMAINV
		"=GAMMAINV(0.5,3,2)":   "5.34812062744712",
		"=GAMMAINV(0.5,0.5,1)": "0.227468211559786",
		// GAMMALN
		"=GAMMALN(4.5)":    "2.45373657084244",
		"=GAMMALN(INT(1))": "0",
		// GAMMALN.PRECISE
		"=GAMMALN.PRECISE(0.4)": "0.796677817701784",
		"=GAMMALN.PRECISE(4.5)": "2.45373657084244",
		// GAUSS
		"=GAUSS(-5)":    "-0.499999713348428",
		"=GAUSS(0)":     "0",
		"=GAUSS(\"0\")": "0",
		"=GAUSS(0.1)":   "0.039827837277029",
		"=GAUSS(2.5)":   "0.493790334674224",
		// GEOMEAN
		"=GEOMEAN(2.5,3,0.5,1,3)": "1.6226711115996",
		// HARMEAN
		"=HARMEAN(2.5,3,0.5,1,3)":               "1.22950819672131",
		"=HARMEAN(\"2.5\",3,0.5,1,INT(3),\"\")": "1.22950819672131",
		// HYPGEOM.DIST
		"=HYPGEOM.DIST(0,3,3,9,TRUE)":   "0.238095238095238",
		"=HYPGEOM.DIST(1,3,3,9,TRUE)":   "0.773809523809524",
		"=HYPGEOM.DIST(2,3,3,9,TRUE)":   "0.988095238095238",
		"=HYPGEOM.DIST(3,3,3,9,TRUE)":   "1",
		"=HYPGEOM.DIST(1,4,4,12,FALSE)": "0.452525252525253",
		"=HYPGEOM.DIST(2,4,4,12,FALSE)": "0.339393939393939",
		"=HYPGEOM.DIST(3,4,4,12,FALSE)": "0.0646464646464646",
		"=HYPGEOM.DIST(4,4,4,12,FALSE)": "0.00202020202020202",
		// HYPGEOMDIST
		"=HYPGEOMDIST(1,4,4,12)": "0.452525252525253",
		"=HYPGEOMDIST(2,4,4,12)": "0.339393939393939",
		"=HYPGEOMDIST(3,4,4,12)": "0.0646464646464646",
		"=HYPGEOMDIST(4,4,4,12)": "0.00202020202020202",
		// INTERCEPT
		"=INTERCEPT(A1:A4,B1:B4)": "-3",
		// KURT
		"=KURT(F1:F9)":           "-1.03350350255137",
		"=KURT(F1,F2:F9)":        "-1.03350350255137",
		"=KURT(INT(1),MUNIT(2))": "-3.33333333333334",
		// EXPON.DIST
		"=EXPON.DIST(0.5,1,TRUE)":  "0.393469340287367",
		"=EXPON.DIST(0.5,1,FALSE)": "0.606530659712633",
		"=EXPON.DIST(2,1,TRUE)":    "0.864664716763387",
		// EXPONDIST
		"=EXPONDIST(0.5,1,TRUE)":  "0.393469340287367",
		"=EXPONDIST(0.5,1,FALSE)": "0.606530659712633",
		"=EXPONDIST(2,1,TRUE)":    "0.864664716763387",
		// FDIST
		"=FDIST(5,1,2)": "0.154845745271483",
		// F.DIST
		"=F.DIST(1,2,5,TRUE)":  "0.568798849628308",
		"=F.DIST(1,2,5,FALSE)": "0.308000821694066",
		// F.DIST.RT
		"=F.DIST.RT(5,1,2)": "0.154845745271483",
		// F.INV
		"=F.INV(0.9,2,5)": "3.77971607877395",
		// FINV
		"=FINV(0.2,1,2)":   "3.55555555555555",
		"=FINV(0.6,1,2)":   "0.380952380952381",
		"=FINV(0.6,2,2)":   "0.666666666666667",
		"=FINV(0.6,4,4)":   "0.763454070045235",
		"=FINV(0.5,4,8)":   "0.914645355977072",
		"=FINV(0.1,79,86)": "1.32646097270444",
		"=FINV(1,40,5)":    "0",
		// F.INV.RT
		"=F.INV.RT(0.2,1,2)":   "3.55555555555555",
		"=F.INV.RT(0.6,1,2)":   "0.380952380952381",
		"=F.INV.RT(0.6,2,2)":   "0.666666666666667",
		"=F.INV.RT(0.6,4,4)":   "0.763454070045235",
		"=F.INV.RT(0.5,4,8)":   "0.914645355977072",
		"=F.INV.RT(0.1,79,86)": "1.32646097270444",
		"=F.INV.RT(1,40,5)":    "0",
		// LOGINV
		"=LOGINV(0.3,2,0.2)": "6.6533460753367",
		// LOGINV
		"=LOGNORM.INV(0.3,2,0.2)": "6.6533460753367",
		// LOGNORM.DIST
		"=LOGNORM.DIST(0.5,10,5,FALSE)": "0.0162104821842127",
		"=LOGNORM.DIST(12,10,5,TRUE)":   "0.0664171147992078",
		// LOGNORMDIST
		"=LOGNORMDIST(12,10,5)": "0.0664171147992078",
		// NEGBINOM.DIST
		"=NEGBINOM.DIST(6,12,0.5,FALSE)":  "0.047210693359375",
		"=NEGBINOM.DIST(12,12,0.5,FALSE)": "0.0805901288986206",
		"=NEGBINOM.DIST(15,12,0.5,FALSE)": "0.057564377784729",
		"=NEGBINOM.DIST(12,12,0.5,TRUE)":  "0.580590128898621",
		"=NEGBINOM.DIST(15,12,0.5,TRUE)":  "0.778965830802917",
		// NEGBINOMDIST
		"=NEGBINOMDIST(6,12,0.5)":  "0.047210693359375",
		"=NEGBINOMDIST(12,12,0.5)": "0.0805901288986206",
		"=NEGBINOMDIST(15,12,0.5)": "0.057564377784729",
		// NORM.DIST
		"=NORM.DIST(0.8,1,0.3,TRUE)": "0.252492537546923",
		"=NORM.DIST(50,40,20,FALSE)": "0.017603266338215",
		// NORMDIST
		"=NORMDIST(0.8,1,0.3,TRUE)": "0.252492537546923",
		"=NORMDIST(50,40,20,FALSE)": "0.017603266338215",
		// NORM.INV
		"=NORM.INV(0.6,5,2)": "5.50669420572",
		// NORMINV
		"=NORMINV(0.6,5,2)":     "5.50669420572",
		"=NORMINV(0.99,40,1.5)": "43.489521811582",
		"=NORMINV(0.02,40,1.5)": "36.9193766364954",
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
		"=MAX(1)":           "1",
		"=MAX(TRUE())":      "1",
		"=MAX(0.5,TRUE())":  "1",
		"=MAX(FALSE())":     "0",
		"=MAX(MUNIT(2))":    "1",
		"=MAX(INT(1))":      "1",
		"=MAX(\"0\",\"2\")": "2",
		// MAXA
		"=MAXA(1)":          "1",
		"=MAXA(TRUE())":     "1",
		"=MAXA(0.5,TRUE())": "1",
		"=MAXA(FALSE())":    "0",
		"=MAXA(MUNIT(2))":   "1",
		"=MAXA(INT(1))":     "1",
		"=MAXA(A1:B4,MUNIT(1),INT(0),1,E1:F2,\"\")": "36693",
		// MAXIFS
		"=MAXIFS(F2:F4,A2:A4,\">0\")": "36693",
		// MEDIAN
		"=MEDIAN(A1:A5,12)":               "2",
		"=MEDIAN(A1:A5)":                  "1.5",
		"=MEDIAN(A1:A5,MEDIAN(A1:A5,12))": "2",
		"=MEDIAN(\"0\",\"2\")":            "1",
		// MIN
		"=MIN(1)":           "1",
		"=MIN(TRUE())":      "1",
		"=MIN(0.5,FALSE())": "0",
		"=MIN(FALSE())":     "0",
		"=MIN(MUNIT(2))":    "0",
		"=MIN(INT(1))":      "1",
		"=MIN(2,\"1\")":     "1",
		// MINA
		"=MINA(1)":           "1",
		"=MINA(TRUE())":      "1",
		"=MINA(0.5,FALSE())": "0",
		"=MINA(FALSE())":     "0",
		"=MINA(MUNIT(2))":    "0",
		"=MINA(INT(1))":      "1",
		"=MINA(A1:B4,MUNIT(1),INT(0),1,E1:F2,\"\")": "0",
		// MINIFS
		"=MINIFS(F2:F4,A2:A4,\">0\")": "22100",
		// PEARSON
		"=PEARSON(A1:A4,B1:B4)": "1",
		// PERCENTILE.EXC
		"=PERCENTILE.EXC(A1:A4,0.2)": "0",
		"=PERCENTILE.EXC(A1:A4,0.6)": "2",
		// PERCENTILE.INC
		"=PERCENTILE.INC(A1:A4,0.2)": "0.6",
		// PERCENTILE
		"=PERCENTILE(A1:A4,0.2)": "0.6",
		"=PERCENTILE(0,0)":       "0",
		// PERCENTRANK.EXC
		"=PERCENTRANK.EXC(A1:B4,0)":     "0.142",
		"=PERCENTRANK.EXC(A1:B4,2)":     "0.428",
		"=PERCENTRANK.EXC(A1:B4,2.5)":   "0.5",
		"=PERCENTRANK.EXC(A1:B4,2.6,1)": "0.5",
		"=PERCENTRANK.EXC(A1:B4,5)":     "0.857",
		// PERCENTRANK.INC
		"=PERCENTRANK.INC(A1:B4,0)":     "0",
		"=PERCENTRANK.INC(A1:B4,2)":     "0.4",
		"=PERCENTRANK.INC(A1:B4,2.5)":   "0.5",
		"=PERCENTRANK.INC(A1:B4,2.6,1)": "0.5",
		"=PERCENTRANK.INC(A1:B4,5)":     "1",
		// PERCENTRANK
		"=PERCENTRANK(A1:B4,0)":     "0",
		"=PERCENTRANK(A1:B4,2)":     "0.4",
		"=PERCENTRANK(A1:B4,2.5)":   "0.5",
		"=PERCENTRANK(A1:B4,2.6,1)": "0.5",
		"=PERCENTRANK(A1:B4,5)":     "1",
		// PERMUT
		"=PERMUT(6,6)":  "720",
		"=PERMUT(7,6)":  "5040",
		"=PERMUT(10,6)": "151200",
		// PERMUTATIONA
		"=PERMUTATIONA(6,6)": "46656",
		"=PERMUTATIONA(7,6)": "117649",
		// PHI
		"=PHI(-1.5)": "0.129517595665892",
		"=PHI(0)":    "0.398942280401433",
		"=PHI(0.1)":  "0.396952547477012",
		"=PHI(1)":    "0.241970724519143",
		// QUARTILE
		"=QUARTILE(A1:A4,2)": "1.5",
		// QUARTILE.EXC
		"=QUARTILE.EXC(A1:A4,1)": "0.25",
		"=QUARTILE.EXC(A1:A4,2)": "1.5",
		"=QUARTILE.EXC(A1:A4,3)": "2.75",
		// QUARTILE.INC
		"=QUARTILE.INC(A1:A4,0)": "0",
		// RANK
		"=RANK(1,A1:B5)":   "5",
		"=RANK(1,A1:B5,0)": "5",
		"=RANK(1,A1:B5,1)": "2",
		// RANK.EQ
		"=RANK.EQ(1,A1:B5)":   "5",
		"=RANK.EQ(1,A1:B5,0)": "5",
		"=RANK.EQ(1,A1:B5,1)": "2",
		// RSQ
		"=RSQ(A1:A4,B1:B4)": "1",
		// SKEW
		"=SKEW(1,2,3,4,3)": "-0.404796008910937",
		"=SKEW(A1:B2)":     "0",
		"=SKEW(A1:D3)":     "0",
		// SKEW.P
		"=SKEW.P(1,2,3,4,3)": "-0.27154541788364",
		"=SKEW.P(A1:B2)":     "0",
		"=SKEW.P(A1:D3)":     "0",
		// SLOPE
		"=SLOPE(A1:A4,B1:B4)": "1",
		// SMALL
		"=SMALL(A1:A5,1)": "0",
		"=SMALL(A1:B5,2)": "1",
		"=SMALL(A1,1)":    "1",
		"=SMALL(A1:F2,1)": "1",
		// STANDARDIZE
		"=STANDARDIZE(5.5,5,2)":   "0.25",
		"=STANDARDIZE(12,15,1.5)": "-2",
		"=STANDARDIZE(-2,0,5)":    "-0.4",
		// STDEVP
		"=STDEVP(A1:B2,6,-1)": "2.40947204913349",
		// STDEV.P
		"=STDEV.P(A1:B2,6,-1)": "2.40947204913349",
		// STDEVPA
		"=STDEVPA(1,3,5,2)":               "1.4790199457749",
		"=STDEVPA(1,3,5,2,1,0)":           "1.63299316185545",
		"=STDEVPA(1,3,5,2,TRUE,\"text\")": "1.63299316185545",
		// T.DIST
		"=T.DIST(1,10,TRUE)":   "0.82955343384897",
		"=T.DIST(-1,10,TRUE)":  "0.17044656615103",
		"=T.DIST(-1,10,FALSE)": "0.230361989229139",
		// T.DIST.2T
		"=T.DIST.2T(1,10)": "0.34089313230206",
		// T.DIST.RT
		"=T.DIST.RT(1,10)":  "0.17044656615103",
		"=T.DIST.RT(-1,10)": "0.82955343384897",
		// TDIST
		"=TDIST(1,10,1)": "0.17044656615103",
		"=TDIST(1,10,2)": "0.34089313230206",
		// T.INV
		"=T.INV(0.25,10)": "-0.699812061312432",
		"=T.INV(0.75,10)": "0.699812061312432",
		// T.INV.2T
		"=T.INV.2T(1,10)":   "0",
		"=T.INV.2T(0.5,10)": "0.699812061312432",
		// TINV
		"=TINV(1,10)":   "0",
		"=TINV(0.5,10)": "0.699812061312432",
		// TRIMMEAN
		"=TRIMMEAN(A1:B4,10%)": "2.5",
		"=TRIMMEAN(A1:B4,70%)": "2.5",
		// VAR
		"=VAR(1,3,5,0,C1)":      "4.91666666666667",
		"=VAR(1,3,5,0,C1,TRUE)": "4",
		// VARA
		"=VARA(1,3,5,0,C1)":      "4.91666666666667",
		"=VARA(1,3,5,0,C1,TRUE)": "4",
		// VARP
		"=VARP(A1:A5)":           "1.25",
		"=VARP(1,3,5,0,C1,TRUE)": "3.2",
		// VAR.P
		"=VAR.P(A1:A5)": "1.25",
		// VAR.S
		"=VAR.S(1,3,5,0,C1)":      "4.91666666666667",
		"=VAR.S(1,3,5,0,C1,TRUE)": "4",
		// VARPA
		"=VARPA(1,3,5,0,C1)":      "3.6875",
		"=VARPA(1,3,5,0,C1,TRUE)": "3.2",
		// WEIBULL
		"=WEIBULL(1,3,1,FALSE)":  "1.10363832351433",
		"=WEIBULL(2,5,1.5,TRUE)": "0.985212776817482",
		// WEIBULL.DIST
		"=WEIBULL.DIST(1,3,1,FALSE)":  "1.10363832351433",
		"=WEIBULL.DIST(2,5,1.5,TRUE)": "0.985212776817482",
		// Information Functions
		// ERROR.TYPE
		"=ERROR.TYPE(1/0)":           "2",
		"=ERROR.TYPE(COT(0))":        "2",
		"=ERROR.TYPE(XOR(\"text\"))": "3",
		"=ERROR.TYPE(HEX2BIN(2,1))":  "6",
		"=ERROR.TYPE(NA())":          "7",
		// ISBLANK
		"=ISBLANK(A1)": "FALSE",
		"=ISBLANK(A5)": "TRUE",
		// ISERR
		"=ISERR(A1)":           "FALSE",
		"=ISERR(NA())":         "FALSE",
		"=ISERR(POWER(0,-1)))": "TRUE",
		// ISERROR
		"=ISERROR(A1)":          "FALSE",
		"=ISERROR(NA())":        "TRUE",
		"=ISERROR(\"#VALUE!\")": "FALSE",
		// ISEVEN
		"=ISEVEN(A1)": "FALSE",
		"=ISEVEN(A2)": "TRUE",
		"=ISEVEN(G1)": "TRUE",
		// ISFORMULA
		"=ISFORMULA(A1)":    "FALSE",
		"=ISFORMULA(\"A\")": "FALSE",
		// ISLOGICAL
		"=ISLOGICAL(TRUE)":      "TRUE",
		"=ISLOGICAL(FALSE)":     "TRUE",
		"=ISLOGICAL(A1=A2)":     "TRUE",
		"=ISLOGICAL(\"true\")":  "TRUE",
		"=ISLOGICAL(\"false\")": "TRUE",
		"=ISLOGICAL(A1)":        "FALSE",
		"=ISLOGICAL(20/5)":      "FALSE",
		// ISNA
		"=ISNA(A1)":   "FALSE",
		"=ISNA(NA())": "TRUE",
		// ISNONTEXT
		"=ISNONTEXT(A1)":           "TRUE",
		"=ISNONTEXT(A5)":           "TRUE",
		"=ISNONTEXT(\"Excelize\")": "FALSE",
		"=ISNONTEXT(NA())":         "TRUE",
		// ISNUMBER
		"=ISNUMBER(A1)":    "TRUE",
		"=ISNUMBER(D1)":    "FALSE",
		"=ISNUMBER(A1:B1)": "TRUE",
		// ISODD
		"=ISODD(A1)": "TRUE",
		"=ISODD(A2)": "FALSE",
		// ISREF
		"=ISREF(B1)":       "TRUE",
		"=ISREF(B1:B2)":    "TRUE",
		"=ISREF(\"text\")": "FALSE",
		"=ISREF(B1*B2)":    "FALSE",
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
		"=SHEET()":           "1",
		"=SHEET(\"Sheet1\")": "1",
		// SHEETS
		"=SHEETS()":   "1",
		"=SHEETS(A1)": "1",
		// TYPE
		"=TYPE(2)":        "1",
		"=TYPE(10/2)":     "1",
		"=TYPE(C2)":       "1",
		"=TYPE(\"text\")": "2",
		"=TYPE(TRUE)":     "4",
		"=TYPE(NA())":     "16",
		"=TYPE(MUNIT(2))": "64",
		// T
		"=T(\"text\")": "text",
		"=T(N(10))":    "",
		// Logical Functions
		// AND
		"=AND(0)":                  "FALSE",
		"=AND(1)":                  "TRUE",
		"=AND(1,0)":                "FALSE",
		"=AND(0,1)":                "FALSE",
		"=AND(1=1)":                "TRUE",
		"=AND(1<2)":                "TRUE",
		"=AND(1>2,2<3,2>0,3>1)":    "FALSE",
		"=AND(1=1),1=1":            "TRUE",
		"=AND(\"TRUE\",\"FALSE\")": "FALSE",
		// FALSE
		"=FALSE()": "FALSE",
		// IFERROR
		"=IFERROR(1/2,0)":             "0.5",
		"=IFERROR(ISERROR(),0)":       "0",
		"=IFERROR(1/0,0)":             "0",
		"=IFERROR(G1,2)":              "0",
		"=IFERROR(B2/MROUND(A2,1),0)": "2.5",
		// IFNA
		"=IFNA(1,\"not found\")":                   "1",
		"=IFNA(NA(),\"not found\")":                "not found",
		"=IFNA(HLOOKUP(D2,D:D,1,2),\"not found\")": "not found",
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
		"=OR(1)":                  "TRUE",
		"=OR(0)":                  "FALSE",
		"=OR(1=2,2=2)":            "TRUE",
		"=OR(1=2,2=3)":            "FALSE",
		"=OR(1=1,2=3)":            "TRUE",
		"=OR(\"TRUE\",\"FALSE\")": "TRUE",
		"=OR(A1:B1)":              "TRUE",
		// SWITCH
		"=SWITCH(1,1,\"A\",2,\"B\",3,\"C\",\"N\")": "A",
		"=SWITCH(3,1,\"A\",2,\"B\",3,\"C\",\"N\")": "C",
		"=SWITCH(4,1,\"A\",2,\"B\",3,\"C\",\"N\")": "N",
		// TRUE
		"=TRUE()": "TRUE",
		// XOR
		"=XOR(1>0,2>0)":                       "FALSE",
		"=XOR(1>0,0>1)":                       "TRUE",
		"=XOR(1>0,0>1,INT(0),INT(1),A1:A4,2)": "FALSE",
		// Date and Time Functions
		// DATE
		"=DATE(2020,10,21)":   "44125",
		"=DATE(2020,10,21)+1": "44126",
		"=DATE(1900,1,1)":     "1",
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
		// DAYS360
		"=DAYS360(\"10/10/2020\", \"10/10/2020\")":       "0",
		"=DAYS360(\"01/30/1999\", \"02/28/1999\")":       "28",
		"=DAYS360(\"01/31/1999\", \"02/28/1999\")":       "28",
		"=DAYS360(\"12/12/1999\", \"08/31/1999\")":       "-101",
		"=DAYS360(\"12/12/1999\", \"11/30/1999\")":       "-12",
		"=DAYS360(\"12/12/1999\", \"11/30/1999\",TRUE)":  "-12",
		"=DAYS360(\"01/31/1999\", \"03/31/1999\",TRUE)":  "60",
		"=DAYS360(\"01/31/1999\", \"03/31/2000\",FALSE)": "420",
		// EDATE
		"=EDATE(\"01/01/2021\",-1)": "44166",
		"=EDATE(\"01/31/2020\",1)":  "43890",
		"=EDATE(\"01/29/2020\",12)": "44225",
		"=EDATE(\"6/12/2021\",-14)": "43933",
		// EOMONTH
		"=EOMONTH(\"01/01/2021\",-1)":  "44196",
		"=EOMONTH(\"01/29/2020\",12)":  "44227",
		"=EOMONTH(\"01/12/2021\",-18)": "43677",
		// HOUR
		"=HOUR(1)":                    "0",
		"=HOUR(43543.5032060185)":     "12",
		"=HOUR(\"43543.5032060185\")": "12",
		"=HOUR(\"13:00:55\")":         "13",
		"=HOUR(\"1:00 PM\")":          "13",
		"=HOUR(\"12/09/2015 08:55\")": "8",
		// ISOWEEKNUM
		"=ISOWEEKNUM(42370)":          "53",
		"=ISOWEEKNUM(\"42370\")":      "53",
		"=ISOWEEKNUM(\"01/01/2005\")": "53",
		"=ISOWEEKNUM(\"02/02/2005\")": "5",
		// MINUTE
		"=MINUTE(1)":                    "0",
		"=MINUTE(0.04)":                 "57",
		"=MINUTE(\"0.04\")":             "57",
		"=MINUTE(\"13:35:55\")":         "35",
		"=MINUTE(\"12/09/2015 08:55\")": "55",
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
		"=YEARFRAC(\"02/29/2000\", \"02/29/2008\",1)": "7.99817518248175",
		"=YEARFRAC(\"02/29/2000\", \"01/29/2001\",1)": "0.915300546448087",
		"=YEARFRAC(\"02/29/2000\", \"03/29/2000\",1)": "0.0792349726775956",
		"=YEARFRAC(\"01/31/2000\", \"03/29/2000\",4)": "0.163888888888889",
		// SECOND
		"=SECOND(\"13:35:55\")":            "55",
		"=SECOND(\"13:10:60\")":            "0",
		"=SECOND(\"13:10:61\")":            "1",
		"=SECOND(\"08:17:00\")":            "0",
		"=SECOND(\"12/09/2015 08:55\")":    "0",
		"=SECOND(\"12/09/2011 08:17:23\")": "23",
		"=SECOND(\"43543.5032060185\")":    "37",
		"=SECOND(43543.5032060185)":        "37",
		// TIME
		"=TIME(5,44,32)":             "0.239259259259259",
		"=TIME(\"5\",\"44\",\"32\")": "0.239259259259259",
		"=TIME(0,0,73)":              "0.000844907407407407",
		// TIMEVALUE
		"=TIMEVALUE(\"2:23\")":             "0.0993055555555555",
		"=TIMEVALUE(\"2:23 am\")":          "0.0993055555555555",
		"=TIMEVALUE(\"2:23 PM\")":          "0.599305555555556",
		"=TIMEVALUE(\"14:23:00\")":         "0.599305555555556",
		"=TIMEVALUE(\"00:02:23\")":         "0.00165509259259259",
		"=TIMEVALUE(\"01/01/2011 02:23\")": "0.0993055555555555",
		// WEEKDAY
		"=WEEKDAY(0)":                 "7",
		"=WEEKDAY(47119)":             "2",
		"=WEEKDAY(\"12/25/2012\")":    "3",
		"=WEEKDAY(\"12/25/2012\",1)":  "3",
		"=WEEKDAY(\"12/25/2012\",2)":  "2",
		"=WEEKDAY(\"12/25/2012\",3)":  "1",
		"=WEEKDAY(\"12/25/2012\",11)": "2",
		"=WEEKDAY(\"12/25/2012\",12)": "1",
		"=WEEKDAY(\"12/25/2012\",13)": "7",
		"=WEEKDAY(\"12/25/2012\",14)": "6",
		"=WEEKDAY(\"12/25/2012\",15)": "5",
		"=WEEKDAY(\"12/25/2012\",16)": "4",
		"=WEEKDAY(\"12/25/2012\",17)": "3",
		// WEEKNUM
		"=WEEKNUM(\"01/01/2011\")":    "1",
		"=WEEKNUM(\"01/03/2011\")":    "2",
		"=WEEKNUM(\"01/13/2008\")":    "3",
		"=WEEKNUM(\"01/21/2008\")":    "4",
		"=WEEKNUM(\"01/30/2008\")":    "5",
		"=WEEKNUM(\"02/04/2008\")":    "6",
		"=WEEKNUM(\"01/02/2017\",2)":  "2",
		"=WEEKNUM(\"01/02/2017\",12)": "1",
		"=WEEKNUM(\"12/31/2017\",21)": "52",
		"=WEEKNUM(\"01/01/2017\",21)": "52",
		"=WEEKNUM(\"01/01/2021\",21)": "53",
		// Text Functions
		// ARRAYTOTEXT
		"=ARRAYTOTEXT(A1:D2)":   "1, 4, , Month, 2, 5, , Jan",
		"=ARRAYTOTEXT(A1:D2,0)": "1, 4, , Month, 2, 5, , Jan",
		"=ARRAYTOTEXT(A1:D2,1)": "{1,4,,\"Month\";2,5,,\"Jan\"}",
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
		"=CONCAT(MUNIT(2))":                      "1001",
		"=CONCAT(A1:B2)":                         "1425",
		// CONCATENATE
		"=CONCATENATE(TRUE(),1,FALSE(),\"0\",INT(2))": "TRUE1FALSE02",
		"=CONCATENATE(MUNIT(2))":                      "1001",
		"=CONCATENATE(A1:B2)":                         "1425",
		// DBCS
		"=DBCS(\"\")":        "",
		"=DBCS(123.456)":     "123.456",
		"=DBCS(\"123.456\")": "123.456",
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
		"=FIND(\"s\",\"Sales\",2)":         "5",
		"=FIND(D1:E2,\"Month\")":           "1",
		// FINDB
		"=FINDB(\"T\",\"Original Text\")":   "10",
		"=FINDB(\"t\",\"Original Text\")":   "13",
		"=FINDB(\"i\",\"Original Text\")":   "3",
		"=FINDB(\"i\",\"Original Text\",4)": "5",
		"=FINDB(\"\",\"Original Text\")":    "1",
		"=FINDB(\"\",\"Original Text\",2)":  "2",
		"=FINDB(\"s\",\"Sales\",2)":         "5",
		// LEFT
		"=LEFT(\"Original Text\")":    "O",
		"=LEFT(\"Original Text\",4)":  "Orig",
		"=LEFT(\"Original Text\",0)":  "",
		"=LEFT(\"Original Text\",13)": "Original Text",
		"=LEFT(\"Original Text\",20)": "Original Text",
		"=LEFT(\"オリジナルテキスト\")":        "オ",
		"=LEFT(\"オリジナルテキスト\",2)":      "オリ",
		"=LEFT(\"オリジナルテキスト\",5)":      "オリジナル",
		"=LEFT(\"オリジナルテキスト\",7)":      "オリジナルテキ",
		"=LEFT(\"オリジナルテキスト\",20)":     "オリジナルテキスト",
		// LEFTB
		"=LEFTB(\"Original Text\")":    "O",
		"=LEFTB(\"Original Text\",4)":  "Orig",
		"=LEFTB(\"Original Text\",0)":  "",
		"=LEFTB(\"Original Text\",13)": "Original Text",
		"=LEFTB(\"Original Text\",20)": "Original Text",
		// LEN
		"=LEN(\"\")":              "0",
		"=LEN(D1)":                "5",
		"=LEN(\"テキスト\")":          "4",
		"=LEN(\"オリジナルテキスト\")":     "9",
		"=LEN(7+LEN(A1&B1&C1))":   "1",
		"=LEN(8+LEN(A1+(C1-B1)))": "2",
		// LENB
		"=LENB(\"\")":          "0",
		"=LENB(D1)":            "5",
		"=LENB(\"テキスト\")":      "8",
		"=LENB(\"オリジナルテキスト\")": "18",
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
		"=MID(\"你好World\",5,1)":       "r",
		"=MID(\"\u30AA\u30EA\u30B8\u30CA\u30EB\u30C6\u30AD\u30B9\u30C8\",6,4)": "\u30C6\u30AD\u30B9\u30C8",
		"=MID(\"\u30AA\u30EA\u30B8\u30CA\u30EB\u30C6\u30AD\u30B9\u30C8\",3,5)": "\u30B8\u30CA\u30EB\u30C6\u30AD",
		// MIDB
		"=MIDB(\"Original Text\",7,1)": "a",
		"=MIDB(\"Original Text\",4,7)": "ginal T",
		"=MIDB(\"255 years\",3,1)":     "5",
		"=MIDB(\"text\",3,6)":          "xt",
		"=MIDB(\"text\",6,0)":          "",
		"=MIDB(\"你好World\",5,1)":       "W",
		"=MIDB(\"\u30AA\u30EA\u30B8\u30CA\u30EB\u30C6\u30AD\u30B9\u30C8\",6,4)": "\u30B8\u30CA",
		"=MIDB(\"\u30AA\u30EA\u30B8\u30CA\u30EB\u30C6\u30AD\u30B9\u30C8\",3,5)": "\u30EA\u30B8\xe3",
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
		"=RIGHT(\"オリジナルテキスト\")":        "ト",
		"=RIGHT(\"オリジナルテキスト\",2)":      "スト",
		"=RIGHT(\"オリジナルテキスト\",4)":      "テキスト",
		"=RIGHT(\"オリジナルテキスト\",7)":      "ジナルテキスト",
		"=RIGHT(\"オリジナルテキスト\",20)":     "オリジナルテキスト",
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
		// TEXT
		"=TEXT(\"07/07/2015\",\"mm/dd/yyyy\")":        "07/07/2015",
		"=TEXT(42192,\"mm/dd/yyyy\")":                 "07/07/2015",
		"=TEXT(42192,\"mmm dd yyyy\")":                "Jul 07 2015",
		"=TEXT(0.75,\"hh:mm\")":                       "18:00",
		"=TEXT(36.363636,\"0.00\")":                   "36.36",
		"=TEXT(567.9,\"$#,##0.00\")":                  "$567.90",
		"=TEXT(-5,\"+ $#,##0.00;- $#,##0.00;$0.00\")": "- $5.00",
		"=TEXT(5,\"+ $#,##0.00;- $#,##0.00;$0.00\")":  "+ $5.00",
		// TEXTAFTER
		"=TEXTAFTER(\"Red riding hood's, red hood\",\"hood\")":               "'s, red hood",
		"=TEXTAFTER(\"Red riding hood's, red hood\",\"HOOD\",1,1)":           "'s, red hood",
		"=TEXTAFTER(\"Red riding hood's, red hood\",\"basket\",1,0,0,\"x\")": "x",
		"=TEXTAFTER(\"Red riding hood's, red hood\",\"basket\",1,0,1,\"x\")": "",
		"=TEXTAFTER(\"Red riding hood's, red hood\",\"hood\",-1)":            "",
		"=TEXTAFTER(\"Jones,Bob\",\",\")":                                    "Bob",
		"=TEXTAFTER(\"12 ft x 20 ft\",\" x \")":                              "20 ft",
		"=TEXTAFTER(\"ABX-112-Red-Y\",\"-\",1)":                              "112-Red-Y",
		"=TEXTAFTER(\"ABX-112-Red-Y\",\"-\",2)":                              "Red-Y",
		"=TEXTAFTER(\"ABX-112-Red-Y\",\"-\",-1)":                             "Y",
		"=TEXTAFTER(\"ABX-112-Red-Y\",\"-\",-2)":                             "Red-Y",
		"=TEXTAFTER(\"ABX-112-Red-Y\",\"-\",-3)":                             "112-Red-Y",
		"=TEXTAFTER(\"ABX-123-Red-XYZ\",\"-\",-4,0,1)":                       "ABX-123-Red-XYZ",
		"=TEXTAFTER(\"ABX-123-Red-XYZ\",\"A\")":                              "BX-123-Red-XYZ",
		// TEXTBEFORE
		"=TEXTBEFORE(\"Red riding hood's, red hood\",\"hood\")":               "Red riding ",
		"=TEXTBEFORE(\"Red riding hood's, red hood\",\"HOOD\",1,1)":           "Red riding ",
		"=TEXTBEFORE(\"Red riding hood's, red hood\",\"basket\",1,0,0,\"x\")": "x",
		"=TEXTBEFORE(\"Red riding hood's, red hood\",\"basket\",1,0,1,\"x\")": "Red riding hood's, red hood",
		"=TEXTBEFORE(\"Red riding hood's, red hood\",\"hood\",-1)":            "Red riding hood's, red ",
		"=TEXTBEFORE(\"Jones,Bob\",\",\")":                                    "Jones",
		"=TEXTBEFORE(\"12 ft x 20 ft\",\" x \")":                              "12 ft",
		"=TEXTBEFORE(\"ABX-112-Red-Y\",\"-\",1)":                              "ABX",
		"=TEXTBEFORE(\"ABX-112-Red-Y\",\"-\",2)":                              "ABX-112",
		"=TEXTBEFORE(\"ABX-112-Red-Y\",\"-\",-1)":                             "ABX-112-Red",
		"=TEXTBEFORE(\"ABX-112-Red-Y\",\"-\",-2)":                             "ABX-112",
		"=TEXTBEFORE(\"ABX-123-Red-XYZ\",\"-\",4,0,1)":                        "ABX-123-Red-XYZ",
		"=TEXTBEFORE(\"ABX-112-Red-Y\",\"A\")":                                "",
		// TEXTJOIN
		"=TEXTJOIN(\"-\",TRUE,1,2,3,4)":  "1-2-3-4",
		"=TEXTJOIN(A4,TRUE,A1:B2)":       "1040205",
		"=TEXTJOIN(\",\",FALSE,A1:C2)":   "1,4,,2,5,",
		"=TEXTJOIN(\",\",TRUE,A1:C2)":    "1,4,2,5",
		"=TEXTJOIN(\",\",TRUE,MUNIT(2))": "1,0,0,1",
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
		// VALUE
		"=VALUE(\"50\")":                  "50",
		"=VALUE(\"1.0E-07\")":             "0.0000001",
		"=VALUE(\"5,000\")":               "5000",
		"=VALUE(\"20%\")":                 "0.2",
		"=VALUE(\"12:00:00\")":            "0.5",
		"=VALUE(\"01/02/2006 15:04:05\")": "38719.6278356481",
		// VALUETOTEXT
		"=VALUETOTEXT(A1)":   "1",
		"=VALUETOTEXT(A1,0)": "1",
		"=VALUETOTEXT(A1,1)": "1",
		"=VALUETOTEXT(D1)":   "Month",
		"=VALUETOTEXT(D1,0)": "Month",
		"=VALUETOTEXT(D1,1)": "\"Month\"",
		// Conditional Functions
		// IF
		"=IF(1=1)":                                   "TRUE",
		"=IF(1<>1)":                                  "FALSE",
		"=IF(5<0, \"negative\", \"positive\")":       "positive",
		"=IF(-2<0, \"negative\", \"positive\")":      "negative",
		"=IF(1=1, \"equal\", \"notequal\")":          "equal",
		"=IF(1<>1, \"equal\", \"notequal\")":         "notequal",
		"=IF(\"A\"=\"A\", \"equal\", \"notequal\")":  "equal",
		"=IF(\"A\"<>\"A\", \"equal\", \"notequal\")": "notequal",
		"=IF(FALSE,0,ROUND(4/2,0))":                  "2",
		"=IF(TRUE,ROUND(4/2,0),0)":                   "2",
		"=IF(A4>0.4,\"TRUE\",\"FALSE\")":             "FALSE",
		// Excel Lookup and Reference Functions
		// ADDRESS
		"=ADDRESS(1,1,1,TRUE)":            "$A$1",
		"=ADDRESS(1,2,1,TRUE)":            "$B$1",
		"=ADDRESS(1,1,1,FALSE)":           "R1C1",
		"=ADDRESS(1,2,1,FALSE)":           "R1C2",
		"=ADDRESS(1,1,2,TRUE)":            "A$1",
		"=ADDRESS(1,2,2,TRUE)":            "B$1",
		"=ADDRESS(1,1,2,FALSE)":           "R1C[1]",
		"=ADDRESS(1,2,2,FALSE)":           "R1C[2]",
		"=ADDRESS(1,1,3,TRUE)":            "$A1",
		"=ADDRESS(1,2,3,TRUE)":            "$B1",
		"=ADDRESS(1,1,3,FALSE)":           "R[1]C1",
		"=ADDRESS(1,2,3,FALSE)":           "R[1]C2",
		"=ADDRESS(1,1,4,TRUE)":            "A1",
		"=ADDRESS(1,2,4,TRUE)":            "B1",
		"=ADDRESS(1,1,4,FALSE)":           "R[1]C[1]",
		"=ADDRESS(1,2,4,FALSE)":           "R[1]C[2]",
		"=ADDRESS(1,1,4,TRUE,\"\")":       "!A1",
		"=ADDRESS(1,2,4,TRUE,\"\")":       "!B1",
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
		// HYPERLINK
		"=HYPERLINK(\"https://github.com/xuri/excelize\")":              "https://github.com/xuri/excelize",
		"=HYPERLINK(\"https://github.com/xuri/excelize\",\"Excelize\")": "Excelize",
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
		"=VLOOKUP(A1:A2,A1:A1,1)":             "1",
		"=VLOOKUP(MUNIT(1),MUNIT(1),1,FALSE)": "1",
		// INDEX
		"=INDEX(0,0,0)":          "0",
		"=INDEX(A1,0,0)":         "1",
		"=INDEX(A1:A1,0,0)":      "1",
		"=SUM(INDEX(A1:B1,1))":   "5",
		"=SUM(INDEX(A1:B1,1,0))": "5",
		"=SUM(INDEX(A1:B2,2,0))": "7",
		"=SUM(INDEX(A1:B4,0,2))": "9",
		"=SUM(INDEX(E1:F5,5,2))": "34440",
		// INDIRECT
		"=INDIRECT(\"E1\")":                   "Team",
		"=INDIRECT(\"E\"&1)":                  "Team",
		"=INDIRECT(\"E\"&ROW())":              "Team",
		"=INDIRECT(\"E\"&ROW(),TRUE)":         "Team",
		"=INDIRECT(\"R1C5\",FALSE)":           "Team",
		"=INDIRECT(\"R\"&1&\"C\"&5,FALSE)":    "Team",
		"=SUM(INDIRECT(\"A1:B2\"))":           "12",
		"=SUM(INDIRECT(\"A1:B2\",TRUE))":      "12",
		"=SUM(INDIRECT(\"R1C1:R2C2\",FALSE))": "12",
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
		// ACCRINT
		"=ACCRINT(\"01/01/2012\",\"04/01/2012\",\"12/31/2013\",8%,10000,4,0,TRUE)":  "1600",
		"=ACCRINT(\"01/01/2012\",\"04/01/2012\",\"12/31/2013\",8%,10000,4,0,FALSE)": "1600",
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
		// COUPDAYBS
		"=COUPDAYBS(\"02/24/2000\",\"11/24/2000\",4,4)": "0",
		"=COUPDAYBS(\"03/27/2000\",\"11/29/2000\",4,4)": "28",
		"=COUPDAYBS(\"02/29/2000\",\"04/01/2000\",4,4)": "58",
		"=COUPDAYBS(\"01/01/2011\",\"10/25/2012\",4)":   "66",
		"=COUPDAYBS(\"01/01/2011\",\"10/25/2012\",4,1)": "68",
		"=COUPDAYBS(\"10/31/2011\",\"02/26/2012\",4,0)": "65",
		// COUPDAYS
		"=COUPDAYS(\"01/01/2011\",\"10/25/2012\",4)":   "90",
		"=COUPDAYS(\"01/01/2011\",\"10/25/2012\",4,1)": "92",
		// COUPDAYSNC
		"=COUPDAYSNC(\"01/01/2011\",\"10/25/2012\",4)": "24",
		"=COUPDAYSNC(\"04/01/2012\",\"03/31/2020\",2)": "179",
		// COUPNCD
		"=COUPNCD(\"01/01/2011\",\"10/25/2012\",4)":   "40568",
		"=COUPNCD(\"01/01/2011\",\"10/25/2012\",4,0)": "40568",
		"=COUPNCD(\"10/25/2011\",\"01/01/2012\",4)":   "40909",
		"=COUPNCD(\"04/01/2012\",\"03/31/2020\",2)":   "41182",
		"=COUPNCD(\"01/01/2000\",\"08/30/2001\",2)":   "36585",
		// COUPNUM
		"=COUPNUM(\"01/01/2011\",\"10/25/2012\",4)":   "8",
		"=COUPNUM(\"01/01/2011\",\"10/25/2012\",4,0)": "8",
		"=COUPNUM(\"09/30/2017\",\"03/31/2021\",4,0)": "14",
		// COUPPCD
		"=COUPPCD(\"01/01/2011\",\"10/25/2012\",4)":   "40476",
		"=COUPPCD(\"01/01/2011\",\"10/25/2012\",4,0)": "40476",
		"=COUPPCD(\"10/25/2011\",\"01/01/2012\",4)":   "40817",
		// CUMIPMT
		"=CUMIPMT(0.05/12,60,50000,1,12,0)":  "-2294.97753732664",
		"=CUMIPMT(0.05/12,60,50000,13,24,0)": "-1833.10006657389",
		// CUMPRINC
		"=CUMPRINC(0.05/12,60,50000,1,12,0)":  "-9027.76264907988",
		"=CUMPRINC(0.05/12,60,50000,13,24,0)": "-9489.64011983263",
		// DB
		"=DB(0,1000,5,1)":       "0",
		"=DB(10000,1000,5,1)":   "3690",
		"=DB(10000,1000,5,2)":   "2328.39",
		"=DB(10000,1000,5,1,6)": "1845",
		"=DB(10000,1000,5,6,6)": "238.527124587882",
		// DDB
		"=DDB(0,1000,5,1)":     "0",
		"=DDB(10000,1000,5,1)": "4000",
		"=DDB(10000,1000,5,2)": "2400",
		"=DDB(10000,1000,5,3)": "1440",
		"=DDB(10000,1000,5,4)": "864",
		"=DDB(10000,1000,5,5)": "296",
		// DISC
		"=DISC(\"04/01/2016\",\"03/31/2021\",95,100)": "0.01",
		// DOLLAR
		"=DOLLAR(1234.56)":     "$1,234.56",
		"=DOLLAR(1234.56,0)":   "$1,235",
		"=DOLLAR(1234.56,1)":   "$1,234.6",
		"=DOLLAR(1234.56,2)":   "$1,234.56",
		"=DOLLAR(1234.56,3)":   "$1,234.560",
		"=DOLLAR(1234.56,-2)":  "$1,200",
		"=DOLLAR(1234.56,-3)":  "$1,000",
		"=DOLLAR(-1234.56,3)":  "($1,234.560)",
		"=DOLLAR(-1234.56,-3)": "($1,000)",
		// DOLLARDE
		"=DOLLARDE(1.01,16)": "1.0625",
		// DOLLARFR
		"=DOLLARFR(1.0625,16)": "1.01",
		// DURATION
		"=DURATION(\"04/01/2015\",\"03/31/2025\",10%,8%,4)": "6.67442279848313",
		// EFFECT
		"=EFFECT(0.1,4)":   "0.103812890625",
		"=EFFECT(0.025,2)": "0.02515625",
		// EUROCONVERT
		"=EUROCONVERT(1.47,\"EUR\",\"EUR\")":         "1.47",
		"=EUROCONVERT(1.47,\"EUR\",\"DEM\")":         "2.88",
		"=EUROCONVERT(1.47,\"FRF\",\"DEM\")":         "0.44",
		"=EUROCONVERT(1.47,\"FRF\",\"DEM\",FALSE)":   "0.44",
		"=EUROCONVERT(1.47,\"FRF\",\"DEM\",FALSE,3)": "0.44",
		"=EUROCONVERT(1.47,\"FRF\",\"DEM\",TRUE,3)":  "0.43810592",
		// FV
		"=FV(0.05/12,60,-1000)":   "68006.0828408434",
		"=FV(0.1/4,16,-2000,0,1)": "39729.4608941662",
		"=FV(0,16,-2000)":         "32000",
		// FVSCHEDULE
		"=FVSCHEDULE(10000,A1:A5)": "240000",
		"=FVSCHEDULE(10000,0.5)":   "15000",
		// INTRATE
		"=INTRATE(\"04/01/2005\",\"03/31/2010\",1000,2125)": "0.225",
		// IPMT
		"=IPMT(0.05/12,2,60,50000)":   "-205.26988187972",
		"=IPMT(0.035/4,2,8,0,5000,1)": "5.25745523782908",
		// ISPMT
		"=ISPMT(0.05/12,1,60,50000)": "-204.861111111111",
		"=ISPMT(0.05/12,2,60,50000)": "-201.388888888889",
		"=ISPMT(0.05/12,2,1,50000)":  "208.333333333333",
		// MDURATION
		"=MDURATION(\"04/01/2015\",\"03/31/2025\",10%,8%,4)": "6.54355176321876",
		// NOMINAL
		"=NOMINAL(0.025,12)": "0.0247180352381129",
		// NPER
		"=NPER(0.04,-6000,50000)":           "10.3380350715077",
		"=NPER(0,-6000,50000)":              "8.33333333333333",
		"=NPER(0.06/4,-2000,60000,30000,1)": "52.7947737092748",
		// NPV
		"=NPV(0.02,-5000,\"\",800)": "-4133.02575932334",
		// ODDFPRICE
		"=ODDFPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,100,2)":       "107.691830256629",
		"=ODDFPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,100,4,1)":     "106.766915010929",
		"=ODDFPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,100,4,3)":     "106.7819138147",
		"=ODDFPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,100,4,4)":     "106.771913772467",
		"=ODDFPRICE(\"11/11/2008\",\"03/01/2021\",\"10/15/2008\",\"03/01/2009\",7.85%,6.25%,100,2,1)":   "113.597717474079",
		"=ODDFPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"09/30/2017\",5.5%,3.5%,100,4,0)":     "106.72930611878",
		"=ODDFPRICE(\"11/11/2008\",\"03/29/2021\",\"08/15/2008\",\"03/29/2009\",0.0785,0.0625,100,2,1)": "113.61826640814",
		// ODDFYIELD
		"=ODDFYIELD(\"05/01/2017\",\"06/30/2021\",\"03/15/2017\",\"06/30/2017\",5.5%,102,100,1)":   "0.0495998049937776",
		"=ODDFYIELD(\"05/01/2017\",\"06/30/2021\",\"03/15/2017\",\"06/30/2017\",5.5%,102,100,2)":   "0.0496289417392839",
		"=ODDFYIELD(\"05/01/2017\",\"06/30/2021\",\"03/15/2017\",\"06/30/2017\",5.5%,102,100,4,1)": "0.0464750282973541",
		// ODDLPRICE
		"=ODDLPRICE(\"04/20/2008\",\"06/15/2008\",\"12/24/2007\",3.75%,99.875,100,2)":   "5.0517841252892",
		"=ODDLPRICE(\"04/20/2008\",\"06/15/2008\",\"12/24/2007\",3.75%,99.875,100,4,1)": "10.3667274303228",
		// ODDLYIELD
		"=ODDLYIELD(\"04/20/2008\",\"06/15/2008\",\"12/24/2007\",3.75%,99.875,100,2)":   "0.0451922356291692",
		"=ODDLYIELD(\"04/20/2008\",\"06/15/2008\",\"12/24/2007\",3.75%,99.875,100,4,1)": "0.0882287538349037",
		// PDURATION
		"=PDURATION(0.04,10000,15000)": "10.3380350715076",
		// PMT
		"=PMT(0,8,0,5000,1)":       "-625",
		"=PMT(0.035/4,8,0,5000,1)": "-600.852027180466",
		// PRICE
		"=PRICE(\"04/01/2012\",\"02/01/2020\",12%,10%,100,2)":   "110.655105178443",
		"=PRICE(\"04/01/2012\",\"02/01/2020\",12%,10%,100,2,4)": "110.655105178443",
		"=PRICE(\"04/01/2012\",\"03/31/2020\",12%,10%,100,2)":   "110.834483593216",
		"=PRICE(\"01/01/2010\",\"06/30/2010\",0.5,1,1,1,4)":     "8.92419088847661",
		// PPMT
		"=PPMT(0.05/12,2,60,50000)":   "-738.291800320824",
		"=PPMT(0.035/4,2,8,0,5000,1)": "-606.109482418295",
		// PRICEDISC
		"=PRICEDISC(\"04/01/2017\",\"03/31/2021\",2.5%,100)":   "90",
		"=PRICEDISC(\"04/01/2017\",\"03/31/2021\",2.5%,100,3)": "90",
		"=PRICEDISC(\"42826\",\"03/31/2021\",2.5%,100,3)":      "90",
		// PRICEMAT
		"=PRICEMAT(\"04/01/2017\",\"03/31/2021\",\"01/01/2017\",4.5%,2.5%)":   "107.170454545455",
		"=PRICEMAT(\"04/01/2017\",\"03/31/2021\",\"01/01/2017\",4.5%,2.5%,0)": "107.170454545455",
		// PV
		"=PV(0,60,1000)":         "-60000",
		"=PV(5%/12,60,1000)":     "-52990.7063239275",
		"=PV(10%/4,16,2000,0,1)": "-26762.7554528811",
		// RATE
		"=RATE(60,-1000,50000)":       "0.0061834131621292",
		"=RATE(24,-800,0,20000,1)":    "0.00325084350160374",
		"=RATE(48,-200,8000,3,1,0.5)": "0.0080412665831637",
		// RECEIVED
		"=RECEIVED(\"04/01/2011\",\"03/31/2016\",1000,4.5%)":   "1290.32258064516",
		"=RECEIVED(\"04/01/2011\",\"03/31/2016\",1000,4.5%,0)": "1290.32258064516",
		// RRI
		"=RRI(10,10000,15000)": "0.0413797439924106",
		// SLN
		"=SLN(10000,1000,5)": "1800",
		// SYD
		"=SYD(10000,1000,5,1)": "3000",
		"=SYD(10000,1000,5,2)": "2400",
		// TBILLEQ
		"=TBILLEQ(\"01/01/2017\",\"06/30/2017\",2.5%)": "0.0256680731364276",
		// TBILLPRICE
		"=TBILLPRICE(\"02/01/2017\",\"06/30/2017\",2.75%)": "98.8618055555556",
		// TBILLYIELD
		"=TBILLYIELD(\"02/01/2017\",\"06/30/2017\",99)": "0.024405125076266",
		// VDB
		"=VDB(10000,1000,5,0,1)":           "4000",
		"=VDB(10000,1000,5,1,3)":           "3840",
		"=VDB(10000,1000,5,3,5)":           "1160",
		"=VDB(10000,1000,5,3,5,0.2,FALSE)": "3600",
		"=VDB(10000,1000,5,3,5,0.2,TRUE)":  "693.633024",
		"=VDB(24000,3000,10,0,0.875,2)":    "4200",
		"=VDB(24000,3000,10,0.1,1)":        "4233.6",
		"=VDB(24000,3000,10,0.1,1,1)":      "2138.4",
		"=VDB(24000,3000,100,50,100,1)":    "10377.2944184652",
		"=VDB(24000,3000,100,50,100,2)":    "5740.0723220908",
		// YIELD
		"=YIELD(\"01/01/2010\",\"06/30/2015\",10%,101,100,4)":               "0.0975631546829798",
		"=YIELD(\"01/01/2010\",\"06/30/2015\",10%,101,100,4,4)":             "0.0976269355643988",
		"=YIELD(\"01/01/2010\",\"06/30/2010\",0.5,1,1,1,4)":                 "1.91285866099894",
		"=YIELD(\"01/01/2010\",\"06/30/2010\",0,1,1,1,4)":                   "0",
		"=YIELD(\"01/01/2010\",\"01/02/2020\",100,68.15518653988686,1,1,1)": "64",
		// YIELDDISC
		"=YIELDDISC(\"01/01/2017\",\"06/30/2017\",97,100)":   "0.0622012325059031",
		"=YIELDDISC(\"01/01/2017\",\"06/30/2017\",97,100,0)": "0.0622012325059031",
		// YIELDMAT
		"=YIELDMAT(\"01/01/2017\",\"06/30/2018\",\"06/01/2014\",5.5%,101)":   "0.0419422478838651",
		"=YIELDMAT(\"01/01/2017\",\"06/30/2018\",\"06/01/2014\",5.5%,101,0)": "0.0419422478838651",
		// DISPIMG
		"=_xlfn.DISPIMG(\"ID_********************************\",1)": "ID_********************************",
	}
	for formula, expected := range mathCalc {
		f := prepareCalcData(cellData)
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	mathCalcError := map[string][]string{
		"=1/0":       {"", "#DIV/0!"},
		"1^\"text\"": {"", "strconv.ParseFloat: parsing \"text\": invalid syntax"},
		"\"text\"^1": {"", "strconv.ParseFloat: parsing \"text\": invalid syntax"},
		"1+\"text\"": {"", "strconv.ParseFloat: parsing \"text\": invalid syntax"},
		"\"text\"+1": {"", "strconv.ParseFloat: parsing \"text\": invalid syntax"},
		"1-\"text\"": {"", "strconv.ParseFloat: parsing \"text\": invalid syntax"},
		"\"text\"-1": {"", "strconv.ParseFloat: parsing \"text\": invalid syntax"},
		"1*\"text\"": {"", "strconv.ParseFloat: parsing \"text\": invalid syntax"},
		"\"text\"*1": {"", "strconv.ParseFloat: parsing \"text\": invalid syntax"},
		"1/\"text\"": {"", "strconv.ParseFloat: parsing \"text\": invalid syntax"},
		"\"text\"/1": {"", "strconv.ParseFloat: parsing \"text\": invalid syntax"},
		// Engineering Functions
		// BESSELI
		"=BESSELI()":       {"#VALUE!", "BESSELI requires 2 numeric arguments"},
		"=BESSELI(\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BESSELI(0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// BESSELJ
		"=BESSELJ()":       {"#VALUE!", "BESSELJ requires 2 numeric arguments"},
		"=BESSELJ(\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BESSELJ(0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// BESSELK
		"=BESSELK()":       {"#VALUE!", "BESSELK requires 2 numeric arguments"},
		"=BESSELK(\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BESSELK(0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BESSELK(-1,0)":   {"#NUM!", "#NUM!"},
		"=BESSELK(1,-1)":   {"#NUM!", "#NUM!"},
		// BESSELY
		"=BESSELY()":       {"#VALUE!", "BESSELY requires 2 numeric arguments"},
		"=BESSELY(\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BESSELY(0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BESSELY(-1,0)":   {"#NUM!", "#NUM!"},
		"=BESSELY(1,-1)":   {"#NUM!", "#NUM!"},
		// BIN2DEC
		"=BIN2DEC()":     {"#VALUE!", "BIN2DEC requires 1 numeric argument"},
		"=BIN2DEC(\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// BIN2HEX
		"=BIN2HEX()":               {"#VALUE!", "BIN2HEX requires at least 1 argument"},
		"=BIN2HEX(1,1,1)":          {"#VALUE!", "BIN2HEX allows at most 2 arguments"},
		"=BIN2HEX(\"\",1)":         {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BIN2HEX(1,\"\")":         {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BIN2HEX(12345678901,10)": {"#NUM!", "#NUM!"},
		"=BIN2HEX(1,-1)":           {"#NUM!", "#NUM!"},
		"=BIN2HEX(31,1)":           {"#NUM!", "#NUM!"},
		// BIN2OCT
		"=BIN2OCT()":                 {"#VALUE!", "BIN2OCT requires at least 1 argument"},
		"=BIN2OCT(1,1,1)":            {"#VALUE!", "BIN2OCT allows at most 2 arguments"},
		"=BIN2OCT(\"\",1)":           {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BIN2OCT(1,\"\")":           {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BIN2OCT(-12345678901 ,10)": {"#NUM!", "#NUM!"},
		"=BIN2OCT(1,-1)":             {"#NUM!", "#NUM!"},
		"=BIN2OCT(8,1)":              {"#NUM!", "#NUM!"},
		// BITAND
		"=BITAND()":        {"#VALUE!", "BITAND requires 2 numeric arguments"},
		"=BITAND(-1,2)":    {"#NUM!", "#NUM!"},
		"=BITAND(2^48,2)":  {"#NUM!", "#NUM!"},
		"=BITAND(1,-1)":    {"#NUM!", "#NUM!"},
		"=BITAND(\"\",-1)": {"#NUM!", "#NUM!"},
		"=BITAND(1,\"\")":  {"#NUM!", "#NUM!"},
		"=BITAND(1,2^48)":  {"#NUM!", "#NUM!"},
		// BITLSHIFT
		"=BITLSHIFT()":        {"#VALUE!", "BITLSHIFT requires 2 numeric arguments"},
		"=BITLSHIFT(-1,2)":    {"#NUM!", "#NUM!"},
		"=BITLSHIFT(2^48,2)":  {"#NUM!", "#NUM!"},
		"=BITLSHIFT(1,-1)":    {"#NUM!", "#NUM!"},
		"=BITLSHIFT(\"\",-1)": {"#NUM!", "#NUM!"},
		"=BITLSHIFT(1,\"\")":  {"#NUM!", "#NUM!"},
		"=BITLSHIFT(1,2^48)":  {"#NUM!", "#NUM!"},
		// BITOR
		"=BITOR()":        {"#VALUE!", "BITOR requires 2 numeric arguments"},
		"=BITOR(-1,2)":    {"#NUM!", "#NUM!"},
		"=BITOR(2^48,2)":  {"#NUM!", "#NUM!"},
		"=BITOR(1,-1)":    {"#NUM!", "#NUM!"},
		"=BITOR(\"\",-1)": {"#NUM!", "#NUM!"},
		"=BITOR(1,\"\")":  {"#NUM!", "#NUM!"},
		"=BITOR(1,2^48)":  {"#NUM!", "#NUM!"},
		// BITRSHIFT
		"=BITRSHIFT()":        {"#VALUE!", "BITRSHIFT requires 2 numeric arguments"},
		"=BITRSHIFT(-1,2)":    {"#NUM!", "#NUM!"},
		"=BITRSHIFT(2^48,2)":  {"#NUM!", "#NUM!"},
		"=BITRSHIFT(1,-1)":    {"#NUM!", "#NUM!"},
		"=BITRSHIFT(\"\",-1)": {"#NUM!", "#NUM!"},
		"=BITRSHIFT(1,\"\")":  {"#NUM!", "#NUM!"},
		"=BITRSHIFT(1,2^48)":  {"#NUM!", "#NUM!"},
		// BITXOR
		"=BITXOR()":        {"#VALUE!", "BITXOR requires 2 numeric arguments"},
		"=BITXOR(-1,2)":    {"#NUM!", "#NUM!"},
		"=BITXOR(2^48,2)":  {"#NUM!", "#NUM!"},
		"=BITXOR(1,-1)":    {"#NUM!", "#NUM!"},
		"=BITXOR(\"\",-1)": {"#NUM!", "#NUM!"},
		"=BITXOR(1,\"\")":  {"#NUM!", "#NUM!"},
		"=BITXOR(1,2^48)":  {"#NUM!", "#NUM!"},
		// COMPLEX
		"=COMPLEX()":              {"#VALUE!", "COMPLEX requires at least 2 arguments"},
		"=COMPLEX(10,-5,\"\")":    {"#VALUE!", "#VALUE!"},
		"=COMPLEX(\"\",0)":        {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=COMPLEX(0,\"\")":        {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=COMPLEX(10,-5,\"i\",0)": {"#VALUE!", "COMPLEX allows at most 3 arguments"},
		// CONVERT
		"=CONVERT()":                          {"#VALUE!", "CONVERT requires 3 arguments"},
		"=CONVERT(\"\",\"m\",\"yd\")":         {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CONVERT(20.2,\"m\",\"C\")":          {"#N/A", "#N/A"},
		"=CONVERT(20.2,\"\",\"C\")":           {"#N/A", "#N/A"},
		"=CONVERT(100,\"dapt\",\"pt\")":       {"#N/A", "#N/A"},
		"=CONVERT(1,\"ft\",\"day\")":          {"#N/A", "#N/A"},
		"=CONVERT(234.56,\"kpt\",\"lt\")":     {"#N/A", "#N/A"},
		"=CONVERT(234.56,\"lt\",\"kpt\")":     {"#N/A", "#N/A"},
		"=CONVERT(234.56,\"kiqt\",\"pt\")":    {"#N/A", "#N/A"},
		"=CONVERT(234.56,\"pt\",\"kiqt\")":    {"#N/A", "#N/A"},
		"=CONVERT(12345.6,\"baton\",\"cwt\")": {"#N/A", "#N/A"},
		"=CONVERT(12345.6,\"cwt\",\"baton\")": {"#N/A", "#N/A"},
		"=CONVERT(234.56,\"xxxx\",\"m\")":     {"#N/A", "#N/A"},
		"=CONVERT(234.56,\"m\",\"xxxx\")":     {"#N/A", "#N/A"},
		// DEC2BIN
		"=DEC2BIN()":        {"#VALUE!", "DEC2BIN requires at least 1 argument"},
		"=DEC2BIN(1,1,1)":   {"#VALUE!", "DEC2BIN allows at most 2 arguments"},
		"=DEC2BIN(\"\",1)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DEC2BIN(1,\"\")":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DEC2BIN(-513,10)": {"#NUM!", "#NUM!"},
		"=DEC2BIN(1,-1)":    {"#NUM!", "#NUM!"},
		"=DEC2BIN(2,1)":     {"#NUM!", "#NUM!"},
		// DEC2HEX
		"=DEC2HEX()":                 {"#VALUE!", "DEC2HEX requires at least 1 argument"},
		"=DEC2HEX(1,1,1)":            {"#VALUE!", "DEC2HEX allows at most 2 arguments"},
		"=DEC2HEX(\"\",1)":           {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DEC2HEX(1,\"\")":           {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DEC2HEX(-549755813888,10)": {"#NUM!", "#NUM!"},
		"=DEC2HEX(1,-1)":             {"#NUM!", "#NUM!"},
		"=DEC2HEX(31,1)":             {"#NUM!", "#NUM!"},
		// DEC2OCT
		"=DEC2OCT()":               {"#VALUE!", "DEC2OCT requires at least 1 argument"},
		"=DEC2OCT(1,1,1)":          {"#VALUE!", "DEC2OCT allows at most 2 arguments"},
		"=DEC2OCT(\"\",1)":         {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DEC2OCT(1,\"\")":         {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DEC2OCT(-536870912 ,10)": {"#NUM!", "#NUM!"},
		"=DEC2OCT(1,-1)":           {"#NUM!", "#NUM!"},
		"=DEC2OCT(8,1)":            {"#NUM!", "#NUM!"},
		// DELTA
		"=DELTA()":       {"#VALUE!", "DELTA requires at least 1 argument"},
		"=DELTA(0,0,0)":  {"#VALUE!", "DELTA allows at most 2 arguments"},
		"=DELTA(\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DELTA(0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// ERF
		"=ERF()":       {"#VALUE!", "ERF requires at least 1 argument"},
		"=ERF(0,0,0)":  {"#VALUE!", "ERF allows at most 2 arguments"},
		"=ERF(\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ERF(0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// ERF.PRECISE
		"=ERF.PRECISE()":     {"#VALUE!", "ERF.PRECISE requires 1 argument"},
		"=ERF.PRECISE(\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// ERFC
		"=ERFC()":     {"#VALUE!", "ERFC requires 1 argument"},
		"=ERFC(\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// ERFC.PRECISE
		"=ERFC.PRECISE()":     {"#VALUE!", "ERFC.PRECISE requires 1 argument"},
		"=ERFC.PRECISE(\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// GESTEP
		"=GESTEP()":       {"#VALUE!", "GESTEP requires at least 1 argument"},
		"=GESTEP(0,0,0)":  {"#VALUE!", "GESTEP allows at most 2 arguments"},
		"=GESTEP(\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=GESTEP(0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// HEX2BIN
		"=HEX2BIN()":        {"#VALUE!", "HEX2BIN requires at least 1 argument"},
		"=HEX2BIN(1,1,1)":   {"#VALUE!", "HEX2BIN allows at most 2 arguments"},
		"=HEX2BIN(\"X\",1)": {"#NUM!", "strconv.ParseInt: parsing \"X\": invalid syntax"},
		"=HEX2BIN(1,\"\")":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=HEX2BIN(-513,10)": {"#NUM!", "strconv.ParseInt: parsing \"-\": invalid syntax"},
		"=HEX2BIN(1,-1)":    {"#NUM!", "#NUM!"},
		"=HEX2BIN(2,1)":     {"#NUM!", "#NUM!"},
		// HEX2DEC
		"=HEX2DEC()":      {"#VALUE!", "HEX2DEC requires 1 numeric argument"},
		"=HEX2DEC(\"X\")": {"#NUM!", "strconv.ParseInt: parsing \"X\": invalid syntax"},
		// HEX2OCT
		"=HEX2OCT()":        {"#VALUE!", "HEX2OCT requires at least 1 argument"},
		"=HEX2OCT(1,1,1)":   {"#VALUE!", "HEX2OCT allows at most 2 arguments"},
		"=HEX2OCT(\"X\",1)": {"#NUM!", "strconv.ParseInt: parsing \"X\": invalid syntax"},
		"=HEX2OCT(1,\"\")":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=HEX2OCT(-513,10)": {"#NUM!", "strconv.ParseInt: parsing \"-\": invalid syntax"},
		"=HEX2OCT(1,-1)":    {"#NUM!", "#NUM!"},
		// IMABS
		"=IMABS()":     {"#VALUE!", "IMABS requires 1 argument"},
		"=IMABS(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		// IMAGINARY
		"=IMAGINARY()":     {"#VALUE!", "IMAGINARY requires 1 argument"},
		"=IMAGINARY(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		// IMARGUMENT
		"=IMARGUMENT()":     {"#VALUE!", "IMARGUMENT requires 1 argument"},
		"=IMARGUMENT(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		// IMCONJUGATE
		"=IMCONJUGATE()":     {"#VALUE!", "IMCONJUGATE requires 1 argument"},
		"=IMCONJUGATE(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		// IMCOS
		"=IMCOS()":     {"#VALUE!", "IMCOS requires 1 argument"},
		"=IMCOS(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		// IMCOSH
		"=IMCOSH()":     {"#VALUE!", "IMCOSH requires 1 argument"},
		"=IMCOSH(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		// IMCOT
		"=IMCOT()":     {"#VALUE!", "IMCOT requires 1 argument"},
		"=IMCOT(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		// IMCSC
		"=IMCSC()":     {"#VALUE!", "IMCSC requires 1 argument"},
		"=IMCSC(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		"=IMCSC(0)":    {"#NUM!", "#NUM!"},
		// IMCSCH
		"=IMCSCH()":     {"#VALUE!", "IMCSCH requires 1 argument"},
		"=IMCSCH(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		"=IMCSCH(0)":    {"#NUM!", "#NUM!"},
		// IMDIV
		"=IMDIV()":       {"#VALUE!", "IMDIV requires 2 arguments"},
		"=IMDIV(0,\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		"=IMDIV(\"\",0)": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		"=IMDIV(1,0)":    {"#NUM!", "#NUM!"},
		// IMEXP
		"=IMEXP()":     {"#VALUE!", "IMEXP requires 1 argument"},
		"=IMEXP(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		// IMLN
		"=IMLN()":     {"#VALUE!", "IMLN requires 1 argument"},
		"=IMLN(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		"=IMLN(0)":    {"#NUM!", "#NUM!"},
		// IMLOG10
		"=IMLOG10()":     {"#VALUE!", "IMLOG10 requires 1 argument"},
		"=IMLOG10(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		"=IMLOG10(0)":    {"#NUM!", "#NUM!"},
		// IMLOG2
		"=IMLOG2()":     {"#VALUE!", "IMLOG2 requires 1 argument"},
		"=IMLOG2(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		"=IMLOG2(0)":    {"#NUM!", "#NUM!"},
		// IMPOWER
		"=IMPOWER()":       {"#VALUE!", "IMPOWER requires 2 arguments"},
		"=IMPOWER(0,\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		"=IMPOWER(\"\",0)": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		"=IMPOWER(0,0)":    {"#NUM!", "#NUM!"},
		"=IMPOWER(0,-1)":   {"#NUM!", "#NUM!"},
		// IMPRODUCT
		"=IMPRODUCT(\"x\")": {"#NUM!", "strconv.ParseComplex: parsing \"x\": invalid syntax"},
		"=IMPRODUCT(A1:D1)": {"#NUM!", "strconv.ParseComplex: parsing \"Month\": invalid syntax"},
		// IMREAL
		"=IMREAL()":     {"#VALUE!", "IMREAL requires 1 argument"},
		"=IMREAL(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		// IMSEC
		"=IMSEC()":     {"#VALUE!", "IMSEC requires 1 argument"},
		"=IMSEC(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		// IMSECH
		"=IMSECH()":     {"#VALUE!", "IMSECH requires 1 argument"},
		"=IMSECH(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		// IMSIN
		"=IMSIN()":     {"#VALUE!", "IMSIN requires 1 argument"},
		"=IMSIN(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		// IMSINH
		"=IMSINH()":     {"#VALUE!", "IMSINH requires 1 argument"},
		"=IMSINH(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		// IMSQRT
		"=IMSQRT()":     {"#VALUE!", "IMSQRT requires 1 argument"},
		"=IMSQRT(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		// IMSUB
		"=IMSUB()":       {"#VALUE!", "IMSUB requires 2 arguments"},
		"=IMSUB(0,\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		"=IMSUB(\"\",0)": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		// IMSUM
		"=IMSUM()":     {"#VALUE!", "IMSUM requires at least 1 argument"},
		"=IMSUM(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		// IMTAN
		"=IMTAN()":     {"#VALUE!", "IMTAN requires 1 argument"},
		"=IMTAN(\"\")": {"#NUM!", "strconv.ParseComplex: parsing \"\": invalid syntax"},
		// OCT2BIN
		"=OCT2BIN()":               {"#VALUE!", "OCT2BIN requires at least 1 argument"},
		"=OCT2BIN(1,1,1)":          {"#VALUE!", "OCT2BIN allows at most 2 arguments"},
		"=OCT2BIN(\"\",1)":         {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=OCT2BIN(1,\"\")":         {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=OCT2BIN(-536870912 ,10)": {"#NUM!", "#NUM!"},
		"=OCT2BIN(1,-1)":           {"#NUM!", "#NUM!"},
		// OCT2DEC
		"=OCT2DEC()":     {"#VALUE!", "OCT2DEC requires 1 numeric argument"},
		"=OCT2DEC(\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// OCT2HEX
		"=OCT2HEX()":               {"#VALUE!", "OCT2HEX requires at least 1 argument"},
		"=OCT2HEX(1,1,1)":          {"#VALUE!", "OCT2HEX allows at most 2 arguments"},
		"=OCT2HEX(\"\",1)":         {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=OCT2HEX(1,\"\")":         {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=OCT2HEX(-536870912 ,10)": {"#NUM!", "#NUM!"},
		"=OCT2HEX(1,-1)":           {"#NUM!", "#NUM!"},
		// Math and Trigonometric Functions
		// ABS
		"=ABS()":      {"#VALUE!", "ABS requires 1 numeric argument"},
		"=ABS(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=ABS(~)":     {"#NAME?", "invalid reference"},
		// ACOS
		"=ACOS()":        {"#VALUE!", "ACOS requires 1 numeric argument"},
		"=ACOS(\"X\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=ACOS(ACOS(0))": {"#NUM!", "#NUM!"},
		// ACOSH
		"=ACOSH()":      {"#VALUE!", "ACOSH requires 1 numeric argument"},
		"=ACOSH(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// _xlfn.ACOT
		"=_xlfn.ACOT()":      {"#VALUE!", "ACOT requires 1 numeric argument"},
		"=_xlfn.ACOT(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// _xlfn.ACOTH
		"=_xlfn.ACOTH()":               {"#VALUE!", "ACOTH requires 1 numeric argument"},
		"=_xlfn.ACOTH(\"X\")":          {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=_xlfn.ACOTH(_xlfn.ACOTH(2))": {"#NUM!", "#NUM!"},
		// _xlfn.AGGREGATE
		"=_xlfn.AGGREGATE()":             {"#VALUE!", "AGGREGATE requires at least 3 arguments"},
		"=_xlfn.AGGREGATE(\"\",0,A4:A5)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=_xlfn.AGGREGATE(1,\"\",A4:A5)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=_xlfn.AGGREGATE(0,A4:A5)":      {"#VALUE!", "AGGREGATE has invalid function_num"},
		"=_xlfn.AGGREGATE(1,8,A4:A5)":    {"#VALUE!", "AGGREGATE has invalid options"},
		"=_xlfn.AGGREGATE(1,0,A5:A6)":    {"#DIV/0!", "#DIV/0!"},
		"=_xlfn.AGGREGATE(13,0,A1:A6)":   {"#N/A", "#N/A"},
		"=_xlfn.AGGREGATE(18,0,A1:A6,1)": {"#NUM!", "#NUM!"},
		// _xlfn.ARABIC
		"=_xlfn.ARABIC()": {"#VALUE!", "ARABIC requires 1 numeric argument"},
		"=_xlfn.ARABIC(\"" + strings.Repeat("I", 256) + "\")": {"#VALUE!", "#VALUE!"},
		// ASIN
		"=ASIN()":      {"#VALUE!", "ASIN requires 1 numeric argument"},
		"=ASIN(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// ASINH
		"=ASINH()":      {"#VALUE!", "ASINH requires 1 numeric argument"},
		"=ASINH(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// ATAN
		"=ATAN()":      {"#VALUE!", "ATAN requires 1 numeric argument"},
		"=ATAN(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// ATANH
		"=ATANH()":      {"#VALUE!", "ATANH requires 1 numeric argument"},
		"=ATANH(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// ATAN2
		"=ATAN2()":        {"#VALUE!", "ATAN2 requires 2 numeric arguments"},
		"=ATAN2(\"X\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=ATAN2(0,\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// BASE
		"=BASE()":          {"#VALUE!", "BASE requires at least 2 arguments"},
		"=BASE(1,2,3,4)":   {"#VALUE!", "BASE allows at most 3 arguments"},
		"=BASE(1,1)":       {"#VALUE!", "radix must be an integer >= 2 and <= 36"},
		"=BASE(\"X\",2)":   {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=BASE(1,\"X\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=BASE(1,2,\"X\")": {"#VALUE!", "strconv.Atoi: parsing \"X\": invalid syntax"},
		// CEILING
		"=CEILING()":        {"#VALUE!", "CEILING requires at least 1 argument"},
		"=CEILING(1,2,3)":   {"#VALUE!", "CEILING allows at most 2 arguments"},
		"=CEILING(1,-1)":    {"#VALUE!", "negative sig to CEILING invalid"},
		"=CEILING(\"X\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=CEILING(0,\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// _xlfn.CEILING.MATH
		"=_xlfn.CEILING.MATH()":          {"#VALUE!", "CEILING.MATH requires at least 1 argument"},
		"=_xlfn.CEILING.MATH(1,2,3,4)":   {"#VALUE!", "CEILING.MATH allows at most 3 arguments"},
		"=_xlfn.CEILING.MATH(\"X\")":     {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=_xlfn.CEILING.MATH(1,\"X\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=_xlfn.CEILING.MATH(1,2,\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// _xlfn.CEILING.PRECISE
		"=_xlfn.CEILING.PRECISE()":        {"#VALUE!", "CEILING.PRECISE requires at least 1 argument"},
		"=_xlfn.CEILING.PRECISE(1,2,3)":   {"#VALUE!", "CEILING.PRECISE allows at most 2 arguments"},
		"=_xlfn.CEILING.PRECISE(\"X\",2)": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=_xlfn.CEILING.PRECISE(1,\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// COMBIN
		"=COMBIN()":         {"#VALUE!", "COMBIN requires 2 argument"},
		"=COMBIN(-1,1)":     {"#VALUE!", "COMBIN requires number >= number_chosen"},
		"=COMBIN(\"X\",1)":  {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=COMBIN(-1,\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// _xlfn.COMBINA
		"=_xlfn.COMBINA()":         {"#VALUE!", "COMBINA requires 2 argument"},
		"=_xlfn.COMBINA(-1,1)":     {"#VALUE!", "COMBINA requires number > number_chosen"},
		"=_xlfn.COMBINA(-1,-1)":    {"#VALUE!", "COMBIN requires number >= number_chosen"},
		"=_xlfn.COMBINA(\"X\",1)":  {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=_xlfn.COMBINA(-1,\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// COS
		"=COS()":      {"#VALUE!", "COS requires 1 numeric argument"},
		"=COS(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// COSH
		"=COSH()":      {"#VALUE!", "COSH requires 1 numeric argument"},
		"=COSH(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// _xlfn.COT
		"=COT()":      {"#VALUE!", "COT requires 1 numeric argument"},
		"=COT(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=COT(0)":     {"#DIV/0!", "#DIV/0!"},
		// _xlfn.COTH
		"=COTH()":      {"#VALUE!", "COTH requires 1 numeric argument"},
		"=COTH(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=COTH(0)":     {"#DIV/0!", "#DIV/0!"},
		// _xlfn.CSC
		"=_xlfn.CSC()":      {"#VALUE!", "CSC requires 1 numeric argument"},
		"=_xlfn.CSC(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=_xlfn.CSC(0)":     {"#DIV/0!", "#DIV/0!"},
		// _xlfn.CSCH
		"=_xlfn.CSCH()":      {"#VALUE!", "CSCH requires 1 numeric argument"},
		"=_xlfn.CSCH(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=_xlfn.CSCH(0)":     {"#DIV/0!", "#DIV/0!"},
		// _xlfn.DECIMAL
		"=_xlfn.DECIMAL()":           {"#VALUE!", "DECIMAL requires 2 numeric arguments"},
		"=_xlfn.DECIMAL(\"X\",2)":    {"#VALUE!", "strconv.ParseInt: parsing \"X\": invalid syntax"},
		"=_xlfn.DECIMAL(2000,\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// DEGREES
		"=DEGREES()":      {"#VALUE!", "DEGREES requires 1 numeric argument"},
		"=DEGREES(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=DEGREES(0)":     {"#DIV/0!", "#DIV/0!"},
		// EVEN
		"=EVEN()":      {"#VALUE!", "EVEN requires 1 numeric argument"},
		"=EVEN(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// EXP
		"=EXP()":      {"#VALUE!", "EXP requires 1 numeric argument"},
		"=EXP(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// FACT
		"=FACT()":      {"#VALUE!", "FACT requires 1 numeric argument"},
		"=FACT(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=FACT(-1)":    {"#NUM!", "#NUM!"},
		// FACTDOUBLE
		"=FACTDOUBLE()":      {"#VALUE!", "FACTDOUBLE requires 1 numeric argument"},
		"=FACTDOUBLE(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=FACTDOUBLE(-1)":    {"#NUM!", "#NUM!"},
		// FLOOR
		"=FLOOR()":         {"#VALUE!", "FLOOR requires 2 numeric arguments"},
		"=FLOOR(\"X\",-1)": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=FLOOR(1,\"X\")":  {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=FLOOR(1,-1)":     {"#NUM!", "invalid arguments to FLOOR"},
		// _xlfn.FLOOR.MATH
		"=_xlfn.FLOOR.MATH()":          {"#VALUE!", "FLOOR.MATH requires at least 1 argument"},
		"=_xlfn.FLOOR.MATH(1,2,3,4)":   {"#VALUE!", "FLOOR.MATH allows at most 3 arguments"},
		"=_xlfn.FLOOR.MATH(\"X\",2,3)": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=_xlfn.FLOOR.MATH(1,\"X\",3)": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=_xlfn.FLOOR.MATH(1,2,\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// _xlfn.FLOOR.PRECISE
		"=_xlfn.FLOOR.PRECISE()":        {"#VALUE!", "FLOOR.PRECISE requires at least 1 argument"},
		"=_xlfn.FLOOR.PRECISE(1,2,3)":   {"#VALUE!", "FLOOR.PRECISE allows at most 2 arguments"},
		"=_xlfn.FLOOR.PRECISE(\"X\",2)": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=_xlfn.FLOOR.PRECISE(1,\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// GCD
		"=GCD()":      {"#VALUE!", "GCD requires at least 1 argument"},
		"=GCD(\"\")":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=GCD(-1)":    {"#VALUE!", "GCD only accepts positive arguments"},
		"=GCD(1,-1)":  {"#VALUE!", "GCD only accepts positive arguments"},
		"=GCD(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// INT
		"=INT()":      {"#VALUE!", "INT requires 1 numeric argument"},
		"=INT(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// ISO.CEILING
		"=ISO.CEILING()":        {"#VALUE!", "ISO.CEILING requires at least 1 argument"},
		"=ISO.CEILING(1,2,3)":   {"#VALUE!", "ISO.CEILING allows at most 2 arguments"},
		"=ISO.CEILING(\"X\",2)": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=ISO.CEILING(1,\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// LCM
		"=LCM()":      {"#VALUE!", "LCM requires at least 1 argument"},
		"=LCM(-1)":    {"#VALUE!", "LCM only accepts positive arguments"},
		"=LCM(1,-1)":  {"#VALUE!", "LCM only accepts positive arguments"},
		"=LCM(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// LN
		"=LN()":      {"#VALUE!", "LN requires 1 numeric argument"},
		"=LN(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// LOG
		"=LOG()":        {"#VALUE!", "LOG requires at least 1 argument"},
		"=LOG(1,2,3)":   {"#VALUE!", "LOG allows at most 2 arguments"},
		"=LOG(\"X\",1)": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=LOG(1,\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=LOG(0,0)":     {"#NUM!", "#DIV/0!"},
		"=LOG(1,0)":     {"#NUM!", "#DIV/0!"},
		"=LOG(1,1)":     {"#DIV/0!", "#DIV/0!"},
		// LOG10
		"=LOG10()":      {"#VALUE!", "LOG10 requires 1 numeric argument"},
		"=LOG10(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// MDETERM
		"=MDETERM()": {"#VALUE!", "MDETERM requires 1 argument"},
		// MINVERSE
		"=MINVERSE()":      {"#VALUE!", "MINVERSE requires 1 argument"},
		"=MINVERSE(B3:C4)": {"#VALUE!", "#VALUE!"},
		"=MINVERSE(A1:C2)": {"#VALUE!", "#VALUE!"},
		"=MINVERSE(A4:A4)": {"#NUM!", "#NUM!"},
		// MMULT
		"=MMULT()":            {"#VALUE!", "MMULT requires 2 argument"},
		"=MMULT(A1:B2,B3:C4)": {"#VALUE!", "#VALUE!"},
		"=MMULT(B3:C4,A1:B2)": {"#VALUE!", "#VALUE!"},
		"=MMULT(A1:A2,B1:B2)": {"#VALUE!", "#VALUE!"},
		// MOD
		"=MOD()":        {"#VALUE!", "MOD requires 2 numeric arguments"},
		"=MOD(6,0)":     {"#DIV/0!", "MOD divide by zero"},
		"=MOD(\"X\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=MOD(6,\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// MROUND
		"=MROUND()":        {"#VALUE!", "MROUND requires 2 numeric arguments"},
		"=MROUND(1,0)":     {"#NUM!", "#NUM!"},
		"=MROUND(1,-1)":    {"#NUM!", "#NUM!"},
		"=MROUND(\"X\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=MROUND(1,\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// MULTINOMIAL
		"=MULTINOMIAL(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// _xlfn.MUNIT
		"=_xlfn.MUNIT()":      {"#VALUE!", "MUNIT requires 1 numeric argument"},
		"=_xlfn.MUNIT(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=_xlfn.MUNIT(-1)":    {"#VALUE!", ""},
		// ODD
		"=ODD()":      {"#VALUE!", "ODD requires 1 numeric argument"},
		"=ODD(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// PI
		"=PI(1)": {"#VALUE!", "PI accepts no arguments"},
		// POWER
		"=POWER(\"X\",1)": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=POWER(1,\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=POWER(0,0)":     {"#NUM!", "#NUM!"},
		"=POWER(0,-1)":    {"#DIV/0!", "#DIV/0!"},
		"=POWER(1)":       {"#VALUE!", "POWER requires 2 numeric arguments"},
		// PRODUCT
		"=PRODUCT(\"X\")":    {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=PRODUCT(\"\",3,6)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// QUOTIENT
		"=QUOTIENT(\"X\",1)": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=QUOTIENT(1,\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=QUOTIENT(1,0)":     {"#DIV/0!", "#DIV/0!"},
		"=QUOTIENT(1)":       {"#VALUE!", "QUOTIENT requires 2 numeric arguments"},
		// RADIANS
		"=RADIANS(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=RADIANS()":      {"#VALUE!", "RADIANS requires 1 numeric argument"},
		// RAND
		"=RAND(1)": {"#VALUE!", "RAND accepts no arguments"},
		// RANDBETWEEN
		"=RANDBETWEEN(\"X\",1)": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=RANDBETWEEN(1,\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=RANDBETWEEN()":        {"#VALUE!", "RANDBETWEEN requires 2 numeric arguments"},
		"=RANDBETWEEN(2,1)":     {"#NUM!", "#NUM!"},
		// ROMAN
		"=ROMAN()":       {"#VALUE!", "ROMAN requires at least 1 argument"},
		"=ROMAN(1,2,3)":  {"#VALUE!", "ROMAN allows at most 2 arguments"},
		"=ROMAN(1,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ROMAN(\"\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ROMAN(\"\",1)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// ROUND
		"=ROUND()":        {"#VALUE!", "ROUND requires 2 numeric arguments"},
		"=ROUND(\"X\",1)": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=ROUND(1,\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// ROUNDDOWN
		"=ROUNDDOWN()":        {"#VALUE!", "ROUNDDOWN requires 2 numeric arguments"},
		"=ROUNDDOWN(\"X\",1)": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=ROUNDDOWN(1,\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// ROUNDUP
		"=ROUNDUP()":        {"#VALUE!", "ROUNDUP requires 2 numeric arguments"},
		"=ROUNDUP(\"X\",1)": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=ROUNDUP(1,\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// SEARCH
		"=SEARCH()":          {"#VALUE!", "SEARCH requires at least 2 arguments"},
		"=SEARCH(1,A1,1,1)":  {"#VALUE!", "SEARCH allows at most 3 arguments"},
		"=SEARCH(2,A1)":      {"#VALUE!", "#VALUE!"},
		"=SEARCH(1,A1,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// SEARCHB
		"=SEARCHB()":                   {"#VALUE!", "SEARCHB requires at least 2 arguments"},
		"=SEARCHB(1,A1,1,1)":           {"#VALUE!", "SEARCHB allows at most 3 arguments"},
		"=SEARCHB(2,A1)":               {"#VALUE!", "#VALUE!"},
		"=SEARCHB(\"?w\",\"你好world\")": {"#VALUE!", "#VALUE!"},
		"=SEARCHB(1,A1,\"\")":          {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// SEC
		"=_xlfn.SEC()":      {"#VALUE!", "SEC requires 1 numeric argument"},
		"=_xlfn.SEC(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// _xlfn.SECH
		"=_xlfn.SECH()":      {"#VALUE!", "SECH requires 1 numeric argument"},
		"=_xlfn.SECH(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// SERIESSUM
		"=SERIESSUM()":               {"#VALUE!", "SERIESSUM requires 4 arguments"},
		"=SERIESSUM(\"\",2,3,A1:A4)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=SERIESSUM(1,\"\",3,A1:A4)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=SERIESSUM(1,2,\"\",A1:A4)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=SERIESSUM(1,2,3,A1:D1)":    {"#VALUE!", "strconv.ParseFloat: parsing \"Month\": invalid syntax"},
		// SIGN
		"=SIGN()":      {"#VALUE!", "SIGN requires 1 numeric argument"},
		"=SIGN(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// SIN
		"=SIN()":      {"#VALUE!", "SIN requires 1 numeric argument"},
		"=SIN(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// SINH
		"=SINH()":      {"#VALUE!", "SINH requires 1 numeric argument"},
		"=SINH(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// SQRT
		"=SQRT()":      {"#VALUE!", "SQRT requires 1 numeric argument"},
		"=SQRT(\"\")":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=SQRT(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=SQRT(-1)":    {"#NUM!", "#NUM!"},
		// SQRTPI
		"=SQRTPI()":      {"#VALUE!", "SQRTPI requires 1 numeric argument"},
		"=SQRTPI(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// STDEV
		"=STDEV()":      {"#VALUE!", "STDEV requires at least 1 argument"},
		"=STDEV(E2:E9)": {"#DIV/0!", "#DIV/0!"},
		// STDEV.S
		"=STDEV.S()": {"#VALUE!", "STDEV.S requires at least 1 argument"},
		// STDEVA
		"=STDEVA()":      {"#VALUE!", "STDEVA requires at least 1 argument"},
		"=STDEVA(E2:E9)": {"#DIV/0!", "#DIV/0!"},
		// POISSON.DIST
		"=POISSON.DIST()": {"#VALUE!", "POISSON.DIST requires 3 arguments"},
		// POISSON
		"=POISSON()":             {"#VALUE!", "POISSON requires 3 arguments"},
		"=POISSON(\"\",0,FALSE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=POISSON(0,\"\",FALSE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=POISSON(0,0,\"\")":     {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=POISSON(0,-1,TRUE)":    {"#N/A", "#N/A"},
		// PROB
		"=PROB()":                   {"#VALUE!", "PROB requires at least 3 arguments"},
		"=PROB(A1:A2,B1:B2,1,1,1)":  {"#VALUE!", "PROB requires at most 4 arguments"},
		"=PROB(A1:A2,B1:B2,\"\")":   {"#VALUE!", "#VALUE!"},
		"=PROB(A1:A2,B1:B2,1,\"\")": {"#VALUE!", "#VALUE!"},
		"=PROB(A1,B1,1)":            {"#NUM!", "#NUM!"},
		"=PROB(A1:A2,B1:B3,1)":      {"#N/A", "#N/A"},
		"=PROB(A1:A2,B1:C2,1)":      {"#N/A", "#N/A"},
		"=PROB(A1:A2,B1:B2,1)":      {"#NUM!", "#NUM!"},
		// SUBTOTAL
		"=SUBTOTAL()":           {"#VALUE!", "SUBTOTAL requires at least 2 arguments"},
		"=SUBTOTAL(\"\",A4:A5)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=SUBTOTAL(0,A4:A5)":    {"#VALUE!", "SUBTOTAL has invalid function_num"},
		"=SUBTOTAL(1,A5:A6)":    {"#DIV/0!", "#DIV/0!"},
		// SUM
		"=SUM((":             {"", ErrInvalidFormula.Error()},
		"=SUM(-)":            {ErrInvalidFormula.Error(), ErrInvalidFormula.Error()},
		"=SUM(1+)":           {ErrInvalidFormula.Error(), ErrInvalidFormula.Error()},
		"=SUM(1-)":           {ErrInvalidFormula.Error(), ErrInvalidFormula.Error()},
		"=SUM(1*)":           {ErrInvalidFormula.Error(), ErrInvalidFormula.Error()},
		"=SUM(1/)":           {ErrInvalidFormula.Error(), ErrInvalidFormula.Error()},
		"=SUM(1*SUM(1/0))":   {"#DIV/0!", "#DIV/0!"},
		"=SUM(1*SUM(1/0)*1)": {"", "#DIV/0!"},
		// SUMIF
		"=SUMIF()": {"#VALUE!", "SUMIF requires at least 2 arguments"},
		// SUMSQ
		"=SUMSQ(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=SUMSQ(C1:D2)": {"#VALUE!", "strconv.ParseFloat: parsing \"Month\": invalid syntax"},
		// SUMPRODUCT
		"=SUMPRODUCT()":            {"#VALUE!", "SUMPRODUCT requires at least 1 argument"},
		"=SUMPRODUCT(A1,B1:B2)":    {"#VALUE!", "#VALUE!"},
		"=SUMPRODUCT(A1,D1)":       {"#VALUE!", "#VALUE!"},
		"=SUMPRODUCT(A1:A3,D1:D3)": {"#VALUE!", "#VALUE!"},
		"=SUMPRODUCT(A1:A2,B1:B3)": {"#VALUE!", "#VALUE!"},
		"=SUMPRODUCT(\"\")":        {"#VALUE!", "#VALUE!"},
		"=SUMPRODUCT(A1,NA())":     {"#N/A", "#N/A"},
		// SUMX2MY2
		"=SUMX2MY2()":         {"#VALUE!", "SUMX2MY2 requires 2 arguments"},
		"=SUMX2MY2(A1,B1:B2)": {"#N/A", "#N/A"},
		// SUMX2PY2
		"=SUMX2PY2()":         {"#VALUE!", "SUMX2PY2 requires 2 arguments"},
		"=SUMX2PY2(A1,B1:B2)": {"#N/A", "#N/A"},
		// SUMXMY2
		"=SUMXMY2()":         {"#VALUE!", "SUMXMY2 requires 2 arguments"},
		"=SUMXMY2(A1,B1:B2)": {"#N/A", "#N/A"},
		// TAN
		"=TAN()":      {"#VALUE!", "TAN requires 1 numeric argument"},
		"=TAN(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// TANH
		"=TANH()":      {"#VALUE!", "TANH requires 1 numeric argument"},
		"=TANH(\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// TRUNC
		"=TRUNC()":        {"#VALUE!", "TRUNC requires at least 1 argument"},
		"=TRUNC(\"X\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		"=TRUNC(1,\"X\")": {"#VALUE!", "strconv.ParseFloat: parsing \"X\": invalid syntax"},
		// Statistical Functions
		// AVEDEV
		"=AVEDEV()":       {"#VALUE!", "AVEDEV requires at least 1 argument"},
		"=AVEDEV(\"\")":   {"#VALUE!", "#VALUE!"},
		"=AVEDEV(1,\"\")": {"#VALUE!", "#VALUE!"},
		// AVERAGE
		"=AVERAGE(H1)": {"#DIV/0!", "#DIV/0!"},
		// AVERAGEA
		"=AVERAGEA(H1)": {"#DIV/0!", "#DIV/0!"},
		// AVERAGEIF
		"=AVERAGEIF()":                      {"#VALUE!", "AVERAGEIF requires at least 2 arguments"},
		"=AVERAGEIF(H1,\"\")":               {"#DIV/0!", "#DIV/0!"},
		"=AVERAGEIF(D1:D3,\"Month\",D1:D3)": {"#DIV/0!", "#DIV/0!"},
		"=AVERAGEIF(C1:C3,\"Month\",D1:D3)": {"#DIV/0!", "#DIV/0!"},
		// BETA.DIST
		"=BETA.DIST()":                     {"#VALUE!", "BETA.DIST requires at least 4 arguments"},
		"=BETA.DIST(0.4,4,5,TRUE,0,1,0)":   {"#VALUE!", "BETA.DIST requires at most 6 arguments"},
		"=BETA.DIST(\"\",4,5,TRUE,0,1)":    {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BETA.DIST(0.4,\"\",5,TRUE,0,1)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BETA.DIST(0.4,4,\"\",TRUE,0,1)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BETA.DIST(0.4,4,5,\"\",0,1)":     {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=BETA.DIST(0.4,4,5,TRUE,\"\",1)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BETA.DIST(0.4,4,5,TRUE,0,\"\")":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BETA.DIST(0.4,0,5,TRUE,0,1)":     {"#NUM!", "#NUM!"},
		"=BETA.DIST(0.4,4,0,TRUE,0,0)":     {"#NUM!", "#NUM!"},
		"=BETA.DIST(0.4,4,5,TRUE,0.5,1)":   {"#NUM!", "#NUM!"},
		"=BETA.DIST(0.4,4,5,TRUE,0,0.3)":   {"#NUM!", "#NUM!"},
		"=BETA.DIST(0.4,4,5,TRUE,0.4,0.4)": {"#NUM!", "#NUM!"},
		// BETADIST
		"=BETADIST()":                {"#VALUE!", "BETADIST requires at least 3 arguments"},
		"=BETADIST(0.4,4,5,0,1,0)":   {"#VALUE!", "BETADIST requires at most 5 arguments"},
		"=BETADIST(\"\",4,5,0,1)":    {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BETADIST(0.4,\"\",5,0,1)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BETADIST(0.4,4,\"\",0,1)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BETADIST(0.4,4,5,\"\",1)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BETADIST(0.4,4,5,0,\"\")":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BETADIST(2,4,5,3,1)":       {"#NUM!", "#NUM!"},
		"=BETADIST(2,4,5,0,1)":       {"#NUM!", "#NUM!"},
		"=BETADIST(0.4,0,5,0,1)":     {"#NUM!", "#NUM!"},
		"=BETADIST(0.4,4,0,0,1)":     {"#NUM!", "#NUM!"},
		"=BETADIST(0.4,4,5,0.4,0.4)": {"#NUM!", "#NUM!"},
		// BETAINV
		"=BETAINV()":               {"#VALUE!", "BETAINV requires at least 3 arguments"},
		"=BETAINV(0.2,4,5,0,1,0)":  {"#VALUE!", "BETAINV requires at most 5 arguments"},
		"=BETAINV(\"\",4,5,0,1)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BETAINV(0.2,\"\",5,0,1)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BETAINV(0.2,4,\"\",0,1)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BETAINV(0.2,4,5,\"\",1)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BETAINV(0.2,4,5,0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BETAINV(0,4,5,0,1)":      {"#NUM!", "#NUM!"},
		"=BETAINV(1,4,5,0,1)":      {"#NUM!", "#NUM!"},
		"=BETAINV(0.2,0,5,0,1)":    {"#NUM!", "#NUM!"},
		"=BETAINV(0.2,4,0,0,1)":    {"#NUM!", "#NUM!"},
		"=BETAINV(0.2,4,5,2,2)":    {"#NUM!", "#NUM!"},
		// BETA.INV
		"=BETA.INV()":               {"#VALUE!", "BETA.INV requires at least 3 arguments"},
		"=BETA.INV(0.2,4,5,0,1,0)":  {"#VALUE!", "BETA.INV requires at most 5 arguments"},
		"=BETA.INV(\"\",4,5,0,1)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BETA.INV(0.2,\"\",5,0,1)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BETA.INV(0.2,4,\"\",0,1)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BETA.INV(0.2,4,5,\"\",1)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BETA.INV(0.2,4,5,0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BETA.INV(0,4,5,0,1)":      {"#NUM!", "#NUM!"},
		"=BETA.INV(1,4,5,0,1)":      {"#NUM!", "#NUM!"},
		"=BETA.INV(0.2,0,5,0,1)":    {"#NUM!", "#NUM!"},
		"=BETA.INV(0.2,4,0,0,1)":    {"#NUM!", "#NUM!"},
		"=BETA.INV(0.2,4,5,2,2)":    {"#NUM!", "#NUM!"},
		// BINOMDIST
		"=BINOMDIST()":                   {"#VALUE!", "BINOMDIST requires 4 arguments"},
		"=BINOMDIST(\"\",100,0.5,FALSE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BINOMDIST(10,\"\",0.5,FALSE)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BINOMDIST(10,100,\"\",FALSE)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BINOMDIST(10,100,0.5,\"\")":    {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=BINOMDIST(-1,100,0.5,FALSE)":   {"#NUM!", "#NUM!"},
		"=BINOMDIST(110,100,0.5,FALSE)":  {"#NUM!", "#NUM!"},
		"=BINOMDIST(10,100,-1,FALSE)":    {"#NUM!", "#NUM!"},
		"=BINOMDIST(10,100,2,FALSE)":     {"#NUM!", "#NUM!"},
		// BINOM.DIST
		"=BINOM.DIST()":                   {"#VALUE!", "BINOM.DIST requires 4 arguments"},
		"=BINOM.DIST(\"\",100,0.5,FALSE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BINOM.DIST(10,\"\",0.5,FALSE)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BINOM.DIST(10,100,\"\",FALSE)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BINOM.DIST(10,100,0.5,\"\")":    {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=BINOM.DIST(-1,100,0.5,FALSE)":   {"#NUM!", "#NUM!"},
		"=BINOM.DIST(110,100,0.5,FALSE)":  {"#NUM!", "#NUM!"},
		"=BINOM.DIST(10,100,-1,FALSE)":    {"#NUM!", "#NUM!"},
		"=BINOM.DIST(10,100,2,FALSE)":     {"#NUM!", "#NUM!"},
		// BINOM.DIST.RANGE
		"=BINOM.DIST.RANGE()":                {"#VALUE!", "BINOM.DIST.RANGE requires at least 3 arguments"},
		"=BINOM.DIST.RANGE(100,0.5,0,40,0)":  {"#VALUE!", "BINOM.DIST.RANGE requires at most 4 arguments"},
		"=BINOM.DIST.RANGE(\"\",0.5,0,40)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BINOM.DIST.RANGE(100,\"\",0,40)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BINOM.DIST.RANGE(100,0.5,\"\",40)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BINOM.DIST.RANGE(100,0.5,0,\"\")":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BINOM.DIST.RANGE(100,-1,0,40)":     {"#NUM!", "#NUM!"},
		"=BINOM.DIST.RANGE(100,2,0,40)":      {"#NUM!", "#NUM!"},
		"=BINOM.DIST.RANGE(100,0.5,-1,40)":   {"#NUM!", "#NUM!"},
		"=BINOM.DIST.RANGE(100,0.5,110,40)":  {"#NUM!", "#NUM!"},
		"=BINOM.DIST.RANGE(100,0.5,0,-1)":    {"#NUM!", "#NUM!"},
		"=BINOM.DIST.RANGE(100,0.5,0,110)":   {"#NUM!", "#NUM!"},
		// BINOM.INV
		"=BINOM.INV()":             {"#VALUE!", "BINOM.INV requires 3 numeric arguments"},
		"=BINOM.INV(\"\",0.5,20%)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BINOM.INV(100,\"\",20%)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BINOM.INV(100,0.5,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=BINOM.INV(-1,0.5,20%)":   {"#NUM!", "#NUM!"},
		"=BINOM.INV(100,-1,20%)":   {"#NUM!", "#NUM!"},
		"=BINOM.INV(100,2,20%)":    {"#NUM!", "#NUM!"},
		"=BINOM.INV(100,0.5,-1)":   {"#NUM!", "#NUM!"},
		"=BINOM.INV(100,0.5,2)":    {"#NUM!", "#NUM!"},
		"=BINOM.INV(1,1,20%)":      {"#NUM!", "#NUM!"},
		// CHIDIST
		"=CHIDIST()":         {"#VALUE!", "CHIDIST requires 2 numeric arguments"},
		"=CHIDIST(\"\",3)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CHIDIST(0.5,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// CHIINV
		"=CHIINV()":         {"#VALUE!", "CHIINV requires 2 numeric arguments"},
		"=CHIINV(\"\",1)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CHIINV(0.5,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CHIINV(0,1)":      {"#NUM!", "#NUM!"},
		"=CHIINV(2,1)":      {"#NUM!", "#NUM!"},
		"=CHIINV(0.5,0.5)":  {"#NUM!", "#NUM!"},
		// CHISQ.DIST
		"=CHISQ.DIST()":            {"#VALUE!", "CHISQ.DIST requires 3 arguments"},
		"=CHISQ.DIST(\"\",2,TRUE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CHISQ.DIST(3,\"\",TRUE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CHISQ.DIST(3,2,\"\")":    {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=CHISQ.DIST(-1,2,TRUE)":   {"#NUM!", "#NUM!"},
		"=CHISQ.DIST(3,0,TRUE)":    {"#NUM!", "#NUM!"},
		// CHISQ.DIST.RT
		"=CHISQ.DIST.RT()":         {"#VALUE!", "CHISQ.DIST.RT requires 2 numeric arguments"},
		"=CHISQ.DIST.RT(\"\",3)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CHISQ.DIST.RT(0.5,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// CHISQ.INV
		"=CHISQ.INV()":                {"#VALUE!", "CHISQ.INV requires 2 numeric arguments"},
		"=CHISQ.INV(\"\",1)":          {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CHISQ.INV(0.5,\"\")":        {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CHISQ.INV(-1,1)":            {"#NUM!", "#NUM!"},
		"=CHISQ.INV(1,1)":             {"#NUM!", "#NUM!"},
		"=CHISQ.INV(0.5,0.5)":         {"#NUM!", "#NUM!"},
		"=CHISQ.INV(0.5,10000000001)": {"#NUM!", "#NUM!"},
		// CHISQ.INV.RT
		"=CHISQ.INV.RT()":         {"#VALUE!", "CHISQ.INV.RT requires 2 numeric arguments"},
		"=CHISQ.INV.RT(\"\",1)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CHISQ.INV.RT(0.5,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CHISQ.INV.RT(0,1)":      {"#NUM!", "#NUM!"},
		"=CHISQ.INV.RT(2,1)":      {"#NUM!", "#NUM!"},
		"=CHISQ.INV.RT(0.5,0.5)":  {"#NUM!", "#NUM!"},
		// CONFIDENCE
		"=CONFIDENCE()":               {"#VALUE!", "CONFIDENCE requires 3 numeric arguments"},
		"=CONFIDENCE(\"\",0.07,100)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CONFIDENCE(0.05,\"\",100)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CONFIDENCE(0.05,0.07,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CONFIDENCE(0,0.07,100)":     {"#NUM!", "#NUM!"},
		"=CONFIDENCE(1,0.07,100)":     {"#NUM!", "#NUM!"},
		"=CONFIDENCE(0.05,0,100)":     {"#NUM!", "#NUM!"},
		"=CONFIDENCE(0.05,0.07,0.5)":  {"#NUM!", "#NUM!"},
		// CONFIDENCE.NORM
		"=CONFIDENCE.NORM()":               {"#VALUE!", "CONFIDENCE.NORM requires 3 numeric arguments"},
		"=CONFIDENCE.NORM(\"\",0.07,100)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CONFIDENCE.NORM(0.05,\"\",100)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CONFIDENCE.NORM(0.05,0.07,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CONFIDENCE.NORM(0,0.07,100)":     {"#NUM!", "#NUM!"},
		"=CONFIDENCE.NORM(1,0.07,100)":     {"#NUM!", "#NUM!"},
		"=CONFIDENCE.NORM(0.05,0,100)":     {"#NUM!", "#NUM!"},
		"=CONFIDENCE.NORM(0.05,0.07,0.5)":  {"#NUM!", "#NUM!"},
		// CORREL
		"=CORREL()":            {"#VALUE!", "CORREL requires 2 arguments"},
		"=CORREL(A1:A3,B1:B5)": {"#N/A", "#N/A"},
		"=CORREL(A1:A1,B1:B1)": {"#DIV/0!", "#DIV/0!"},
		// CONFIDENCE.T
		"=CONFIDENCE.T()":               {"#VALUE!", "CONFIDENCE.T requires 3 arguments"},
		"=CONFIDENCE.T(\"\",0.07,100)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CONFIDENCE.T(0.05,\"\",100)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CONFIDENCE.T(0.05,0.07,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CONFIDENCE.T(0,0.07,100)":     {"#NUM!", "#NUM!"},
		"=CONFIDENCE.T(1,0.07,100)":     {"#NUM!", "#NUM!"},
		"=CONFIDENCE.T(0.05,0,100)":     {"#NUM!", "#NUM!"},
		"=CONFIDENCE.T(0.05,0.07,0)":    {"#NUM!", "#NUM!"},
		"=CONFIDENCE.T(0.05,0.07,1)":    {"#DIV/0!", "#DIV/0!"},
		// COUNTBLANK
		"=COUNTBLANK()":    {"#VALUE!", "COUNTBLANK requires 1 argument"},
		"=COUNTBLANK(1,2)": {"#VALUE!", "COUNTBLANK requires 1 argument"},
		// COUNTIF
		"=COUNTIF()": {"#VALUE!", "COUNTIF requires 2 arguments"},
		// COUNTIFS
		"=COUNTIFS()":              {"#VALUE!", "COUNTIFS requires at least 2 arguments"},
		"=COUNTIFS(A1:A9,2,D1:D9)": {"#N/A", "#N/A"},
		// CRITBINOM
		"=CRITBINOM()":             {"#VALUE!", "CRITBINOM requires 3 numeric arguments"},
		"=CRITBINOM(\"\",0.5,20%)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CRITBINOM(100,\"\",20%)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CRITBINOM(100,0.5,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CRITBINOM(-1,0.5,20%)":   {"#NUM!", "#NUM!"},
		"=CRITBINOM(100,-1,20%)":   {"#NUM!", "#NUM!"},
		"=CRITBINOM(100,2,20%)":    {"#NUM!", "#NUM!"},
		"=CRITBINOM(100,0.5,-1)":   {"#NUM!", "#NUM!"},
		"=CRITBINOM(100,0.5,2)":    {"#NUM!", "#NUM!"},
		"=CRITBINOM(1,1,20%)":      {"#NUM!", "#NUM!"},
		// DEVSQ
		"=DEVSQ()":      {"#VALUE!", "DEVSQ requires at least 1 numeric argument"},
		"=DEVSQ(D1:D2)": {"#N/A", "#N/A"},
		// FISHER
		"=FISHER()":         {"#VALUE!", "FISHER requires 1 numeric argument"},
		"=FISHER(2)":        {"#N/A", "#N/A"},
		"=FISHER(\"2\")":    {"#N/A", "#N/A"},
		"=FISHER(INT(-2)))": {"#N/A", "#N/A"},
		"=FISHER(F1)":       {"#VALUE!", "FISHER requires 1 numeric argument"},
		// FISHERINV
		"=FISHERINV()":   {"#VALUE!", "FISHERINV requires 1 numeric argument"},
		"=FISHERINV(F1)": {"#VALUE!", "FISHERINV requires 1 numeric argument"},
		// FORECAST
		"=FORECAST()":                 {"#VALUE!", "FORECAST requires 3 arguments"},
		"=FORECAST(\"\",A1:A7,B1:B7)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=FORECAST(1,A1:A2,B1:B1)":    {"#N/A", "#N/A"},
		"=FORECAST(1,A4,A4)":          {"#DIV/0!", "#DIV/0!"},
		// FORECAST.LINEAR
		"=FORECAST.LINEAR()":                 {"#VALUE!", "FORECAST.LINEAR requires 3 arguments"},
		"=FORECAST.LINEAR(\"\",A1:A7,B1:B7)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=FORECAST.LINEAR(1,A1:A2,B1:B1)":    {"#N/A", "#N/A"},
		"=FORECAST.LINEAR(1,A4,A4)":          {"#DIV/0!", "#DIV/0!"},
		// FREQUENCY
		"=FREQUENCY()":           {"#VALUE!", "FREQUENCY requires 2 arguments"},
		"=FREQUENCY(NA(),A1:A3)": {"#N/A", "#N/A"},
		"=FREQUENCY(A1:A3,NA())": {"#N/A", "#N/A"},
		// GAMMA
		"=GAMMA()":       {"#VALUE!", "GAMMA requires 1 numeric argument"},
		"=GAMMA(F1)":     {"#VALUE!", "GAMMA requires 1 numeric argument"},
		"=GAMMA(0)":      {"#N/A", "#N/A"},
		"=GAMMA(\"0\")":  {"#N/A", "#N/A"},
		"=GAMMA(INT(0))": {"#N/A", "#N/A"},
		// GAMMA.DIST
		"=GAMMA.DIST()":               {"#VALUE!", "GAMMA.DIST requires 4 arguments"},
		"=GAMMA.DIST(\"\",3,2,FALSE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=GAMMA.DIST(6,\"\",2,FALSE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=GAMMA.DIST(6,3,\"\",FALSE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=GAMMA.DIST(6,3,2,\"\")":     {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=GAMMA.DIST(-1,3,2,FALSE)":   {"#NUM!", "#NUM!"},
		"=GAMMA.DIST(6,0,2,FALSE)":    {"#NUM!", "#NUM!"},
		"=GAMMA.DIST(6,3,0,FALSE)":    {"#NUM!", "#NUM!"},
		// GAMMADIST
		"=GAMMADIST()":               {"#VALUE!", "GAMMADIST requires 4 arguments"},
		"=GAMMADIST(\"\",3,2,FALSE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=GAMMADIST(6,\"\",2,FALSE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=GAMMADIST(6,3,\"\",FALSE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=GAMMADIST(6,3,2,\"\")":     {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=GAMMADIST(-1,3,2,FALSE)":   {"#NUM!", "#NUM!"},
		"=GAMMADIST(6,0,2,FALSE)":    {"#NUM!", "#NUM!"},
		"=GAMMADIST(6,3,0,FALSE)":    {"#NUM!", "#NUM!"},
		// GAMMA.INV
		"=GAMMA.INV()":           {"#VALUE!", "GAMMA.INV requires 3 arguments"},
		"=GAMMA.INV(\"\",3,2)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=GAMMA.INV(0.5,\"\",2)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=GAMMA.INV(0.5,3,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=GAMMA.INV(-1,3,2)":     {"#NUM!", "#NUM!"},
		"=GAMMA.INV(2,3,2)":      {"#NUM!", "#NUM!"},
		"=GAMMA.INV(0.5,0,2)":    {"#NUM!", "#NUM!"},
		"=GAMMA.INV(0.5,3,0)":    {"#NUM!", "#NUM!"},
		// GAMMAINV
		"=GAMMAINV()":           {"#VALUE!", "GAMMAINV requires 3 arguments"},
		"=GAMMAINV(\"\",3,2)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=GAMMAINV(0.5,\"\",2)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=GAMMAINV(0.5,3,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=GAMMAINV(-1,3,2)":     {"#NUM!", "#NUM!"},
		"=GAMMAINV(2,3,2)":      {"#NUM!", "#NUM!"},
		"=GAMMAINV(0.5,0,2)":    {"#NUM!", "#NUM!"},
		"=GAMMAINV(0.5,3,0)":    {"#NUM!", "#NUM!"},
		// GAMMALN
		"=GAMMALN()":       {"#VALUE!", "GAMMALN requires 1 numeric argument"},
		"=GAMMALN(F1)":     {"#VALUE!", "GAMMALN requires 1 numeric argument"},
		"=GAMMALN(0)":      {"#N/A", "#N/A"},
		"=GAMMALN(INT(0))": {"#N/A", "#N/A"},
		// GAMMALN.PRECISE
		"=GAMMALN.PRECISE()":     {"#VALUE!", "GAMMALN.PRECISE requires 1 numeric argument"},
		"=GAMMALN.PRECISE(\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=GAMMALN.PRECISE(0)":    {"#NUM!", "#NUM!"},
		// GAUSS
		"=GAUSS()":     {"#VALUE!", "GAUSS requires 1 numeric argument"},
		"=GAUSS(\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// GEOMEAN
		"=GEOMEAN()":      {"#VALUE!", "GEOMEAN requires at least 1 numeric argument"},
		"=GEOMEAN(0)":     {"#NUM!", "#NUM!"},
		"=GEOMEAN(D1:D2)": {"#NUM!", "#NUM!"},
		"=GEOMEAN(\"\")":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// HARMEAN
		"=HARMEAN()":   {"#VALUE!", "HARMEAN requires at least 1 argument"},
		"=HARMEAN(-1)": {"#N/A", "#N/A"},
		"=HARMEAN(0)":  {"#N/A", "#N/A"},
		// HYPGEOM.DIST
		"=HYPGEOM.DIST()":                  {"#VALUE!", "HYPGEOM.DIST requires 5 arguments"},
		"=HYPGEOM.DIST(\"\",4,4,12,FALSE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=HYPGEOM.DIST(1,\"\",4,12,FALSE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=HYPGEOM.DIST(1,4,\"\",12,FALSE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=HYPGEOM.DIST(1,4,4,\"\",FALSE)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=HYPGEOM.DIST(1,4,4,12,\"\")":     {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=HYPGEOM.DIST(-1,4,4,12,FALSE)":   {"#NUM!", "#NUM!"},
		"=HYPGEOM.DIST(2,1,4,12,FALSE)":    {"#NUM!", "#NUM!"},
		"=HYPGEOM.DIST(2,4,1,12,FALSE)":    {"#NUM!", "#NUM!"},
		"=HYPGEOM.DIST(2,2,2,1,FALSE)":     {"#NUM!", "#NUM!"},
		"=HYPGEOM.DIST(1,0,4,12,FALSE)":    {"#NUM!", "#NUM!"},
		"=HYPGEOM.DIST(1,4,4,2,FALSE)":     {"#NUM!", "#NUM!"},
		"=HYPGEOM.DIST(1,4,0,12,FALSE)":    {"#NUM!", "#NUM!"},
		"=HYPGEOM.DIST(1,4,4,0,FALSE)":     {"#NUM!", "#NUM!"},
		// HYPGEOMDIST
		"=HYPGEOMDIST()":            {"#VALUE!", "HYPGEOMDIST requires 4 numeric arguments"},
		"=HYPGEOMDIST(\"\",4,4,12)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=HYPGEOMDIST(1,\"\",4,12)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=HYPGEOMDIST(1,4,\"\",12)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=HYPGEOMDIST(1,4,4,\"\")":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=HYPGEOMDIST(-1,4,4,12)":   {"#NUM!", "#NUM!"},
		"=HYPGEOMDIST(2,1,4,12)":    {"#NUM!", "#NUM!"},
		"=HYPGEOMDIST(2,4,1,12)":    {"#NUM!", "#NUM!"},
		"=HYPGEOMDIST(2,2,2,1)":     {"#NUM!", "#NUM!"},
		"=HYPGEOMDIST(1,0,4,12)":    {"#NUM!", "#NUM!"},
		"=HYPGEOMDIST(1,4,4,2)":     {"#NUM!", "#NUM!"},
		"=HYPGEOMDIST(1,4,0,12)":    {"#NUM!", "#NUM!"},
		"=HYPGEOMDIST(1,4,4,0)":     {"#NUM!", "#NUM!"},
		// INTERCEPT
		"=INTERCEPT()":            {"#VALUE!", "INTERCEPT requires 2 arguments"},
		"=INTERCEPT(A1:A2,B1:B1)": {"#N/A", "#N/A"},
		"=INTERCEPT(A4,A4)":       {"#DIV/0!", "#DIV/0!"},
		// KURT
		"=KURT()":          {"#VALUE!", "KURT requires at least 1 argument"},
		"=KURT(F1,INT(1))": {"#DIV/0!", "#DIV/0!"},
		// EXPON.DIST
		"=EXPON.DIST()":            {"#VALUE!", "EXPON.DIST requires 3 arguments"},
		"=EXPON.DIST(\"\",1,TRUE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=EXPON.DIST(0,\"\",TRUE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=EXPON.DIST(0,1,\"\")":    {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=EXPON.DIST(-1,1,TRUE)":   {"#NUM!", "#NUM!"},
		"=EXPON.DIST(1,0,TRUE)":    {"#NUM!", "#NUM!"},
		// EXPONDIST
		"=EXPONDIST()":            {"#VALUE!", "EXPONDIST requires 3 arguments"},
		"=EXPONDIST(\"\",1,TRUE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=EXPONDIST(0,\"\",TRUE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=EXPONDIST(0,1,\"\")":    {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=EXPONDIST(-1,1,TRUE)":   {"#NUM!", "#NUM!"},
		"=EXPONDIST(1,0,TRUE)":    {"#NUM!", "#NUM!"},
		// FDIST
		"=FDIST()":                {"#VALUE!", "FDIST requires 3 arguments"},
		"=FDIST(\"\",1,2)":        {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=FDIST(5,\"\",2)":        {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=FDIST(5,1,\"\")":        {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=FDIST(-1,1,2)":          {"#NUM!", "#NUM!"},
		"=FDIST(5,0,2)":           {"#NUM!", "#NUM!"},
		"=FDIST(5,10000000000,2)": {"#NUM!", "#NUM!"},
		"=FDIST(5,1,0)":           {"#NUM!", "#NUM!"},
		"=FDIST(5,1,10000000000)": {"#NUM!", "#NUM!"},
		// F.DIST
		"=F.DIST()":                     {"#VALUE!", "F.DIST requires 4 arguments"},
		"=F.DIST(\"\",2,5,TRUE)":        {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=F.DIST(1,\"\",5,TRUE)":        {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=F.DIST(1,2,\"\",TRUE)":        {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=F.DIST(1,2,5,\"\")":           {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=F.DIST(-1,1,2,TRUE)":          {"#NUM!", "#NUM!"},
		"=F.DIST(5,0,2,TRUE)":           {"#NUM!", "#NUM!"},
		"=F.DIST(5,10000000000,2,TRUE)": {"#NUM!", "#NUM!"},
		"=F.DIST(5,1,0,TRUE)":           {"#NUM!", "#NUM!"},
		"=F.DIST(5,1,10000000000,TRUE)": {"#NUM!", "#NUM!"},
		// F.DIST.RT
		"=F.DIST.RT()":                {"#VALUE!", "F.DIST.RT requires 3 arguments"},
		"=F.DIST.RT(\"\",1,2)":        {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=F.DIST.RT(5,\"\",2)":        {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=F.DIST.RT(5,1,\"\")":        {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=F.DIST.RT(-1,1,2)":          {"#NUM!", "#NUM!"},
		"=F.DIST.RT(5,0,2)":           {"#NUM!", "#NUM!"},
		"=F.DIST.RT(5,10000000000,2)": {"#NUM!", "#NUM!"},
		"=F.DIST.RT(5,1,0)":           {"#NUM!", "#NUM!"},
		"=F.DIST.RT(5,1,10000000000)": {"#NUM!", "#NUM!"},
		// F.INV
		"=F.INV()":           {"#VALUE!", "F.INV requires 3 arguments"},
		"=F.INV(\"\",1,2)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=F.INV(0.2,\"\",2)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=F.INV(0.2,1,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=F.INV(0,1,2)":      {"#NUM!", "#NUM!"},
		"=F.INV(0.2,0.5,2)":  {"#NUM!", "#NUM!"},
		"=F.INV(0.2,1,0.5)":  {"#NUM!", "#NUM!"},
		// FINV
		"=FINV()":           {"#VALUE!", "FINV requires 3 arguments"},
		"=FINV(\"\",1,2)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=FINV(0.2,\"\",2)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=FINV(0.2,1,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=FINV(0,1,2)":      {"#NUM!", "#NUM!"},
		"=FINV(0.2,0.5,2)":  {"#NUM!", "#NUM!"},
		"=FINV(0.2,1,0.5)":  {"#NUM!", "#NUM!"},
		// F.INV.RT
		"=F.INV.RT()":           {"#VALUE!", "F.INV.RT requires 3 arguments"},
		"=F.INV.RT(\"\",1,2)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=F.INV.RT(0.2,\"\",2)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=F.INV.RT(0.2,1,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=F.INV.RT(0,1,2)":      {"#NUM!", "#NUM!"},
		"=F.INV.RT(0.2,0.5,2)":  {"#NUM!", "#NUM!"},
		"=F.INV.RT(0.2,1,0.5)":  {"#NUM!", "#NUM!"},
		// LOGINV
		"=LOGINV()":             {"#VALUE!", "LOGINV requires 3 arguments"},
		"=LOGINV(\"\",2,0.2)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=LOGINV(0.3,\"\",0.2)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=LOGINV(0.3,2,\"\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=LOGINV(0,2,0.2)":      {"#NUM!", "#NUM!"},
		"=LOGINV(1,2,0.2)":      {"#NUM!", "#NUM!"},
		"=LOGINV(0.3,2,0)":      {"#NUM!", "#NUM!"},
		// LOGNORM.INV
		"=LOGNORM.INV()":             {"#VALUE!", "LOGNORM.INV requires 3 arguments"},
		"=LOGNORM.INV(\"\",2,0.2)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=LOGNORM.INV(0.3,\"\",0.2)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=LOGNORM.INV(0.3,2,\"\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=LOGNORM.INV(0,2,0.2)":      {"#NUM!", "#NUM!"},
		"=LOGNORM.INV(1,2,0.2)":      {"#NUM!", "#NUM!"},
		"=LOGNORM.INV(0.3,2,0)":      {"#NUM!", "#NUM!"},
		// LOGNORM.DIST
		"=LOGNORM.DIST()":                  {"#VALUE!", "LOGNORM.DIST requires 4 arguments"},
		"=LOGNORM.DIST(\"\",10,5,FALSE)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=LOGNORM.DIST(0.5,\"\",5,FALSE)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=LOGNORM.DIST(0.5,10,\"\",FALSE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=LOGNORM.DIST(0.5,10,5,\"\")":     {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=LOGNORM.DIST(0,10,5,FALSE)":      {"#NUM!", "#NUM!"},
		"=LOGNORM.DIST(0.5,10,0,FALSE)":    {"#NUM!", "#NUM!"},
		// LOGNORMDIST
		"=LOGNORMDIST()":           {"#VALUE!", "LOGNORMDIST requires 3 arguments"},
		"=LOGNORMDIST(\"\",10,5)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=LOGNORMDIST(12,\"\",5)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=LOGNORMDIST(12,10,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=LOGNORMDIST(0,2,5)":      {"#NUM!", "#NUM!"},
		"=LOGNORMDIST(12,10,0)":    {"#NUM!", "#NUM!"},
		// NEGBINOM.DIST
		"=NEGBINOM.DIST()":                 {"#VALUE!", "NEGBINOM.DIST requires 4 arguments"},
		"=NEGBINOM.DIST(\"\",12,0.5,TRUE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=NEGBINOM.DIST(6,\"\",0.5,TRUE)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=NEGBINOM.DIST(6,12,\"\",TRUE)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=NEGBINOM.DIST(6,12,0.5,\"\")":    {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=NEGBINOM.DIST(-1,12,0.5,TRUE)":   {"#NUM!", "#NUM!"},
		"=NEGBINOM.DIST(6,0,0.5,TRUE)":     {"#NUM!", "#NUM!"},
		"=NEGBINOM.DIST(6,12,-1,TRUE)":     {"#NUM!", "#NUM!"},
		"=NEGBINOM.DIST(6,12,2,TRUE)":      {"#NUM!", "#NUM!"},
		// NEGBINOMDIST
		"=NEGBINOMDIST()":            {"#VALUE!", "NEGBINOMDIST requires 3 arguments"},
		"=NEGBINOMDIST(\"\",12,0.5)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=NEGBINOMDIST(6,\"\",0.5)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=NEGBINOMDIST(6,12,\"\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=NEGBINOMDIST(-1,12,0.5)":   {"#NUM!", "#NUM!"},
		"=NEGBINOMDIST(6,0,0.5)":     {"#NUM!", "#NUM!"},
		"=NEGBINOMDIST(6,12,-1)":     {"#NUM!", "#NUM!"},
		"=NEGBINOMDIST(6,12,2)":      {"#NUM!", "#NUM!"},
		// NORM.DIST
		"=NORM.DIST()": {"#VALUE!", "NORM.DIST requires 4 arguments"},
		// NORMDIST
		"=NORMDIST()":               {"#VALUE!", "NORMDIST requires 4 arguments"},
		"=NORMDIST(\"\",0,0,FALSE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=NORMDIST(0,\"\",0,FALSE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=NORMDIST(0,0,\"\",FALSE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=NORMDIST(0,0,0,\"\")":     {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=NORMDIST(0,0,-1,TRUE)":    {"#N/A", "#N/A"},
		// NORM.INV
		"=NORM.INV()": {"#VALUE!", "NORM.INV requires 3 arguments"},
		// NORMINV
		"=NORMINV()":         {"#VALUE!", "NORMINV requires 3 arguments"},
		"=NORMINV(\"\",0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=NORMINV(0,\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=NORMINV(0,0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=NORMINV(0,0,-1)":   {"#N/A", "#N/A"},
		"=NORMINV(-1,0,0)":   {"#N/A", "#N/A"},
		"=NORMINV(0,0,0)":    {"#NUM!", "#NUM!"},
		// NORM.S.DIST
		"=NORM.S.DIST()": {"#VALUE!", "NORM.S.DIST requires 2 numeric arguments"},
		// NORMSDIST
		"=NORMSDIST()": {"#VALUE!", "NORMSDIST requires 1 numeric argument"},
		// NORM.S.INV
		"=NORM.S.INV()": {"#VALUE!", "NORM.S.INV requires 1 numeric argument"},
		// NORMSINV
		"=NORMSINV()": {"#VALUE!", "NORMSINV requires 1 numeric argument"},
		// LARGE
		"=LARGE()":           {"#VALUE!", "LARGE requires 2 arguments"},
		"=LARGE(A1:A5,0)":    {"#NUM!", "k should be > 0"},
		"=LARGE(A1:A5,6)":    {"#NUM!", "k should be <= length of array"},
		"=LARGE(A1:A5,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// MAX
		"=MAX()":     {"#VALUE!", "MAX requires at least 1 argument"},
		"=MAX(NA())": {"#N/A", "#N/A"},
		// MAXA
		"=MAXA()":     {"#VALUE!", "MAXA requires at least 1 argument"},
		"=MAXA(NA())": {"#N/A", "#N/A"},
		// MAXIFS
		"=MAXIFS()":                         {"#VALUE!", "MAXIFS requires at least 3 arguments"},
		"=MAXIFS(F2:F4,A2:A4,\">0\",D2:D9)": {"#N/A", "#N/A"},
		// MEDIAN
		"=MEDIAN()":      {"#VALUE!", "MEDIAN requires at least 1 argument"},
		"=MEDIAN(\"\")":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=MEDIAN(D1:D2)": {"#NUM!", "#NUM!"},
		// MIN
		"=MIN()":     {"#VALUE!", "MIN requires at least 1 argument"},
		"=MIN(NA())": {"#N/A", "#N/A"},
		// MINA
		"=MINA()":     {"#VALUE!", "MINA requires at least 1 argument"},
		"=MINA(NA())": {"#N/A", "#N/A"},
		// MINIFS
		"=MINIFS()":                         {"#VALUE!", "MINIFS requires at least 3 arguments"},
		"=MINIFS(F2:F4,A2:A4,\"<0\",D2:D9)": {"#N/A", "#N/A"},
		// PEARSON
		"=PEARSON()":            {"#VALUE!", "PEARSON requires 2 arguments"},
		"=PEARSON(A1:A2,B1:B1)": {"#N/A", "#N/A"},
		"=PEARSON(A4,A4)":       {"#DIV/0!", "#DIV/0!"},
		// PERCENTILE.EXC
		"=PERCENTILE.EXC()":           {"#VALUE!", "PERCENTILE.EXC requires 2 arguments"},
		"=PERCENTILE.EXC(A1:A4,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PERCENTILE.EXC(A1:A4,-1)":   {"#NUM!", "#NUM!"},
		"=PERCENTILE.EXC(A1:A4,0)":    {"#NUM!", "#NUM!"},
		"=PERCENTILE.EXC(A1:A4,1)":    {"#NUM!", "#NUM!"},
		"=PERCENTILE.EXC(NA(),0.5)":   {"#NUM!", "#NUM!"},
		// PERCENTILE.INC
		"=PERCENTILE.INC()": {"#VALUE!", "PERCENTILE.INC requires 2 arguments"},
		// PERCENTILE
		"=PERCENTILE()":       {"#VALUE!", "PERCENTILE requires 2 arguments"},
		"=PERCENTILE(0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PERCENTILE(0,-1)":   {"#N/A", "#N/A"},
		"=PERCENTILE(NA(),1)": {"#N/A", "#N/A"},
		// PERCENTRANK.EXC
		"=PERCENTRANK.EXC()":             {"#VALUE!", "PERCENTRANK.EXC requires 2 or 3 arguments"},
		"=PERCENTRANK.EXC(A1:B4,\"\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PERCENTRANK.EXC(A1:B4,0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PERCENTRANK.EXC(A1:B4,0,0)":    {"#NUM!", "PERCENTRANK.EXC arguments significance should be > 1"},
		"=PERCENTRANK.EXC(A1:B4,6)":      {"#N/A", "#N/A"},
		"=PERCENTRANK.EXC(NA(),1)":       {"#N/A", "#N/A"},
		// PERCENTRANK.INC
		"=PERCENTRANK.INC()":             {"#VALUE!", "PERCENTRANK.INC requires 2 or 3 arguments"},
		"=PERCENTRANK.INC(A1:B4,\"\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PERCENTRANK.INC(A1:B4,0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PERCENTRANK.INC(A1:B4,0,0)":    {"#NUM!", "PERCENTRANK.INC arguments significance should be > 1"},
		"=PERCENTRANK.INC(A1:B4,6)":      {"#N/A", "#N/A"},
		"=PERCENTRANK.INC(NA(),1)":       {"#N/A", "#N/A"},
		// PERCENTRANK
		"=PERCENTRANK()":             {"#VALUE!", "PERCENTRANK requires 2 or 3 arguments"},
		"=PERCENTRANK(A1:B4,\"\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PERCENTRANK(A1:B4,0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PERCENTRANK(A1:B4,0,0)":    {"#NUM!", "PERCENTRANK arguments significance should be > 1"},
		"=PERCENTRANK(A1:B4,6)":      {"#N/A", "#N/A"},
		"=PERCENTRANK(NA(),1)":       {"#N/A", "#N/A"},
		// PERMUT
		"=PERMUT()":       {"#VALUE!", "PERMUT requires 2 numeric arguments"},
		"=PERMUT(\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PERMUT(0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PERMUT(6,8)":    {"#N/A", "#N/A"},
		// PERMUTATIONA
		"=PERMUTATIONA()":       {"#VALUE!", "PERMUTATIONA requires 2 numeric arguments"},
		"=PERMUTATIONA(\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PERMUTATIONA(0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PERMUTATIONA(-1,0)":   {"#N/A", "#N/A"},
		"=PERMUTATIONA(0,-1)":   {"#N/A", "#N/A"},
		// PHI
		"=PHI()":     {"#VALUE!", "PHI requires 1 argument"},
		"=PHI(\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// QUARTILE
		"=QUARTILE()":           {"#VALUE!", "QUARTILE requires 2 arguments"},
		"=QUARTILE(A1:A4,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=QUARTILE(A1:A4,-1)":   {"#NUM!", "#NUM!"},
		"=QUARTILE(A1:A4,5)":    {"#NUM!", "#NUM!"},
		// QUARTILE.EXC
		"=QUARTILE.EXC()":           {"#VALUE!", "QUARTILE.EXC requires 2 arguments"},
		"=QUARTILE.EXC(A1:A4,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=QUARTILE.EXC(A1:A4,0)":    {"#NUM!", "#NUM!"},
		"=QUARTILE.EXC(A1:A4,4)":    {"#NUM!", "#NUM!"},
		// QUARTILE.INC
		"=QUARTILE.INC()": {"#VALUE!", "QUARTILE.INC requires 2 arguments"},
		// RANK
		"=RANK()":             {"#VALUE!", "RANK requires at least 2 arguments"},
		"=RANK(1,A1:B5,0,0)":  {"#VALUE!", "RANK requires at most 3 arguments"},
		"=RANK(-1,A1:B5)":     {"#N/A", "#N/A"},
		"=RANK(\"\",A1:B5)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=RANK(1,A1:B5,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// RANK.EQ
		"=RANK.EQ()":             {"#VALUE!", "RANK.EQ requires at least 2 arguments"},
		"=RANK.EQ(1,A1:B5,0,0)":  {"#VALUE!", "RANK.EQ requires at most 3 arguments"},
		"=RANK.EQ(-1,A1:B5)":     {"#N/A", "#N/A"},
		"=RANK.EQ(\"\",A1:B5)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=RANK.EQ(1,A1:B5,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// RSQ
		"=RSQ()":            {"#VALUE!", "RSQ requires 2 arguments"},
		"=RSQ(A1:A2,B1:B1)": {"#N/A", "#N/A"},
		"=RSQ(A4,A4)":       {"#DIV/0!", "#DIV/0!"},
		// SKEW
		"=SKEW()":     {"#VALUE!", "SKEW requires at least 1 argument"},
		"=SKEW(\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=SKEW(0)":    {"#DIV/0!", "#DIV/0!"},
		// SKEW.P
		"=SKEW.P()":     {"#VALUE!", "SKEW.P requires at least 1 argument"},
		"=SKEW.P(\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=SKEW.P(0)":    {"#DIV/0!", "#DIV/0!"},
		// SLOPE
		"=SLOPE()":            {"#VALUE!", "SLOPE requires 2 arguments"},
		"=SLOPE(A1:A2,B1:B1)": {"#N/A", "#N/A"},
		"=SLOPE(A4,A4)":       {"#DIV/0!", "#DIV/0!"},
		// SMALL
		"=SMALL()":           {"#VALUE!", "SMALL requires 2 arguments"},
		"=SMALL(A1:A5,0)":    {"#NUM!", "k should be > 0"},
		"=SMALL(A1:A5,6)":    {"#NUM!", "k should be <= length of array"},
		"=SMALL(A1:A5,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// STANDARDIZE
		"=STANDARDIZE()":         {"#VALUE!", "STANDARDIZE requires 3 arguments"},
		"=STANDARDIZE(\"\",0,5)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=STANDARDIZE(0,\"\",5)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=STANDARDIZE(0,0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=STANDARDIZE(0,0,0)":    {"#N/A", "#N/A"},
		// STDEVP
		"=STDEVP()":     {"#VALUE!", "STDEVP requires at least 1 argument"},
		"=STDEVP(\"\")": {"#DIV/0!", "#DIV/0!"},
		// STDEV.P
		"=STDEV.P()":     {"#VALUE!", "STDEV.P requires at least 1 argument"},
		"=STDEV.P(\"\")": {"#DIV/0!", "#DIV/0!"},
		// STDEVPA
		"=STDEVPA()":     {"#VALUE!", "STDEVPA requires at least 1 argument"},
		"=STDEVPA(\"\")": {"#DIV/0!", "#DIV/0!"},
		// T.DIST
		"=T.DIST()":             {"#VALUE!", "T.DIST requires 3 arguments"},
		"=T.DIST(\"\",10,TRUE)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=T.DIST(1,\"\",TRUE)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=T.DIST(1,10,\"\")":    {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=T.DIST(1,0,TRUE)":     {"#NUM!", "#NUM!"},
		"=T.DIST(1,-1,FALSE)":   {"#NUM!", "#NUM!"},
		"=T.DIST(1,0,FALSE)":    {"#DIV/0!", "#DIV/0!"},
		// T.DIST.2T
		"=T.DIST.2T()":        {"#VALUE!", "T.DIST.2T requires 2 arguments"},
		"=T.DIST.2T(\"\",10)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=T.DIST.2T(1,\"\")":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=T.DIST.2T(-1,10)":   {"#NUM!", "#NUM!"},
		"=T.DIST.2T(1,0)":     {"#NUM!", "#NUM!"},
		// T.DIST.RT
		"=T.DIST.RT()":        {"#VALUE!", "T.DIST.RT requires 2 arguments"},
		"=T.DIST.RT(\"\",10)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=T.DIST.RT(1,\"\")":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=T.DIST.RT(1,0)":     {"#NUM!", "#NUM!"},
		// TDIST
		"=TDIST()":          {"#VALUE!", "TDIST requires 3 arguments"},
		"=TDIST(\"\",10,1)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=TDIST(1,\"\",1)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=TDIST(1,10,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=TDIST(-1,10,1)":   {"#NUM!", "#NUM!"},
		"=TDIST(1,0,1)":     {"#NUM!", "#NUM!"},
		"=TDIST(1,10,0)":    {"#NUM!", "#NUM!"},
		// T.INV
		"=T.INV()":          {"#VALUE!", "T.INV requires 2 arguments"},
		"=T.INV(\"\",10)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=T.INV(0.25,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=T.INV(0,10)":      {"#NUM!", "#NUM!"},
		"=T.INV(1,10)":      {"#NUM!", "#NUM!"},
		"=T.INV(0.25,0.5)":  {"#NUM!", "#NUM!"},
		// T.INV.2T
		"=T.INV.2T()":          {"#VALUE!", "T.INV.2T requires 2 arguments"},
		"=T.INV.2T(\"\",10)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=T.INV.2T(0.25,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=T.INV.2T(0,10)":      {"#NUM!", "#NUM!"},
		"=T.INV.2T(0.25,0.5)":  {"#NUM!", "#NUM!"},
		// TINV
		"=TINV()":          {"#VALUE!", "TINV requires 2 arguments"},
		"=TINV(\"\",10)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=TINV(0.25,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=TINV(0,10)":      {"#NUM!", "#NUM!"},
		"=TINV(0.25,0.5)":  {"#NUM!", "#NUM!"},
		// TRIMMEAN
		"=TRIMMEAN()":        {"#VALUE!", "TRIMMEAN requires 2 arguments"},
		"=TRIMMEAN(A1,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=TRIMMEAN(A1,1)":    {"#NUM!", "#NUM!"},
		"=TRIMMEAN(A1,-1)":   {"#NUM!", "#NUM!"},
		// VAR
		"=VAR()": {"#VALUE!", "VAR requires at least 1 argument"},
		// VARA
		"=VARA()": {"#VALUE!", "VARA requires at least 1 argument"},
		// VARP
		"=VARP()":     {"#VALUE!", "VARP requires at least 1 argument"},
		"=VARP(\"\")": {"#DIV/0!", "#DIV/0!"},
		// VAR.P
		"=VAR.P()":     {"#VALUE!", "VAR.P requires at least 1 argument"},
		"=VAR.P(\"\")": {"#DIV/0!", "#DIV/0!"},
		// VAR.S
		"=VAR.S()": {"#VALUE!", "VAR.S requires at least 1 argument"},
		// VARPA
		"=VARPA()": {"#VALUE!", "VARPA requires at least 1 argument"},
		// WEIBULL
		"=WEIBULL()":               {"#VALUE!", "WEIBULL requires 4 arguments"},
		"=WEIBULL(\"\",1,1,FALSE)": {"#VALUE!", "#VALUE!"},
		"=WEIBULL(1,0,1,FALSE)":    {"#N/A", "#N/A"},
		"=WEIBULL(1,1,-1,FALSE)":   {"#N/A", "#N/A"},
		// WEIBULL.DIST
		"=WEIBULL.DIST()":               {"#VALUE!", "WEIBULL.DIST requires 4 arguments"},
		"=WEIBULL.DIST(\"\",1,1,FALSE)": {"#VALUE!", "#VALUE!"},
		"=WEIBULL.DIST(1,0,1,FALSE)":    {"#N/A", "#N/A"},
		"=WEIBULL.DIST(1,1,-1,FALSE)":   {"#N/A", "#N/A"},
		// Z.TEST
		"=Z.TEST(A1)":        {"#VALUE!", "Z.TEST requires at least 2 arguments"},
		"=Z.TEST(A1,0,0,0)":  {"#VALUE!", "Z.TEST accepts at most 3 arguments"},
		"=Z.TEST(H1,0)":      {"#N/A", "#N/A"},
		"=Z.TEST(A1,\"\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=Z.TEST(A1,1)":      {"#DIV/0!", "#DIV/0!"},
		"=Z.TEST(A1,1,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// ZTEST
		"=ZTEST(A1)":        {"#VALUE!", "ZTEST requires at least 2 arguments"},
		"=ZTEST(A1,0,0,0)":  {"#VALUE!", "ZTEST accepts at most 3 arguments"},
		"=ZTEST(H1,0)":      {"#N/A", "#N/A"},
		"=ZTEST(A1,\"\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ZTEST(A1,1)":      {"#DIV/0!", "#DIV/0!"},
		"=ZTEST(A1,1,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// Information Functions
		// ERROR.TYPE
		"=ERROR.TYPE()":  {"#VALUE!", "ERROR.TYPE requires 1 argument"},
		"=ERROR.TYPE(1)": {"#N/A", "#N/A"},
		// ISBLANK
		"=ISBLANK(A1,A2)": {"#VALUE!", "ISBLANK requires 1 argument"},
		// ISERR
		"=ISERR()": {"#VALUE!", "ISERR requires 1 argument"},
		// ISERROR
		"=ISERROR()": {"#VALUE!", "ISERROR requires 1 argument"},
		// ISEVEN
		"=ISEVEN()":         {"#VALUE!", "ISEVEN requires 1 argument"},
		"=ISEVEN(\"text\")": {"#VALUE!", "#VALUE!"},
		"=ISEVEN(A1:A2)":    {"#VALUE!", "#VALUE!"},
		// ISFORMULA
		"=ISFORMULA()": {"#VALUE!", "ISFORMULA requires 1 argument"},
		// ISLOGICAL
		"=ISLOGICAL()": {"#VALUE!", "ISLOGICAL requires 1 argument"},
		// ISNA
		"=ISNA()": {"#VALUE!", "ISNA requires 1 argument"},
		// ISNONTEXT
		"=ISNONTEXT()": {"#VALUE!", "ISNONTEXT requires 1 argument"},
		// ISNUMBER
		"=ISNUMBER()": {"#VALUE!", "ISNUMBER requires 1 argument"},
		// ISODD
		"=ISODD()":         {"#VALUE!", "ISODD requires 1 argument"},
		"=ISODD(\"text\")": {"#VALUE!", "#VALUE!"},
		// ISREF
		"=ISREF()": {"#VALUE!", "ISREF requires 1 argument"},
		// ISTEXT
		"=ISTEXT()": {"#VALUE!", "ISTEXT requires 1 argument"},
		// N
		"=N()":     {"#VALUE!", "N requires 1 argument"},
		"=N(NA())": {"#N/A", "#N/A"},
		// NA
		"=NA()":  {"#N/A", "#N/A"},
		"=NA(1)": {"#VALUE!", "NA accepts no arguments"},
		// SHEET
		"=SHEET(\"\",\"\")":  {"#VALUE!", "SHEET accepts at most 1 argument"},
		"=SHEET(\"Sheet2\")": {"#N/A", "#N/A"},
		// SHEETS
		"=SHEETS(\"\",\"\")":  {"#VALUE!", "SHEETS accepts at most 1 argument"},
		"=SHEETS(\"Sheet1\")": {"#N/A", "#N/A"},
		// TYPE
		"=TYPE()": {"#VALUE!", "TYPE requires 1 argument"},
		// T
		"=T()":     {"#VALUE!", "T requires 1 argument"},
		"=T(NA())": {"#N/A", "#N/A"},
		// Logical Functions
		// AND
		"=AND(\"text\")":                 {"#VALUE!", "#VALUE!"},
		"=AND(A1:B1)":                    {"#VALUE!", "#VALUE!"},
		"=AND(\"1\",\"TRUE\",\"FALSE\")": {"#VALUE!", "#VALUE!"},
		"=AND()":                         {"#VALUE!", "AND requires at least 1 argument"},
		"=AND(1" + strings.Repeat(",1", 30) + ")": {"#VALUE!", "AND accepts at most 30 arguments"},
		// FALSE
		"=FALSE(A1)": {"#VALUE!", "FALSE takes no arguments"},
		// IFERROR
		"=IFERROR()": {"#VALUE!", "IFERROR requires 2 arguments"},
		// IFNA
		"=IFNA()": {"#VALUE!", "IFNA requires 2 arguments"},
		// IFS
		"=IFS()":            {"#VALUE!", "IFS requires at least 2 arguments"},
		"=IFS(FALSE,FALSE)": {"#N/A", "#N/A"},
		// NOT
		"=NOT()":      {"#VALUE!", "NOT requires 1 argument"},
		"=NOT(NOT())": {"#VALUE!", "NOT requires 1 argument"},
		"=NOT(\"\")":  {"#VALUE!", "NOT expects 1 boolean or numeric argument"},
		// OR
		"=OR(\"text\")":                          {"#VALUE!", "#VALUE!"},
		"=OR(\"1\",\"TRUE\",\"FALSE\")":          {"#VALUE!", "#VALUE!"},
		"=OR()":                                  {"#VALUE!", "OR requires at least 1 argument"},
		"=OR(1" + strings.Repeat(",1", 30) + ")": {"#VALUE!", "OR accepts at most 30 arguments"},
		// SWITCH
		"=SWITCH()":      {"#VALUE!", "SWITCH requires at least 3 arguments"},
		"=SWITCH(0,1,2)": {"#N/A", "#N/A"},
		// TRUE
		"=TRUE(A1)": {"#VALUE!", "TRUE takes no arguments"},
		// XOR
		"=XOR()":              {"#VALUE!", "XOR requires at least 1 argument"},
		"=XOR(\"1\")":         {"#VALUE!", "#VALUE!"},
		"=XOR(\"text\")":      {"#VALUE!", "#VALUE!"},
		"=XOR(XOR(\"text\"))": {"#VALUE!", "#VALUE!"},
		// Date and Time Functions
		// DATE
		"=DATE()":                 {"#VALUE!", "DATE requires 3 number arguments"},
		"=DATE(\"text\",10,21)":   {"#VALUE!", "DATE requires 3 number arguments"},
		"=DATE(2020,\"text\",21)": {"#VALUE!", "DATE requires 3 number arguments"},
		"=DATE(2020,10,\"text\")": {"#VALUE!", "DATE requires 3 number arguments"},
		// DATEDIF
		"=DATEDIF()":                  {"#VALUE!", "DATEDIF requires 3 number arguments"},
		"=DATEDIF(\"\",\"\",\"\")":    {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DATEDIF(43891,43101,\"Y\")": {"#NUM!", "start_date > end_date"},
		"=DATEDIF(43101,43891,\"x\")": {"#VALUE!", "DATEDIF has invalid unit"},
		// DATEVALUE
		"=DATEVALUE()":             {"#VALUE!", "DATEVALUE requires 1 argument"},
		"=DATEVALUE(\"01/01\")":    {"#VALUE!", "#VALUE!"}, // valid in Excel, which uses years by the system date
		"=DATEVALUE(\"1900-0-0\")": {"#VALUE!", "#VALUE!"},
		// DAY
		"=DAY()":         {"#VALUE!", "DAY requires exactly 1 argument"},
		"=DAY(-1)":       {"#NUM!", "DAY only accepts positive argument"},
		"=DAY(0,0)":      {"#VALUE!", "DAY requires exactly 1 argument"},
		"=DAY(\"text\")": {"#VALUE!", "#VALUE!"},
		"=DAY(\"January 25, 2020 9223372036854775808 AM\")":                   {"#VALUE!", "#VALUE!"},
		"=DAY(\"January 25, 2020 9223372036854775808:00 AM\")":                {"#VALUE!", "#VALUE!"},
		"=DAY(\"January 25, 2020 00:9223372036854775808 AM\")":                {"#VALUE!", "#VALUE!"},
		"=DAY(\"January 25, 2020 9223372036854775808:00.0 AM\")":              {"#VALUE!", "#VALUE!"},
		"=DAY(\"January 25, 2020 0:1" + strings.Repeat("0", 309) + ".0 AM\")": {"#VALUE!", "#VALUE!"},
		"=DAY(\"January 25, 2020 9223372036854775808:00:00 AM\")":             {"#VALUE!", "#VALUE!"},
		"=DAY(\"January 25, 2020 0:9223372036854775808:0 AM\")":               {"#VALUE!", "#VALUE!"},
		"=DAY(\"January 25, 2020 0:0:1" + strings.Repeat("0", 309) + " AM\")": {"#VALUE!", "#VALUE!"},
		"=DAY(\"January 25, 2020 0:61:0 AM\")":                                {"#VALUE!", "#VALUE!"},
		"=DAY(\"January 25, 2020 0:00:60 AM\")":                               {"#VALUE!", "#VALUE!"},
		"=DAY(\"January 25, 2020 24:00:00\")":                                 {"#VALUE!", "#VALUE!"},
		"=DAY(\"January 25, 2020 00:00:10001\")":                              {"#VALUE!", "#VALUE!"},
		"=DAY(\"9223372036854775808/25/2020\")":                               {"#VALUE!", "#VALUE!"},
		"=DAY(\"01/9223372036854775808/2020\")":                               {"#VALUE!", "#VALUE!"},
		"=DAY(\"01/25/9223372036854775808\")":                                 {"#VALUE!", "#VALUE!"},
		"=DAY(\"01/25/10000\")":                                               {"#VALUE!", "#VALUE!"},
		"=DAY(\"01/25/100\")":                                                 {"#VALUE!", "#VALUE!"},
		"=DAY(\"January 9223372036854775808, 2020\")":                         {"#VALUE!", "#VALUE!"},
		"=DAY(\"January 25, 9223372036854775808\")":                           {"#VALUE!", "#VALUE!"},
		"=DAY(\"January 25, 10000\")":                                         {"#VALUE!", "#VALUE!"},
		"=DAY(\"January 25, 100\")":                                           {"#VALUE!", "#VALUE!"},
		"=DAY(\"9223372036854775808-25-2020\")":                               {"#VALUE!", "#VALUE!"},
		"=DAY(\"01-9223372036854775808-2020\")":                               {"#VALUE!", "#VALUE!"},
		"=DAY(\"01-25-9223372036854775808\")":                                 {"#VALUE!", "#VALUE!"},
		"=DAY(\"1900-0-0\")":                                                  {"#VALUE!", "#VALUE!"},
		"=DAY(\"14-25-1900\")":                                                {"#VALUE!", "#VALUE!"},
		"=DAY(\"3-January-9223372036854775808\")":                             {"#VALUE!", "#VALUE!"},
		"=DAY(\"9223372036854775808-January-1900\")":                          {"#VALUE!", "#VALUE!"},
		"=DAY(\"0-January-1900\")":                                            {"#VALUE!", "#VALUE!"},
		// DAYS
		"=DAYS()":       {"#VALUE!", "DAYS requires 2 arguments"},
		"=DAYS(\"\",0)": {"#VALUE!", "#VALUE!"},
		"=DAYS(0,\"\")": {"#VALUE!", "#VALUE!"},
		"=DAYS(NA(),0)": {"#VALUE!", "#VALUE!"},
		"=DAYS(0,NA())": {"#VALUE!", "#VALUE!"},
		// DAYS360
		"=DAYS360(\"12/12/1999\")":                           {"#VALUE!", "DAYS360 requires at least 2 arguments"},
		"=DAYS360(\"12/12/1999\", \"11/30/1999\",TRUE,\"\")": {"#VALUE!", "DAYS360 requires at most 3 arguments"},
		"=DAYS360(\"12/12/1999\", \"11/30/1999\",\"\")":      {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=DAYS360(\"12/12/1999\", \"\")":                     {"#VALUE!", "#VALUE!"},
		"=DAYS360(\"\", \"11/30/1999\")":                     {"#VALUE!", "#VALUE!"},
		// EDATE
		"=EDATE()":                      {"#VALUE!", "EDATE requires 2 arguments"},
		"=EDATE(0,\"\")":                {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=EDATE(-1,0)":                  {"#NUM!", "#NUM!"},
		"=EDATE(\"\",0)":                {"#VALUE!", "#VALUE!"},
		"=EDATE(\"January 25, 100\",0)": {"#VALUE!", "#VALUE!"},
		// EOMONTH
		"=EOMONTH()":                      {"#VALUE!", "EOMONTH requires 2 arguments"},
		"=EOMONTH(0,\"\")":                {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=EOMONTH(-1,0)":                  {"#NUM!", "#NUM!"},
		"=EOMONTH(\"\",0)":                {"#VALUE!", "#VALUE!"},
		"=EOMONTH(\"January 25, 100\",0)": {"#VALUE!", "#VALUE!"},
		// HOUR
		"=HOUR()":             {"#VALUE!", "HOUR requires exactly 1 argument"},
		"=HOUR(-1)":           {"#NUM!", "HOUR only accepts positive argument"},
		"=HOUR(\"\")":         {"#VALUE!", "#VALUE!"},
		"=HOUR(\"25:10:55\")": {"#VALUE!", "#VALUE!"},
		// ISOWEEKNUM
		"=ISOWEEKNUM()":                    {"#VALUE!", "ISOWEEKNUM requires 1 argument"},
		"=ISOWEEKNUM(\"\")":                {"#VALUE!", "#VALUE!"},
		"=ISOWEEKNUM(\"January 25, 100\")": {"#VALUE!", "#VALUE!"},
		"=ISOWEEKNUM(-1)":                  {"#NUM!", "#NUM!"},
		// MINUTE
		"=MINUTE()":             {"#VALUE!", "MINUTE requires exactly 1 argument"},
		"=MINUTE(-1)":           {"#NUM!", "MINUTE only accepts positive argument"},
		"=MINUTE(\"\")":         {"#VALUE!", "#VALUE!"},
		"=MINUTE(\"13:60:55\")": {"#VALUE!", "#VALUE!"},
		// MONTH
		"=MONTH()":                    {"#VALUE!", "MONTH requires exactly 1 argument"},
		"=MONTH(0,0)":                 {"#VALUE!", "MONTH requires exactly 1 argument"},
		"=MONTH(-1)":                  {"#NUM!", "MONTH only accepts positive argument"},
		"=MONTH(\"text\")":            {"#VALUE!", "#VALUE!"},
		"=MONTH(\"January 25, 100\")": {"#VALUE!", "#VALUE!"},
		// YEAR
		"=YEAR()":                    {"#VALUE!", "YEAR requires exactly 1 argument"},
		"=YEAR(0,0)":                 {"#VALUE!", "YEAR requires exactly 1 argument"},
		"=YEAR(-1)":                  {"#NUM!", "YEAR only accepts positive argument"},
		"=YEAR(\"text\")":            {"#VALUE!", "#VALUE!"},
		"=YEAR(\"January 25, 100\")": {"#VALUE!", "#VALUE!"},
		// YEARFRAC
		"=YEARFRAC()":                 {"#VALUE!", "YEARFRAC requires 3 or 4 arguments"},
		"=YEARFRAC(42005,42094,5)":    {"#NUM!", "invalid basis"},
		"=YEARFRAC(\"\",42094,5)":     {"#VALUE!", "#VALUE!"},
		"=YEARFRAC(42005,\"\",5)":     {"#VALUE!", "#VALUE!"},
		"=YEARFRAC(42005,42094,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// NOW
		"=NOW(A1)": {"#VALUE!", "NOW accepts no arguments"},
		// SECOND
		"=SECOND()":          {"#VALUE!", "SECOND requires exactly 1 argument"},
		"=SECOND(-1)":        {"#NUM!", "SECOND only accepts positive argument"},
		"=SECOND(\"\")":      {"#VALUE!", "#VALUE!"},
		"=SECOND(\"25:55\")": {"#VALUE!", "#VALUE!"},
		// TIME
		"=TIME()":         {"#VALUE!", "TIME requires 3 number arguments"},
		"=TIME(\"\",0,0)": {"#VALUE!", "TIME requires 3 number arguments"},
		"=TIME(0,0,-1)":   {"#NUM!", "#NUM!"},
		// TIMEVALUE
		"=TIMEVALUE()":          {"#VALUE!", "TIMEVALUE requires exactly 1 argument"},
		"=TIMEVALUE(1)":         {"#VALUE!", "#VALUE!"},
		"=TIMEVALUE(-1)":        {"#VALUE!", "#VALUE!"},
		"=TIMEVALUE(\"25:55\")": {"#VALUE!", "#VALUE!"},
		// TODAY
		"=TODAY(A1)": {"#VALUE!", "TODAY accepts no arguments"},
		// WEEKDAY
		"=WEEKDAY()":                    {"#VALUE!", "WEEKDAY requires at least 1 argument"},
		"=WEEKDAY(0,1,0)":               {"#VALUE!", "WEEKDAY allows at most 2 arguments"},
		"=WEEKDAY(0,\"\")":              {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=WEEKDAY(\"\",1)":              {"#VALUE!", "#VALUE!"},
		"=WEEKDAY(0,0)":                 {"#VALUE!", "#VALUE!"},
		"=WEEKDAY(\"January 25, 100\")": {"#VALUE!", "#VALUE!"},
		"=WEEKDAY(-1,1)":                {"#NUM!", "#NUM!"},
		// WEEKNUM
		"=WEEKNUM()":                    {"#VALUE!", "WEEKNUM requires at least 1 argument"},
		"=WEEKNUM(0,1,0)":               {"#VALUE!", "WEEKNUM allows at most 2 arguments"},
		"=WEEKNUM(0,\"\")":              {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=WEEKNUM(\"\",1)":              {"#VALUE!", "#VALUE!"},
		"=WEEKNUM(\"January 25, 100\")": {"#VALUE!", "#VALUE!"},
		"=WEEKNUM(0,0)":                 {"#NUM!", "#NUM!"},
		"=WEEKNUM(-1,1)":                {"#NUM!", "#NUM!"},
		// Text Functions
		// ARRAYTOTEXT
		"=ARRAYTOTEXT()":        {"#VALUE!", "ARRAYTOTEXT requires at least 1 argument"},
		"=ARRAYTOTEXT(A1,0,0)":  {"#VALUE!", "ARRAYTOTEXT allows at most 2 arguments"},
		"=ARRAYTOTEXT(A1,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ARRAYTOTEXT(A1,2)":    {"#VALUE!", "#VALUE!"},
		// CHAR
		"=CHAR()":     {"#VALUE!", "CHAR requires 1 argument"},
		"=CHAR(-1)":   {"#VALUE!", "#VALUE!"},
		"=CHAR(256)":  {"#VALUE!", "#VALUE!"},
		"=CHAR(\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// CLEAN
		"=CLEAN()":    {"#VALUE!", "CLEAN requires 1 argument"},
		"=CLEAN(1,2)": {"#VALUE!", "CLEAN requires 1 argument"},
		// CODE
		"=CODE()":    {"#VALUE!", "CODE requires 1 argument"},
		"=CODE(1,2)": {"#VALUE!", "CODE requires 1 argument"},
		// CONCAT
		"=CONCAT(NA())":  {"#N/A", "#N/A"},
		"=CONCAT(1,1/0)": {"#DIV/0!", "#DIV/0!"},
		// CONCATENATE
		"=CONCATENATE(NA())":  {"#N/A", "#N/A"},
		"=CONCATENATE(1,1/0)": {"#DIV/0!", "#DIV/0!"},
		// DBCS
		"=DBCS(NA())": {"#N/A", "#N/A"},
		"=DBCS()":     {"#VALUE!", "DBCS requires 1 argument"},
		// EXACT
		"=EXACT()":      {"#VALUE!", "EXACT requires 2 arguments"},
		"=EXACT(1,2,3)": {"#VALUE!", "EXACT requires 2 arguments"},
		// FIXED
		"=FIXED()":         {"#VALUE!", "FIXED requires at least 1 argument"},
		"=FIXED(0,1,2,3)":  {"#VALUE!", "FIXED allows at most 3 arguments"},
		"=FIXED(\"\")":     {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=FIXED(0,\"\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=FIXED(0,0,\"\")": {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		// FIND
		"=FIND()":                 {"#VALUE!", "FIND requires at least 2 arguments"},
		"=FIND(1,2,3,4)":          {"#VALUE!", "FIND allows at most 3 arguments"},
		"=FIND(\"x\",\"\")":       {"#VALUE!", "#VALUE!"},
		"=FIND(\"x\",\"x\",-1)":   {"#VALUE!", "#VALUE!"},
		"=FIND(\"x\",\"x\",\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// FINDB
		"=FINDB()":                 {"#VALUE!", "FINDB requires at least 2 arguments"},
		"=FINDB(1,2,3,4)":          {"#VALUE!", "FINDB allows at most 3 arguments"},
		"=FINDB(\"x\",\"\")":       {"#VALUE!", "#VALUE!"},
		"=FINDB(\"x\",\"x\",-1)":   {"#VALUE!", "#VALUE!"},
		"=FINDB(\"x\",\"x\",\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// LEFT
		"=LEFT()":          {"#VALUE!", "LEFT requires at least 1 argument"},
		"=LEFT(\"\",2,3)":  {"#VALUE!", "LEFT allows at most 2 arguments"},
		"=LEFT(\"\",\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=LEFT(\"\",-1)":   {"#VALUE!", "#VALUE!"},
		// LEFTB
		"=LEFTB()":          {"#VALUE!", "LEFTB requires at least 1 argument"},
		"=LEFTB(\"\",2,3)":  {"#VALUE!", "LEFTB allows at most 2 arguments"},
		"=LEFTB(\"\",\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=LEFTB(\"\",-1)":   {"#VALUE!", "#VALUE!"},
		// LEN
		"=LEN()": {"#VALUE!", "LEN requires 1 string argument"},
		// LENB
		"=LENB()": {"#VALUE!", "LENB requires 1 string argument"},
		// LOWER
		"=LOWER()":    {"#VALUE!", "LOWER requires 1 argument"},
		"=LOWER(1,2)": {"#VALUE!", "LOWER requires 1 argument"},
		// MID
		"=MID()":            {"#VALUE!", "MID requires 3 arguments"},
		"=MID(\"\",0,1)":    {"#VALUE!", "#VALUE!"},
		"=MID(\"\",1,-1)":   {"#VALUE!", "#VALUE!"},
		"=MID(\"\",\"\",1)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=MID(\"\",1,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// MIDB
		"=MIDB()":            {"#VALUE!", "MIDB requires 3 arguments"},
		"=MIDB(\"\",0,1)":    {"#VALUE!", "#VALUE!"},
		"=MIDB(\"\",1,-1)":   {"#VALUE!", "#VALUE!"},
		"=MIDB(\"\",\"\",1)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=MIDB(\"\",1,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// PROPER
		"=PROPER()":    {"#VALUE!", "PROPER requires 1 argument"},
		"=PROPER(1,2)": {"#VALUE!", "PROPER requires 1 argument"},
		// REPLACE
		"=REPLACE()":                           {"#VALUE!", "REPLACE requires 4 arguments"},
		"=REPLACE(\"text\",0,4,\"string\")":    {"#VALUE!", "#VALUE!"},
		"=REPLACE(\"text\",\"\",0,\"string\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=REPLACE(\"text\",1,\"\",\"string\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// REPLACEB
		"=REPLACEB()":                           {"#VALUE!", "REPLACEB requires 4 arguments"},
		"=REPLACEB(\"text\",0,4,\"string\")":    {"#VALUE!", "#VALUE!"},
		"=REPLACEB(\"text\",\"\",0,\"string\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=REPLACEB(\"text\",1,\"\",\"string\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// REPT
		"=REPT()":            {"#VALUE!", "REPT requires 2 arguments"},
		"=REPT(INT(0),2)":    {"#VALUE!", "REPT requires first argument to be a string"},
		"=REPT(\"*\",\"*\")": {"#VALUE!", "REPT requires second argument to be a number"},
		"=REPT(\"*\",-1)":    {"#VALUE!", "REPT requires second argument to be >= 0"},
		// RIGHT
		"=RIGHT()":          {"#VALUE!", "RIGHT requires at least 1 argument"},
		"=RIGHT(\"\",2,3)":  {"#VALUE!", "RIGHT allows at most 2 arguments"},
		"=RIGHT(\"\",\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=RIGHT(\"\",-1)":   {"#VALUE!", "#VALUE!"},
		// RIGHTB
		"=RIGHTB()":          {"#VALUE!", "RIGHTB requires at least 1 argument"},
		"=RIGHTB(\"\",2,3)":  {"#VALUE!", "RIGHTB allows at most 2 arguments"},
		"=RIGHTB(\"\",\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=RIGHTB(\"\",-1)":   {"#VALUE!", "#VALUE!"},
		// SUBSTITUTE
		"=SUBSTITUTE()":                    {"#VALUE!", "SUBSTITUTE requires 3 or 4 arguments"},
		"=SUBSTITUTE(\"\",\"\",\"\",\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=SUBSTITUTE(\"\",\"\",\"\",0)":    {"#VALUE!", "instance_num should be > 0"},
		// TEXT
		"=TEXT()":          {"#VALUE!", "TEXT requires 2 arguments"},
		"=TEXT(NA(),\"\")": {"#N/A", "#N/A"},
		"=TEXT(0,NA())":    {"#N/A", "#N/A"},
		// TEXTAFTER
		"=TEXTAFTER()": {"#VALUE!", "TEXTAFTER requires at least 2 arguments"},
		"=TEXTAFTER(\"Red riding hood's, red hood\",\"hood\",1,0,0,\"\",0)": {"#VALUE!", "TEXTAFTER accepts at most 6 arguments"},
		"=TEXTAFTER(\"Red riding hood's, red hood\",\"hood\",\"\")":         {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=TEXTAFTER(\"Red riding hood's, red hood\",\"hood\",1,\"\")":       {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=TEXTAFTER(\"Red riding hood's, red hood\",\"hood\",1,0,\"\")":     {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=TEXTAFTER(\"\",\"hood\")":                                         {"#N/A", "#N/A"},
		"=TEXTAFTER(\"Red riding hood's, red hood\",\"hood\",0)":            {"#VALUE!", "#VALUE!"},
		"=TEXTAFTER(\"Red riding hood's, red hood\",\"hood\",28)":           {"#VALUE!", "#VALUE!"},
		// TEXTBEFORE
		"=TEXTBEFORE()": {"#VALUE!", "TEXTBEFORE requires at least 2 arguments"},
		"=TEXTBEFORE(\"Red riding hood's, red hood\",\"hood\",1,0,0,\"\",0)": {"#VALUE!", "TEXTBEFORE accepts at most 6 arguments"},
		"=TEXTBEFORE(\"Red riding hood's, red hood\",\"hood\",\"\")":         {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=TEXTBEFORE(\"Red riding hood's, red hood\",\"hood\",1,\"\")":       {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=TEXTBEFORE(\"Red riding hood's, red hood\",\"hood\",1,0,\"\")":     {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=TEXTBEFORE(\"\",\"hood\")":                                         {"#N/A", "#N/A"},
		"=TEXTBEFORE(\"Red riding hood's, red hood\",\"hood\",0)":            {"#VALUE!", "#VALUE!"},
		"=TEXTBEFORE(\"Red riding hood's, red hood\",\"hood\",28)":           {"#VALUE!", "#VALUE!"},
		// TEXTJOIN
		"=TEXTJOIN()":               {"#VALUE!", "TEXTJOIN requires at least 3 arguments"},
		"=TEXTJOIN(\"\",\"\",1)":    {"#VALUE!", "#VALUE!"},
		"=TEXTJOIN(\"\",TRUE,NA())": {"#N/A", "#N/A"},
		"=TEXTJOIN(\"\",TRUE," + strings.Repeat("0,", 250) + ",0)": {"#VALUE!", "TEXTJOIN accepts at most 252 arguments"},
		"=TEXTJOIN(\",\",FALSE,REPT(\"*\",32768))":                 {"#VALUE!", "TEXTJOIN function exceeds 32767 characters"},
		// TRIM
		"=TRIM()":    {"#VALUE!", "TRIM requires 1 argument"},
		"=TRIM(1,2)": {"#VALUE!", "TRIM requires 1 argument"},
		// UNICHAR
		"=UNICHAR()":      {"#VALUE!", "UNICHAR requires 1 argument"},
		"=UNICHAR(\"\")":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=UNICHAR(55296)": {"#VALUE!", "#VALUE!"},
		"=UNICHAR(0)":     {"#VALUE!", "#VALUE!"},
		// UNICODE
		"=UNICODE()":     {"#VALUE!", "UNICODE requires 1 argument"},
		"=UNICODE(\"\")": {"#VALUE!", "#VALUE!"},
		// VALUE
		"=VALUE()":     {"#VALUE!", "VALUE requires 1 argument"},
		"=VALUE(\"\")": {"#VALUE!", "#VALUE!"},
		// VALUETOTEXT
		"=VALUETOTEXT()":        {"#VALUE!", "VALUETOTEXT requires at least 1 argument"},
		"=VALUETOTEXT(A1,0,0)":  {"#VALUE!", "VALUETOTEXT allows at most 2 arguments"},
		"=VALUETOTEXT(A1,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=VALUETOTEXT(A1,2)":    {"#VALUE!", "#VALUE!"},
		// UPPER
		"=UPPER()":    {"#VALUE!", "UPPER requires 1 argument"},
		"=UPPER(1,2)": {"#VALUE!", "UPPER requires 1 argument"},
		// Conditional Functions
		// IF
		"=IF()":        {"#VALUE!", "IF requires at least 1 argument"},
		"=IF(0,1,2,3)": {"#VALUE!", "IF accepts at most 3 arguments"},
		"=IF(D1,1,2)":  {"#VALUE!", "strconv.ParseBool: parsing \"Month\": invalid syntax"},
		// Excel Lookup and Reference Functions
		// ADDRESS
		"=ADDRESS()":                        {"#VALUE!", "ADDRESS requires at least 2 arguments"},
		"=ADDRESS(1,1,1,TRUE,\"Sheet1\",0)": {"#VALUE!", "ADDRESS requires at most 5 arguments"},
		"=ADDRESS(\"\",1,1,TRUE)":           {"#VALUE!", "#VALUE!"},
		"=ADDRESS(1,\"\",1,TRUE)":           {"#VALUE!", "#VALUE!"},
		"=ADDRESS(1,1,\"\",TRUE)":           {"#VALUE!", "#VALUE!"},
		"=ADDRESS(1,1,1,\"\")":              {"#VALUE!", "#VALUE!"},
		"=ADDRESS(1,1,0,TRUE)":              {"#NUM!", "#NUM!"},
		"=ADDRESS(1,16385,2,TRUE)":          {"#VALUE!", "#VALUE!"},
		"=ADDRESS(1,16385,3,TRUE)":          {"#VALUE!", "#VALUE!"},
		"=ADDRESS(1048577,1,1,TRUE)":        {"#VALUE!", "#VALUE!"},
		// CHOOSE
		"=CHOOSE()":                {"#VALUE!", "CHOOSE requires 2 arguments"},
		"=CHOOSE(\"index_num\",0)": {"#VALUE!", "CHOOSE requires first argument of type number"},
		"=CHOOSE(2,0)":             {"#VALUE!", "index_num should be <= to the number of values"},
		"=CHOOSE(1,NA())":          {"#N/A", "#N/A"},
		// COLUMN
		"=COLUMN(1,2)":                 {"#VALUE!", "COLUMN requires at most 1 argument"},
		"=COLUMN(\"\")":                {"#VALUE!", "invalid reference"},
		"=COLUMN(Sheet1)":              {"#NAME?", "invalid reference"},
		"=COLUMN(Sheet1!A1!B1)":        {"#NAME?", "invalid reference"},
		"=COLUMN(Sheet1!A1:Sheet2!A2)": {"#NAME?", "invalid reference"},
		"=COLUMN(Sheet1!A1:1A)":        {"#NAME?", "invalid reference"},
		// COLUMNS
		"=COLUMNS()":              {"#VALUE!", "COLUMNS requires 1 argument"},
		"=COLUMNS(1)":             {"#VALUE!", "invalid reference"},
		"=COLUMNS(\"\")":          {"#VALUE!", "invalid reference"},
		"=COLUMNS(Sheet1)":        {"#NAME?", "invalid reference"},
		"=COLUMNS(Sheet1!A1!B1)":  {"#NAME?", "invalid reference"},
		"=COLUMNS(Sheet1!Sheet1)": {"#NAME?", "invalid reference"},
		// FORMULATEXT
		"=FORMULATEXT()":  {"#VALUE!", "FORMULATEXT requires 1 argument"},
		"=FORMULATEXT(1)": {"#VALUE!", "#VALUE!"},
		// HLOOKUP
		"=HLOOKUP()":                     {"#VALUE!", "HLOOKUP requires at least 3 arguments"},
		"=HLOOKUP(D2,D1,1,FALSE)":        {"#VALUE!", "HLOOKUP requires second argument of table array"},
		"=HLOOKUP(D2,D:D,FALSE,FALSE)":   {"#VALUE!", "HLOOKUP requires numeric row argument"},
		"=HLOOKUP(D2,D:D,1,FALSE,FALSE)": {"#VALUE!", "HLOOKUP requires at most 4 arguments"},
		"=HLOOKUP(D2,D:D,1,2)":           {"#N/A", "HLOOKUP no result found"},
		"=HLOOKUP(D2,D10:D10,1,FALSE)":   {"#N/A", "HLOOKUP no result found"},
		"=HLOOKUP(D2,D2:D3,4,FALSE)":     {"#N/A", "HLOOKUP has invalid row index"},
		"=HLOOKUP(D2,C:C,1,FALSE)":       {"#N/A", "HLOOKUP no result found"},
		"=HLOOKUP(ISNUMBER(1),F3:F9,1)":  {"#N/A", "HLOOKUP no result found"},
		"=HLOOKUP(INT(1),E2:E9,1)":       {"#N/A", "HLOOKUP no result found"},
		"=HLOOKUP(MUNIT(2),MUNIT(3),1)":  {"#N/A", "HLOOKUP no result found"},
		"=HLOOKUP(A1:B2,B2:B3,1)":        {"#N/A", "HLOOKUP no result found"},
		// MATCH
		"=MATCH()":              {"#VALUE!", "MATCH requires 1 or 2 arguments"},
		"=MATCH(0,A1:A1,0,0)":   {"#VALUE!", "MATCH requires 1 or 2 arguments"},
		"=MATCH(0,A1:A1,\"x\")": {"#VALUE!", "MATCH requires numeric match_type argument"},
		"=MATCH(0,A1)":          {"#N/A", "MATCH arguments lookup_array should be one-dimensional array"},
		"=MATCH(0,A1:B2)":       {"#N/A", "MATCH arguments lookup_array should be one-dimensional array"},
		"=MATCH(0,A1:B1)":       {"#N/A", "#N/A"},
		// TRANSPOSE
		"=TRANSPOSE()": {"#VALUE!", "TRANSPOSE requires 1 argument"},
		// HYPERLINK
		"=HYPERLINK()": {"#VALUE!", "HYPERLINK requires at least 1 argument"},
		"=HYPERLINK(\"https://github.com/xuri/excelize\",\"Excelize\",\"\")": {"#VALUE!", "HYPERLINK allows at most 2 arguments"},
		// VLOOKUP
		"=VLOOKUP()":                     {"#VALUE!", "VLOOKUP requires at least 3 arguments"},
		"=VLOOKUP(D2,D1,1,FALSE)":        {"#VALUE!", "VLOOKUP requires second argument of table array"},
		"=VLOOKUP(D2,D:D,FALSE,FALSE)":   {"#VALUE!", "VLOOKUP requires numeric col argument"},
		"=VLOOKUP(D2,D:D,1,FALSE,FALSE)": {"#VALUE!", "VLOOKUP requires at most 4 arguments"},
		"=VLOOKUP(D2,D10:D10,1,FALSE)":   {"#N/A", "VLOOKUP no result found"},
		"=VLOOKUP(D2,D:D,2,FALSE)":       {"#N/A", "VLOOKUP has invalid column index"},
		"=VLOOKUP(D2,C:C,1,FALSE)":       {"#N/A", "VLOOKUP no result found"},
		"=VLOOKUP(ISNUMBER(1),F3:F9,1)":  {"#N/A", "VLOOKUP no result found"},
		"=VLOOKUP(INT(1),E2:E9,1)":       {"#N/A", "VLOOKUP no result found"},
		"=VLOOKUP(MUNIT(2),MUNIT(3),1)":  {"#N/A", "VLOOKUP no result found"},
		"=VLOOKUP(1,G1:H2,1,FALSE)":      {"#N/A", "VLOOKUP no result found"},
		// INDEX
		"=INDEX()":          {"#VALUE!", "INDEX requires 2 or 3 arguments"},
		"=INDEX(A1,2)":      {"#REF!", "INDEX row_num out of range"},
		"=INDEX(A1,0,2)":    {"#REF!", "INDEX col_num out of range"},
		"=INDEX(A1:A1,2)":   {"#REF!", "INDEX row_num out of range"},
		"=INDEX(A1:A1,0,2)": {"#REF!", "INDEX col_num out of range"},
		"=INDEX(A1:B2,2,3)": {"#REF!", "INDEX col_num out of range"},
		"=INDEX(A1:A2,0,0)": {"#VALUE!", "#VALUE!"},
		"=INDEX(0,\"\")":    {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=INDEX(0,0,\"\")":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// INDIRECT
		"=INDIRECT()":                     {"#VALUE!", "INDIRECT requires 1 or 2 arguments"},
		"=INDIRECT(\"E\"&1,TRUE,1)":       {"#VALUE!", "INDIRECT requires 1 or 2 arguments"},
		"=INDIRECT(\"R1048577C1\",\"\")":  {"#VALUE!", "#VALUE!"},
		"=INDIRECT(\"E1048577\")":         {"#REF!", "#REF!"},
		"=INDIRECT(\"R1048577C1\",FALSE)": {"#REF!", "#REF!"},
		"=INDIRECT(\"R1C16385\",FALSE)":   {"#REF!", "#REF!"},
		"=INDIRECT(\"\",FALSE)":           {"#REF!", "#REF!"},
		"=INDIRECT(\"R C1\",FALSE)":       {"#REF!", "#REF!"},
		"=INDIRECT(\"R1C \",FALSE)":       {"#REF!", "#REF!"},
		"=INDIRECT(\"R1C1:R2C \",FALSE)":  {"#REF!", "#REF!"},
		// LOOKUP
		"=LOOKUP()":                     {"#VALUE!", "LOOKUP requires at least 2 arguments"},
		"=LOOKUP(D2,D1,D2)":             {"#VALUE!", "LOOKUP requires second argument of table array"},
		"=LOOKUP(D2,D1,D2,FALSE)":       {"#VALUE!", "LOOKUP requires at most 3 arguments"},
		"=LOOKUP(1,MUNIT(0))":           {"#VALUE!", "LOOKUP requires not empty range as second argument"},
		"=LOOKUP(D1,MUNIT(1),MUNIT(1))": {"#N/A", "LOOKUP no result found"},
		// ROW
		"=ROW(1,2)":          {"#VALUE!", "ROW requires at most 1 argument"},
		"=ROW(\"\")":         {"#VALUE!", "invalid reference"},
		"=ROW(Sheet1)":       {"#NAME?", "invalid reference"},
		"=ROW(Sheet1!A1!B1)": {"#NAME?", "invalid reference"},
		// ROWS
		"=ROWS()":              {"#VALUE!", "ROWS requires 1 argument"},
		"=ROWS(1)":             {"#VALUE!", "invalid reference"},
		"=ROWS(\"\")":          {"#VALUE!", "invalid reference"},
		"=ROWS(Sheet1)":        {"#NAME?", "invalid reference"},
		"=ROWS(Sheet1!A1!B1)":  {"#NAME?", "invalid reference"},
		"=ROWS(Sheet1!Sheet1)": {"#NAME?", "invalid reference"},
		// Web Functions
		// ENCODEURL
		"=ENCODEURL()": {"#VALUE!", "ENCODEURL requires 1 argument"},
		// Financial Functions
		// ACCRINT
		"=ACCRINT()": {"#VALUE!", "ACCRINT requires at least 6 arguments"},
		"=ACCRINT(\"01/01/2012\",\"04/01/2012\",\"12/31/2013\",8%,10000,4,1,FALSE,0)":  {"#VALUE!", "ACCRINT allows at most 8 arguments"},
		"=ACCRINT(\"\",\"04/01/2012\",\"12/31/2013\",8%,10000,4,1,FALSE)":              {"#VALUE!", "#VALUE!"},
		"=ACCRINT(\"01/01/2012\",\"\",\"12/31/2013\",8%,10000,4,1,FALSE)":              {"#VALUE!", "#VALUE!"},
		"=ACCRINT(\"01/01/2012\",\"04/01/2012\",\"\",8%,10000,4,1,FALSE)":              {"#VALUE!", "#VALUE!"},
		"=ACCRINT(\"01/01/2012\",\"04/01/2012\",\"12/31/2013\",\"\",10000,4,1,FALSE)":  {"#NUM!", "#NUM!"},
		"=ACCRINT(\"01/01/2012\",\"04/01/2012\",\"12/31/2013\",8%,\"\",4,1,FALSE)":     {"#NUM!", "#NUM!"},
		"=ACCRINT(\"01/01/2012\",\"04/01/2012\",\"12/31/2013\",8%,10000,3)":            {"#NUM!", "#NUM!"},
		"=ACCRINT(\"01/01/2012\",\"04/01/2012\",\"12/31/2013\",8%,10000,\"\",1,FALSE)": {"#NUM!", "#NUM!"},
		"=ACCRINT(\"01/01/2012\",\"04/01/2012\",\"12/31/2013\",8%,10000,4,\"\",FALSE)": {"#NUM!", "#NUM!"},
		"=ACCRINT(\"01/01/2012\",\"04/01/2012\",\"12/31/2013\",8%,10000,4,1,\"\")":     {"#VALUE!", "#VALUE!"},
		"=ACCRINT(\"01/01/2012\",\"04/01/2012\",\"12/31/2013\",8%,10000,4,5,FALSE)":    {"#NUM!", "invalid basis"},
		// ACCRINTM
		"=ACCRINTM()": {"#VALUE!", "ACCRINTM requires 4 or 5 arguments"},
		"=ACCRINTM(\"\",\"01/01/2012\",8%,10000)":                {"#VALUE!", "#VALUE!"},
		"=ACCRINTM(\"01/01/2012\",\"\",8%,10000)":                {"#VALUE!", "#VALUE!"},
		"=ACCRINTM(\"12/31/2012\",\"01/01/2012\",8%,10000)":      {"#NUM!", "#NUM!"},
		"=ACCRINTM(\"01/01/2012\",\"12/31/2012\",\"\",10000)":    {"#NUM!", "#NUM!"},
		"=ACCRINTM(\"01/01/2012\",\"12/31/2012\",8%,\"\",10000)": {"#NUM!", "#NUM!"},
		"=ACCRINTM(\"01/01/2012\",\"12/31/2012\",8%,-1,10000)":   {"#NUM!", "#NUM!"},
		"=ACCRINTM(\"01/01/2012\",\"12/31/2012\",8%,10000,\"\")": {"#NUM!", "#NUM!"},
		"=ACCRINTM(\"01/01/2012\",\"12/31/2012\",8%,10000,5)":    {"#NUM!", "invalid basis"},
		// AMORDEGRC
		"=AMORDEGRC()": {"#VALUE!", "AMORDEGRC requires 6 or 7 arguments"},
		"=AMORDEGRC(\"\",\"01/01/2015\",\"09/30/2015\",20,1,20%)":     {"#VALUE!", "AMORDEGRC requires cost to be number argument"},
		"=AMORDEGRC(-1,\"01/01/2015\",\"09/30/2015\",20,1,20%)":       {"#VALUE!", "AMORDEGRC requires cost >= 0"},
		"=AMORDEGRC(150,\"\",\"09/30/2015\",20,1,20%)":                {"#VALUE!", "#VALUE!"},
		"=AMORDEGRC(150,\"01/01/2015\",\"\",20,1,20%)":                {"#VALUE!", "#VALUE!"},
		"=AMORDEGRC(150,\"09/30/2015\",\"01/01/2015\",20,1,20%)":      {"#NUM!", "#NUM!"},
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",\"\",1,20%)":    {"#NUM!", "#NUM!"},
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",-1,1,20%)":      {"#NUM!", "#NUM!"},
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",20,\"\",20%)":   {"#NUM!", "#NUM!"},
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",20,-1,20%)":     {"#NUM!", "#NUM!"},
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",20,1,\"\")":     {"#NUM!", "#NUM!"},
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",20,1,-1)":       {"#NUM!", "#NUM!"},
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",20,1,20%,\"\")": {"#NUM!", "#NUM!"},
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",20,1,50%)":      {"#NUM!", "AMORDEGRC requires rate to be < 0.5"},
		"=AMORDEGRC(150,\"01/01/2015\",\"09/30/2015\",20,1,20%,5)":    {"#NUM!", "invalid basis"},
		// AMORLINC
		"=AMORLINC()": {"#VALUE!", "AMORLINC requires 6 or 7 arguments"},
		"=AMORLINC(\"\",\"01/01/2015\",\"09/30/2015\",20,1,20%)":     {"#VALUE!", "AMORLINC requires cost to be number argument"},
		"=AMORLINC(-1,\"01/01/2015\",\"09/30/2015\",20,1,20%)":       {"#VALUE!", "AMORLINC requires cost >= 0"},
		"=AMORLINC(150,\"\",\"09/30/2015\",20,1,20%)":                {"#VALUE!", "#VALUE!"},
		"=AMORLINC(150,\"01/01/2015\",\"\",20,1,20%)":                {"#VALUE!", "#VALUE!"},
		"=AMORLINC(150,\"09/30/2015\",\"01/01/2015\",20,1,20%)":      {"#NUM!", "#NUM!"},
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",\"\",1,20%)":    {"#NUM!", "#NUM!"},
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",-1,1,20%)":      {"#NUM!", "#NUM!"},
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",20,\"\",20%)":   {"#NUM!", "#NUM!"},
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",20,-1,20%)":     {"#NUM!", "#NUM!"},
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",20,1,\"\")":     {"#NUM!", "#NUM!"},
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",20,1,-1)":       {"#NUM!", "#NUM!"},
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",20,1,20%,\"\")": {"#NUM!", "#NUM!"},
		"=AMORLINC(150,\"01/01/2015\",\"09/30/2015\",20,1,20%,5)":    {"#NUM!", "invalid basis"},
		// COUPDAYBS
		"=COUPDAYBS()":                                     {"#VALUE!", "COUPDAYBS requires 3 or 4 arguments"},
		"=COUPDAYBS(\"\",\"10/25/2012\",4)":                {"#VALUE!", "#VALUE!"},
		"=COUPDAYBS(\"01/01/2011\",\"\",4)":                {"#VALUE!", "#VALUE!"},
		"=COUPDAYBS(\"01/01/2011\",\"10/25/2012\",\"\")":   {"#VALUE!", "#VALUE!"},
		"=COUPDAYBS(\"01/01/2011\",\"10/25/2012\",4,\"\")": {"#NUM!", "#NUM!"},
		"=COUPDAYBS(\"10/25/2012\",\"01/01/2011\",4)":      {"#NUM!", "COUPDAYBS requires maturity > settlement"},
		// COUPDAYS
		"=COUPDAYS()":                                     {"#VALUE!", "COUPDAYS requires 3 or 4 arguments"},
		"=COUPDAYS(\"\",\"10/25/2012\",4)":                {"#VALUE!", "#VALUE!"},
		"=COUPDAYS(\"01/01/2011\",\"\",4)":                {"#VALUE!", "#VALUE!"},
		"=COUPDAYS(\"01/01/2011\",\"10/25/2012\",\"\")":   {"#VALUE!", "#VALUE!"},
		"=COUPDAYS(\"01/01/2011\",\"10/25/2012\",4,\"\")": {"#NUM!", "#NUM!"},
		"=COUPDAYS(\"10/25/2012\",\"01/01/2011\",4)":      {"#NUM!", "COUPDAYS requires maturity > settlement"},
		// COUPDAYSNC
		"=COUPDAYSNC()":                                     {"#VALUE!", "COUPDAYSNC requires 3 or 4 arguments"},
		"=COUPDAYSNC(\"\",\"10/25/2012\",4)":                {"#VALUE!", "#VALUE!"},
		"=COUPDAYSNC(\"01/01/2011\",\"\",4)":                {"#VALUE!", "#VALUE!"},
		"=COUPDAYSNC(\"01/01/2011\",\"10/25/2012\",\"\")":   {"#VALUE!", "#VALUE!"},
		"=COUPDAYSNC(\"01/01/2011\",\"10/25/2012\",4,\"\")": {"#NUM!", "#NUM!"},
		"=COUPDAYSNC(\"10/25/2012\",\"01/01/2011\",4)":      {"#NUM!", "COUPDAYSNC requires maturity > settlement"},
		// COUPNCD
		"=COUPNCD()": {"#VALUE!", "COUPNCD requires 3 or 4 arguments"},
		"=COUPNCD(\"01/01/2011\",\"10/25/2012\",4,0,0)":  {"#VALUE!", "COUPNCD requires 3 or 4 arguments"},
		"=COUPNCD(\"\",\"10/25/2012\",4)":                {"#VALUE!", "#VALUE!"},
		"=COUPNCD(\"01/01/2011\",\"\",4)":                {"#VALUE!", "#VALUE!"},
		"=COUPNCD(\"01/01/2011\",\"10/25/2012\",\"\")":   {"#VALUE!", "#VALUE!"},
		"=COUPNCD(\"01/01/2011\",\"10/25/2012\",4,\"\")": {"#NUM!", "#NUM!"},
		"=COUPNCD(\"01/01/2011\",\"10/25/2012\",3)":      {"#NUM!", "#NUM!"},
		"=COUPNCD(\"10/25/2012\",\"01/01/2011\",4)":      {"#NUM!", "COUPNCD requires maturity > settlement"},
		// COUPNUM
		"=COUPNUM()": {"#VALUE!", "COUPNUM requires 3 or 4 arguments"},
		"=COUPNUM(\"01/01/2011\",\"10/25/2012\",4,0,0)":  {"#VALUE!", "COUPNUM requires 3 or 4 arguments"},
		"=COUPNUM(\"\",\"10/25/2012\",4)":                {"#VALUE!", "#VALUE!"},
		"=COUPNUM(\"01/01/2011\",\"\",4)":                {"#VALUE!", "#VALUE!"},
		"=COUPNUM(\"01/01/2011\",\"10/25/2012\",\"\")":   {"#VALUE!", "#VALUE!"},
		"=COUPNUM(\"01/01/2011\",\"10/25/2012\",4,\"\")": {"#NUM!", "#NUM!"},
		"=COUPNUM(\"01/01/2011\",\"10/25/2012\",3)":      {"#NUM!", "#NUM!"},
		"=COUPNUM(\"10/25/2012\",\"01/01/2011\",4)":      {"#NUM!", "COUPNUM requires maturity > settlement"},
		// COUPPCD
		"=COUPPCD()": {"#VALUE!", "COUPPCD requires 3 or 4 arguments"},
		"=COUPPCD(\"01/01/2011\",\"10/25/2012\",4,0,0)":  {"#VALUE!", "COUPPCD requires 3 or 4 arguments"},
		"=COUPPCD(\"\",\"10/25/2012\",4)":                {"#VALUE!", "#VALUE!"},
		"=COUPPCD(\"01/01/2011\",\"\",4)":                {"#VALUE!", "#VALUE!"},
		"=COUPPCD(\"01/01/2011\",\"10/25/2012\",\"\")":   {"#VALUE!", "#VALUE!"},
		"=COUPPCD(\"01/01/2011\",\"10/25/2012\",4,\"\")": {"#NUM!", "#NUM!"},
		"=COUPPCD(\"01/01/2011\",\"10/25/2012\",3)":      {"#NUM!", "#NUM!"},
		"=COUPPCD(\"10/25/2012\",\"01/01/2011\",4)":      {"#NUM!", "COUPPCD requires maturity > settlement"},
		// CUMIPMT
		"=CUMIPMT()":               {"#VALUE!", "CUMIPMT requires 6 arguments"},
		"=CUMIPMT(0,0,0,0,0,2)":    {"#N/A", "#N/A"},
		"=CUMIPMT(0,0,0,-1,0,0)":   {"#N/A", "#N/A"},
		"=CUMIPMT(0,0,0,1,0,0)":    {"#N/A", "#N/A"},
		"=CUMIPMT(\"\",0,0,0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CUMIPMT(0,\"\",0,0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CUMIPMT(0,0,\"\",0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CUMIPMT(0,0,0,\"\",0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CUMIPMT(0,0,0,0,\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CUMIPMT(0,0,0,0,0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// CUMPRINC
		"=CUMPRINC()":               {"#VALUE!", "CUMPRINC requires 6 arguments"},
		"=CUMPRINC(0,0,0,0,0,2)":    {"#N/A", "#N/A"},
		"=CUMPRINC(0,0,0,-1,0,0)":   {"#N/A", "#N/A"},
		"=CUMPRINC(0,0,0,1,0,0)":    {"#N/A", "#N/A"},
		"=CUMPRINC(\"\",0,0,0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CUMPRINC(0,\"\",0,0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CUMPRINC(0,0,\"\",0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CUMPRINC(0,0,0,\"\",0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CUMPRINC(0,0,0,0,\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=CUMPRINC(0,0,0,0,0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// DB
		"=DB()":             {"#VALUE!", "DB requires at least 4 arguments"},
		"=DB(0,0,0,0,0,0)":  {"#VALUE!", "DB allows at most 5 arguments"},
		"=DB(-1,0,0,0)":     {"#N/A", "#N/A"},
		"=DB(\"\",0,0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DB(0,\"\",0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DB(0,0,\"\",0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DB(0,0,0,\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DB(0,0,0,0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// DDB
		"=DDB()":             {"#VALUE!", "DDB requires at least 4 arguments"},
		"=DDB(0,0,0,0,0,0)":  {"#VALUE!", "DDB allows at most 5 arguments"},
		"=DDB(-1,0,0,0)":     {"#N/A", "#N/A"},
		"=DDB(\"\",0,0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DDB(0,\"\",0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DDB(0,0,\"\",0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DDB(0,0,0,\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DDB(0,0,0,0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// DISC
		"=DISC()":                                          {"#VALUE!", "DISC requires 4 or 5 arguments"},
		"=DISC(\"\",\"03/31/2021\",95,100)":                {"#VALUE!", "#VALUE!"},
		"=DISC(\"04/01/2016\",\"\",95,100)":                {"#VALUE!", "#VALUE!"},
		"=DISC(\"04/01/2016\",\"03/31/2021\",\"\",100)":    {"#VALUE!", "#VALUE!"},
		"=DISC(\"04/01/2016\",\"03/31/2021\",95,\"\")":     {"#VALUE!", "#VALUE!"},
		"=DISC(\"04/01/2016\",\"03/31/2021\",95,100,\"\")": {"#NUM!", "#NUM!"},
		"=DISC(\"03/31/2021\",\"04/01/2016\",95,100)":      {"#NUM!", "DISC requires maturity > settlement"},
		"=DISC(\"04/01/2016\",\"03/31/2021\",0,100)":       {"#NUM!", "DISC requires pr > 0"},
		"=DISC(\"04/01/2016\",\"03/31/2021\",95,0)":        {"#NUM!", "DISC requires redemption > 0"},
		"=DISC(\"04/01/2016\",\"03/31/2021\",95,100,5)":    {"#NUM!", "invalid basis"},
		// DOLLAR
		"DOLLAR()":       {"#VALUE!", "DOLLAR requires at least 1 argument"},
		"DOLLAR(0,0,0)":  {"#VALUE!", "DOLLAR requires 1 or 2 arguments"},
		"DOLLAR(\"\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"DOLLAR(0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"DOLLAR(1,200)":  {"#VALUE!", "decimal value should be less than 128"},
		// DOLLARDE
		"=DOLLARDE()":       {"#VALUE!", "DOLLARDE requires 2 arguments"},
		"=DOLLARDE(\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DOLLARDE(0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DOLLARDE(0,-1)":   {"#NUM!", "#NUM!"},
		"=DOLLARDE(0,0)":    {"#DIV/0!", "#DIV/0!"},
		// DOLLARFR
		"=DOLLARFR()":       {"#VALUE!", "DOLLARFR requires 2 arguments"},
		"=DOLLARFR(\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DOLLARFR(0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DOLLARFR(0,-1)":   {"#NUM!", "#NUM!"},
		"=DOLLARFR(0,0)":    {"#DIV/0!", "#DIV/0!"},
		// DURATION
		"=DURATION()": {"#VALUE!", "DURATION requires 5 or 6 arguments"},
		"=DURATION(\"\",\"03/31/2025\",10%,8%,4)":                {"#VALUE!", "#VALUE!"},
		"=DURATION(\"04/01/2015\",\"\",10%,8%,4)":                {"#VALUE!", "#VALUE!"},
		"=DURATION(\"03/31/2025\",\"04/01/2015\",10%,8%,4)":      {"#NUM!", "DURATION requires maturity > settlement"},
		"=DURATION(\"04/01/2015\",\"03/31/2025\",-1,8%,4)":       {"#NUM!", "DURATION requires coupon >= 0"},
		"=DURATION(\"04/01/2015\",\"03/31/2025\",10%,-1,4)":      {"#NUM!", "DURATION requires yld >= 0"},
		"=DURATION(\"04/01/2015\",\"03/31/2025\",\"\",8%,4)":     {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DURATION(\"04/01/2015\",\"03/31/2025\",10%,\"\",4)":    {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DURATION(\"04/01/2015\",\"03/31/2025\",10%,8%,\"\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=DURATION(\"04/01/2015\",\"03/31/2025\",10%,8%,3)":      {"#NUM!", "#NUM!"},
		"=DURATION(\"04/01/2015\",\"03/31/2025\",10%,8%,4,\"\")": {"#NUM!", "#NUM!"},
		"=DURATION(\"04/01/2015\",\"03/31/2025\",10%,8%,4,5)":    {"#NUM!", "invalid basis"},
		// EFFECT
		"=EFFECT()":       {"#VALUE!", "EFFECT requires 2 arguments"},
		"=EFFECT(\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=EFFECT(0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=EFFECT(0,0)":    {"#NUM!", "#NUM!"},
		"=EFFECT(1,0)":    {"#NUM!", "#NUM!"},
		// EUROCONVERT
		"=EUROCONVERT()": {"#VALUE!", "EUROCONVERT requires at least 3 arguments"},
		"=EUROCONVERT(1.47,\"FRF\",\"DEM\",TRUE,3,1)":  {"#VALUE!", "EUROCONVERT allows at most 5 arguments"},
		"=EUROCONVERT(\"\",\"FRF\",\"DEM\",TRUE,3)":    {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=EUROCONVERT(1.47,\"FRF\",\"DEM\",\"\",3)":    {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=EUROCONVERT(1.47,\"FRF\",\"DEM\",TRUE,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=EUROCONVERT(1.47,\"\",\"DEM\")":              {"#VALUE!", "#VALUE!"},
		"=EUROCONVERT(1.47,\"FRF\",\"\",TRUE,3)":       {"#VALUE!", "#VALUE!"},
		// FV
		"=FV()":              {"#VALUE!", "FV requires at least 3 arguments"},
		"=FV(0,0,0,0,0,0,0)": {"#VALUE!", "FV allows at most 5 arguments"},
		"=FV(0,0,0,0,2)":     {"#N/A", "#N/A"},
		"=FV(\"\",0,0,0,0)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=FV(0,\"\",0,0,0)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=FV(0,0,\"\",0,0)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=FV(0,0,0,\"\",0)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=FV(0,0,0,0,\"\")":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// FVSCHEDULE
		"=FVSCHEDULE()":        {"#VALUE!", "FVSCHEDULE requires 2 arguments"},
		"=FVSCHEDULE(\"\",0)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=FVSCHEDULE(0,\"x\")": {"#VALUE!", "strconv.ParseFloat: parsing \"x\": invalid syntax"},
		// INTRATE
		"=INTRATE()":                                          {"#VALUE!", "INTRATE requires 4 or 5 arguments"},
		"=INTRATE(\"\",\"03/31/2021\",95,100)":                {"#VALUE!", "#VALUE!"},
		"=INTRATE(\"04/01/2016\",\"\",95,100)":                {"#VALUE!", "#VALUE!"},
		"=INTRATE(\"04/01/2016\",\"03/31/2021\",\"\",100)":    {"#VALUE!", "#VALUE!"},
		"=INTRATE(\"04/01/2016\",\"03/31/2021\",95,\"\")":     {"#VALUE!", "#VALUE!"},
		"=INTRATE(\"04/01/2016\",\"03/31/2021\",95,100,\"\")": {"#NUM!", "#NUM!"},
		"=INTRATE(\"03/31/2021\",\"04/01/2016\",95,100)":      {"#NUM!", "INTRATE requires maturity > settlement"},
		"=INTRATE(\"04/01/2016\",\"03/31/2021\",0,100)":       {"#NUM!", "INTRATE requires investment > 0"},
		"=INTRATE(\"04/01/2016\",\"03/31/2021\",95,0)":        {"#NUM!", "INTRATE requires redemption > 0"},
		"=INTRATE(\"04/01/2016\",\"03/31/2021\",95,100,5)":    {"#NUM!", "invalid basis"},
		// IPMT
		"=IPMT()":               {"#VALUE!", "IPMT requires at least 4 arguments"},
		"=IPMT(0,0,0,0,0,0,0)":  {"#VALUE!", "IPMT allows at most 6 arguments"},
		"=IPMT(0,0,0,0,0,2)":    {"#N/A", "#N/A"},
		"=IPMT(0,-1,0,0,0,0)":   {"#N/A", "#N/A"},
		"=IPMT(0,1,0,0,0,0)":    {"#N/A", "#N/A"},
		"=IPMT(\"\",0,0,0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=IPMT(0,\"\",0,0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=IPMT(0,0,\"\",0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=IPMT(0,0,0,\"\",0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=IPMT(0,0,0,0,\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=IPMT(0,0,0,0,0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// ISPMT
		"=ISPMT()":           {"#VALUE!", "ISPMT requires 4 arguments"},
		"=ISPMT(\"\",0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ISPMT(0,\"\",0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ISPMT(0,0,\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ISPMT(0,0,0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// MDURATION
		"=MDURATION()": {"#VALUE!", "MDURATION requires 5 or 6 arguments"},
		"=MDURATION(\"\",\"03/31/2025\",10%,8%,4)":                {"#VALUE!", "#VALUE!"},
		"=MDURATION(\"04/01/2015\",\"\",10%,8%,4)":                {"#VALUE!", "#VALUE!"},
		"=MDURATION(\"03/31/2025\",\"04/01/2015\",10%,8%,4)":      {"#NUM!", "MDURATION requires maturity > settlement"},
		"=MDURATION(\"04/01/2015\",\"03/31/2025\",-1,8%,4)":       {"#NUM!", "MDURATION requires coupon >= 0"},
		"=MDURATION(\"04/01/2015\",\"03/31/2025\",10%,-1,4)":      {"#NUM!", "MDURATION requires yld >= 0"},
		"=MDURATION(\"04/01/2015\",\"03/31/2025\",\"\",8%,4)":     {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=MDURATION(\"04/01/2015\",\"03/31/2025\",10%,\"\",4)":    {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=MDURATION(\"04/01/2015\",\"03/31/2025\",10%,8%,\"\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=MDURATION(\"04/01/2015\",\"03/31/2025\",10%,8%,3)":      {"#NUM!", "#NUM!"},
		"=MDURATION(\"04/01/2015\",\"03/31/2025\",10%,8%,4,\"\")": {"#NUM!", "#NUM!"},
		"=MDURATION(\"04/01/2015\",\"03/31/2025\",10%,8%,4,5)":    {"#NUM!", "invalid basis"},
		// NOMINAL
		"=NOMINAL()":       {"#VALUE!", "NOMINAL requires 2 arguments"},
		"=NOMINAL(\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=NOMINAL(0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=NOMINAL(0,0)":    {"#NUM!", "#NUM!"},
		"=NOMINAL(1,0)":    {"#NUM!", "#NUM!"},
		// NPER
		"=NPER()":             {"#VALUE!", "NPER requires at least 3 arguments"},
		"=NPER(0,0,0,0,0,0)":  {"#VALUE!", "NPER allows at most 5 arguments"},
		"=NPER(0,0,0)":        {"#NUM!", "#NUM!"},
		"=NPER(0,0,0,0,2)":    {"#N/A", "#N/A"},
		"=NPER(\"\",0,0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=NPER(0,\"\",0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=NPER(0,0,\"\",0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=NPER(0,0,0,\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=NPER(0,0,0,0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// NPV
		"=NPV()":       {"#VALUE!", "NPV requires at least 2 arguments"},
		"=NPV(\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// ODDFPRICE
		"=ODDFPRICE()": {"#VALUE!", "ODDFPRICE requires 8 or 9 arguments"},
		"=ODDFPRICE(\"\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,100,2)":                {"#VALUE!", "#VALUE!"},
		"=ODDFPRICE(\"02/01/2017\",\"\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,100,2)":                {"#VALUE!", "#VALUE!"},
		"=ODDFPRICE(\"02/01/2017\",\"03/31/2021\",\"\",\"03/31/2017\",5.5%,3.5%,100,2)":                {"#VALUE!", "#VALUE!"},
		"=ODDFPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"\",5.5%,3.5%,100,2)":                {"#VALUE!", "#VALUE!"},
		"=ODDFPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",\"\",3.5%,100,2)":      {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ODDFPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,\"\",100,2)":      {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ODDFPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,\"\",2)":     {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ODDFPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,100,\"\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ODDFPRICE(\"02/01/2017\",\"03/31/2021\",\"02/01/2017\",\"03/31/2017\",5.5%,3.5%,100,2)":      {"#NUM!", "ODDFPRICE requires settlement > issue"},
		"=ODDFPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"02/01/2017\",5.5%,3.5%,100,2)":      {"#NUM!", "ODDFPRICE requires first_coupon > settlement"},
		"=ODDFPRICE(\"02/01/2017\",\"02/01/2017\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,100,2)":      {"#NUM!", "ODDFPRICE requires maturity > first_coupon"},
		"=ODDFPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",-1,3.5%,100,2)":        {"#NUM!", "ODDFPRICE requires rate >= 0"},
		"=ODDFPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,-1,100,2)":        {"#NUM!", "ODDFPRICE requires yld >= 0"},
		"=ODDFPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,0,2)":        {"#NUM!", "ODDFPRICE requires redemption > 0"},
		"=ODDFPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,100,2,\"\")": {"#NUM!", "#NUM!"},
		"=ODDFPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,100,3)":      {"#NUM!", "#NUM!"},
		"=ODDFPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/30/2017\",5.5%,3.5%,100,4)":      {"#NUM!", "#NUM!"},
		"=ODDFPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,100,2,5)":    {"#NUM!", "invalid basis"},
		// ODDFYIELD
		"=ODDFYIELD()": {"#VALUE!", "ODDFYIELD requires 8 or 9 arguments"},
		"=ODDFYIELD(\"\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,100,2)":                {"#VALUE!", "#VALUE!"},
		"=ODDFYIELD(\"02/01/2017\",\"\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,100,2)":                {"#VALUE!", "#VALUE!"},
		"=ODDFYIELD(\"02/01/2017\",\"03/31/2021\",\"\",\"03/31/2017\",5.5%,3.5%,100,2)":                {"#VALUE!", "#VALUE!"},
		"=ODDFYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"\",5.5%,3.5%,100,2)":                {"#VALUE!", "#VALUE!"},
		"=ODDFYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",\"\",3.5%,100,2)":      {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ODDFYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,\"\",100,2)":      {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ODDFYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,\"\",2)":     {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ODDFYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,100,\"\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ODDFYIELD(\"02/01/2017\",\"03/31/2021\",\"02/01/2017\",\"03/31/2017\",5.5%,3.5%,100,2)":      {"#NUM!", "ODDFYIELD requires settlement > issue"},
		"=ODDFYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"02/01/2017\",5.5%,3.5%,100,2)":      {"#NUM!", "ODDFYIELD requires first_coupon > settlement"},
		"=ODDFYIELD(\"02/01/2017\",\"02/01/2017\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,100,2)":      {"#NUM!", "ODDFYIELD requires maturity > first_coupon"},
		"=ODDFYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",-1,3.5%,100,2)":        {"#NUM!", "ODDFYIELD requires rate >= 0"},
		"=ODDFYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,0,100,2)":         {"#NUM!", "ODDFYIELD requires pr > 0"},
		"=ODDFYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,0,2)":        {"#NUM!", "ODDFYIELD requires redemption > 0"},
		"=ODDFYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,100,2,\"\")": {"#NUM!", "#NUM!"},
		"=ODDFYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,100,3)":      {"#NUM!", "#NUM!"},
		"=ODDFYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/30/2017\",5.5%,3.5%,100,4)":      {"#NUM!", "#NUM!"},
		"=ODDFYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"03/31/2017\",5.5%,3.5%,100,2,5)":    {"#NUM!", "invalid basis"},
		// ODDLPRICE
		"=ODDLPRICE()": {"#VALUE!", "ODDLPRICE requires 7 or 8 arguments"},
		"=ODDLPRICE(\"\",\"03/31/2021\",\"12/01/2016\",5.5%,3.5%,100,2)":                {"#VALUE!", "#VALUE!"},
		"=ODDLPRICE(\"02/01/2017\",\"\",\"12/01/2016\",5.5%,3.5%,100,2)":                {"#VALUE!", "#VALUE!"},
		"=ODDLPRICE(\"02/01/2017\",\"03/31/2021\",\"\",5.5%,3.5%,100,2)":                {"#VALUE!", "#VALUE!"},
		"=ODDLPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"\",3.5%,100,2)":      {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ODDLPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",5.5%,\"\",100,2)":      {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ODDLPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",5.5%,3.5%,\"\",2)":     {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ODDLPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",5.5%,3.5%,100,\"\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ODDLPRICE(\"04/20/2008\",\"06/15/2008\",\"04/30/2008\",3.75%,99.875,100,2)":   {"#NUM!", "ODDLPRICE requires settlement > last_interest"},
		"=ODDLPRICE(\"06/20/2008\",\"06/15/2008\",\"04/30/2008\",3.75%,99.875,100,2)":   {"#NUM!", "ODDLPRICE requires maturity > settlement"},
		"=ODDLPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",-1,3.5%,100,2)":        {"#NUM!", "ODDLPRICE requires rate >= 0"},
		"=ODDLPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",5.5%,-1,100,2)":        {"#NUM!", "ODDLPRICE requires yld >= 0"},
		"=ODDLPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",5.5%,3.5%,0,2)":        {"#NUM!", "ODDLPRICE requires redemption > 0"},
		"=ODDLPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",5.5%,3.5%,100,2,\"\")": {"#NUM!", "#NUM!"},
		"=ODDLPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",5.5%,3.5%,100,3)":      {"#NUM!", "#NUM!"},
		"=ODDLPRICE(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",5.5%,3.5%,100,2,5)":    {"#NUM!", "invalid basis"},
		// ODDLYIELD
		"=ODDLYIELD()": {"#VALUE!", "ODDLYIELD requires 7 or 8 arguments"},
		"=ODDLYIELD(\"\",\"03/31/2021\",\"12/01/2016\",5.5%,3.5%,100,2)":                {"#VALUE!", "#VALUE!"},
		"=ODDLYIELD(\"02/01/2017\",\"\",\"12/01/2016\",5.5%,3.5%,100,2)":                {"#VALUE!", "#VALUE!"},
		"=ODDLYIELD(\"02/01/2017\",\"03/31/2021\",\"\",5.5%,3.5%,100,2)":                {"#VALUE!", "#VALUE!"},
		"=ODDLYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",\"\",3.5%,100,2)":      {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ODDLYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",5.5%,\"\",100,2)":      {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ODDLYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",5.5%,3.5%,\"\",2)":     {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ODDLYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",5.5%,3.5%,100,\"\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=ODDLYIELD(\"04/20/2008\",\"06/15/2008\",\"04/30/2008\",3.75%,99.875,100,2)":   {"#NUM!", "ODDLYIELD requires settlement > last_interest"},
		"=ODDLYIELD(\"06/20/2008\",\"06/15/2008\",\"04/30/2008\",3.75%,99.875,100,2)":   {"#NUM!", "ODDLYIELD requires maturity > settlement"},
		"=ODDLYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",-1,3.5%,100,2)":        {"#NUM!", "ODDLYIELD requires rate >= 0"},
		"=ODDLYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",5.5%,0,100,2)":         {"#NUM!", "ODDLYIELD requires pr > 0"},
		"=ODDLYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",5.5%,3.5%,0,2)":        {"#NUM!", "ODDLYIELD requires redemption > 0"},
		"=ODDLYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",5.5%,3.5%,100,2,\"\")": {"#NUM!", "#NUM!"},
		"=ODDLYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",5.5%,3.5%,100,3)":      {"#NUM!", "#NUM!"},
		"=ODDLYIELD(\"02/01/2017\",\"03/31/2021\",\"12/01/2016\",5.5%,3.5%,100,2,5)":    {"#NUM!", "invalid basis"},
		// PDURATION
		"=PDURATION()":         {"#VALUE!", "PDURATION requires 3 arguments"},
		"=PDURATION(\"\",0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PDURATION(0,\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PDURATION(0,0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PDURATION(0,0,0)":    {"#NUM!", "#NUM!"},
		// PMT
		"=PMT()":             {"#VALUE!", "PMT requires at least 3 arguments"},
		"=PMT(0,0,0,0,0,0)":  {"#VALUE!", "PMT allows at most 5 arguments"},
		"=PMT(0,0,0,0,2)":    {"#N/A", "#N/A"},
		"=PMT(\"\",0,0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PMT(0,\"\",0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PMT(0,0,\"\",0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PMT(0,0,0,\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PMT(0,0,0,0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// PRICE
		"=PRICE()": {"#VALUE!", "PRICE requires 6 or 7 arguments"},
		"=PRICE(\"\",\"02/01/2020\",12%,10%,100,2,4)":              {"#VALUE!", "#VALUE!"},
		"=PRICE(\"04/01/2012\",\"\",12%,10%,100,2,4)":              {"#VALUE!", "#VALUE!"},
		"=PRICE(\"04/01/2012\",\"02/01/2020\",\"\",10%,100,2,4)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PRICE(\"04/01/2012\",\"02/01/2020\",12%,\"\",100,2,4)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PRICE(\"04/01/2012\",\"02/01/2020\",12%,10%,\"\",2,4)":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PRICE(\"04/01/2012\",\"02/01/2020\",12%,10%,100,\"\",4)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PRICE(\"04/01/2012\",\"02/01/2020\",-1,10%,100,2,4)":     {"#NUM!", "PRICE requires rate >= 0"},
		"=PRICE(\"04/01/2012\",\"02/01/2020\",12%,-1,100,2,4)":     {"#NUM!", "PRICE requires yld >= 0"},
		"=PRICE(\"04/01/2012\",\"02/01/2020\",12%,10%,0,2,4)":      {"#NUM!", "PRICE requires redemption > 0"},
		"=PRICE(\"04/01/2012\",\"02/01/2020\",12%,10%,100,2,\"\")": {"#NUM!", "#NUM!"},
		"=PRICE(\"04/01/2012\",\"02/01/2020\",12%,10%,100,3,4)":    {"#NUM!", "#NUM!"},
		"=PRICE(\"04/01/2012\",\"02/01/2020\",12%,10%,100,2,5)":    {"#NUM!", "invalid basis"},
		// PPMT
		"=PPMT()":               {"#VALUE!", "PPMT requires at least 4 arguments"},
		"=PPMT(0,0,0,0,0,0,0)":  {"#VALUE!", "PPMT allows at most 6 arguments"},
		"=PPMT(0,0,0,0,0,2)":    {"#N/A", "#N/A"},
		"=PPMT(0,-1,0,0,0,0)":   {"#N/A", "#N/A"},
		"=PPMT(0,1,0,0,0,0)":    {"#N/A", "#N/A"},
		"=PPMT(\"\",0,0,0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PPMT(0,\"\",0,0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PPMT(0,0,\"\",0,0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PPMT(0,0,0,\"\",0,0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PPMT(0,0,0,0,\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PPMT(0,0,0,0,0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// PRICEDISC
		"=PRICEDISC()":                                          {"#VALUE!", "PRICEDISC requires 4 or 5 arguments"},
		"=PRICEDISC(\"\",\"03/31/2021\",95,100)":                {"#VALUE!", "#VALUE!"},
		"=PRICEDISC(\"04/01/2016\",\"\",95,100)":                {"#VALUE!", "#VALUE!"},
		"=PRICEDISC(\"04/01/2016\",\"03/31/2021\",\"\",100)":    {"#VALUE!", "#VALUE!"},
		"=PRICEDISC(\"04/01/2016\",\"03/31/2021\",95,\"\")":     {"#VALUE!", "#VALUE!"},
		"=PRICEDISC(\"04/01/2016\",\"03/31/2021\",95,100,\"\")": {"#NUM!", "#NUM!"},
		"=PRICEDISC(\"03/31/2021\",\"04/01/2016\",95,100)":      {"#NUM!", "PRICEDISC requires maturity > settlement"},
		"=PRICEDISC(\"04/01/2016\",\"03/31/2021\",0,100)":       {"#NUM!", "PRICEDISC requires discount > 0"},
		"=PRICEDISC(\"04/01/2016\",\"03/31/2021\",95,0)":        {"#NUM!", "PRICEDISC requires redemption > 0"},
		"=PRICEDISC(\"04/01/2016\",\"03/31/2021\",95,100,5)":    {"#NUM!", "invalid basis"},
		// PRICEMAT
		"=PRICEMAT()": {"#VALUE!", "PRICEMAT requires 5 or 6 arguments"},
		"=PRICEMAT(\"\",\"03/31/2021\",\"01/01/2017\",4.5%,2.5%)":                {"#VALUE!", "#VALUE!"},
		"=PRICEMAT(\"04/01/2017\",\"\",\"01/01/2017\",4.5%,2.5%)":                {"#VALUE!", "#VALUE!"},
		"=PRICEMAT(\"04/01/2017\",\"03/31/2021\",\"\",4.5%,2.5%)":                {"#VALUE!", "#VALUE!"},
		"=PRICEMAT(\"04/01/2017\",\"03/31/2021\",\"01/01/2017\",\"\",2.5%)":      {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PRICEMAT(\"04/01/2017\",\"03/31/2021\",\"01/01/2017\",4.5%,\"\")":      {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PRICEMAT(\"04/01/2017\",\"03/31/2021\",\"01/01/2017\",4.5%,2.5%,\"\")": {"#NUM!", "#NUM!"},
		"=PRICEMAT(\"03/31/2021\",\"04/01/2017\",\"01/01/2017\",4.5%,2.5%)":      {"#NUM!", "PRICEMAT requires maturity > settlement"},
		"=PRICEMAT(\"01/01/2017\",\"03/31/2021\",\"04/01/2017\",4.5%,2.5%)":      {"#NUM!", "PRICEMAT requires settlement > issue"},
		"=PRICEMAT(\"04/01/2017\",\"03/31/2021\",\"01/01/2017\",-1,2.5%)":        {"#NUM!", "PRICEMAT requires rate >= 0"},
		"=PRICEMAT(\"04/01/2017\",\"03/31/2021\",\"01/01/2017\",4.5%,-1)":        {"#NUM!", "PRICEMAT requires yld >= 0"},
		"=PRICEMAT(\"04/01/2017\",\"03/31/2021\",\"01/01/2017\",4.5%,2.5%,5)":    {"#NUM!", "invalid basis"},
		// PV
		"=PV()":                     {"#VALUE!", "PV requires at least 3 arguments"},
		"=PV(10%/4,16,2000,0,1,0)":  {"#VALUE!", "PV allows at most 5 arguments"},
		"=PV(\"\",16,2000,0,1)":     {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PV(10%/4,\"\",2000,0,1)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PV(10%/4,16,\"\",0,1)":    {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PV(10%/4,16,2000,\"\",1)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=PV(10%/4,16,2000,0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// RATE
		"=RATE()":                        {"#VALUE!", "RATE requires at least 3 arguments"},
		"=RATE(48,-200,8000,3,1,0.5,0)":  {"#VALUE!", "RATE allows at most 6 arguments"},
		"=RATE(\"\",-200,8000,3,1,0.5)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=RATE(48,\"\",8000,3,1,0.5)":    {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=RATE(48,-200,\"\",3,1,0.5)":    {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=RATE(48,-200,8000,\"\",1,0.5)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=RATE(48,-200,8000,3,\"\",0.5)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=RATE(48,-200,8000,3,1,\"\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		// RECEIVED
		"=RECEIVED()": {"#VALUE!", "RECEIVED requires at least 4 arguments"},
		"=RECEIVED(\"04/01/2011\",\"03/31/2016\",1000,4.5%,1,0)":  {"#VALUE!", "RECEIVED allows at most 5 arguments"},
		"=RECEIVED(\"\",\"03/31/2016\",1000,4.5%,1)":              {"#VALUE!", "#VALUE!"},
		"=RECEIVED(\"04/01/2011\",\"\",1000,4.5%,1)":              {"#VALUE!", "#VALUE!"},
		"=RECEIVED(\"04/01/2011\",\"03/31/2016\",\"\",4.5%,1)":    {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=RECEIVED(\"04/01/2011\",\"03/31/2016\",1000,\"\",1)":    {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=RECEIVED(\"04/01/2011\",\"03/31/2016\",1000,4.5%,\"\")": {"#NUM!", "#NUM!"},
		"=RECEIVED(\"04/01/2011\",\"03/31/2016\",1000,0)":         {"#NUM!", "RECEIVED requires discount > 0"},
		"=RECEIVED(\"04/01/2011\",\"03/31/2016\",1000,4.5%,5)":    {"#NUM!", "invalid basis"},
		// RRI
		"=RRI()":               {"#VALUE!", "RRI requires 3 arguments"},
		"=RRI(\"\",\"\",\"\")": {"#NUM!", "#NUM!"},
		"=RRI(0,10000,15000)":  {"#NUM!", "RRI requires nper argument to be > 0"},
		"=RRI(10,0,15000)":     {"#NUM!", "RRI requires pv argument to be > 0"},
		"=RRI(10,10000,-1)":    {"#NUM!", "RRI requires fv argument to be >= 0"},
		// SLN
		"=SLN()":               {"#VALUE!", "SLN requires 3 arguments"},
		"=SLN(\"\",\"\",\"\")": {"#NUM!", "#NUM!"},
		"=SLN(10000,1000,0)":   {"#NUM!", "SLN requires life argument to be > 0"},
		// SYD
		"=SYD()":                    {"#VALUE!", "SYD requires 4 arguments"},
		"=SYD(\"\",\"\",\"\",\"\")": {"#NUM!", "#NUM!"},
		"=SYD(10000,1000,0,1)":      {"#NUM!", "SYD requires life argument to be > 0"},
		"=SYD(10000,1000,5,0)":      {"#NUM!", "SYD requires per argument to be > 0"},
		"=SYD(10000,1000,1,5)":      {"#NUM!", "#NUM!"},
		// TBILLEQ
		"=TBILLEQ()":                                   {"#VALUE!", "TBILLEQ requires 3 arguments"},
		"=TBILLEQ(\"\",\"06/30/2017\",2.5%)":           {"#VALUE!", "#VALUE!"},
		"=TBILLEQ(\"01/01/2017\",\"\",2.5%)":           {"#VALUE!", "#VALUE!"},
		"=TBILLEQ(\"01/01/2017\",\"06/30/2017\",\"\")": {"#VALUE!", "#VALUE!"},
		"=TBILLEQ(\"01/01/2017\",\"06/30/2017\",0)":    {"#NUM!", "#NUM!"},
		"=TBILLEQ(\"01/01/2017\",\"06/30/2018\",2.5%)": {"#NUM!", "#NUM!"},
		"=TBILLEQ(\"06/30/2017\",\"01/01/2017\",2.5%)": {"#NUM!", "#NUM!"},
		// TBILLPRICE
		"=TBILLPRICE()":                                   {"#VALUE!", "TBILLPRICE requires 3 arguments"},
		"=TBILLPRICE(\"\",\"06/30/2017\",2.5%)":           {"#VALUE!", "#VALUE!"},
		"=TBILLPRICE(\"01/01/2017\",\"\",2.5%)":           {"#VALUE!", "#VALUE!"},
		"=TBILLPRICE(\"01/01/2017\",\"06/30/2017\",\"\")": {"#VALUE!", "#VALUE!"},
		"=TBILLPRICE(\"01/01/2017\",\"06/30/2017\",0)":    {"#NUM!", "#NUM!"},
		"=TBILLPRICE(\"01/01/2017\",\"06/30/2018\",2.5%)": {"#NUM!", "#NUM!"},
		"=TBILLPRICE(\"06/30/2017\",\"01/01/2017\",2.5%)": {"#NUM!", "#NUM!"},
		// TBILLYIELD
		"=TBILLYIELD()":                                   {"#VALUE!", "TBILLYIELD requires 3 arguments"},
		"=TBILLYIELD(\"\",\"06/30/2017\",2.5%)":           {"#VALUE!", "#VALUE!"},
		"=TBILLYIELD(\"01/01/2017\",\"\",2.5%)":           {"#VALUE!", "#VALUE!"},
		"=TBILLYIELD(\"01/01/2017\",\"06/30/2017\",\"\")": {"#VALUE!", "#VALUE!"},
		"=TBILLYIELD(\"01/01/2017\",\"06/30/2017\",0)":    {"#NUM!", "#NUM!"},
		"=TBILLYIELD(\"01/01/2017\",\"06/30/2018\",2.5%)": {"#NUM!", "#NUM!"},
		"=TBILLYIELD(\"06/30/2017\",\"01/01/2017\",2.5%)": {"#NUM!", "#NUM!"},
		// VDB
		"=VDB()":                          {"#VALUE!", "VDB requires 5 or 7 arguments"},
		"=VDB(-1,1000,5,0,1)":             {"#NUM!", "VDB requires cost >= 0"},
		"=VDB(10000,-1,5,0,1)":            {"#NUM!", "VDB requires salvage >= 0"},
		"=VDB(10000,1000,0,0,1)":          {"#NUM!", "VDB requires life > 0"},
		"=VDB(10000,1000,5,-1,1)":         {"#NUM!", "VDB requires start_period > 0"},
		"=VDB(10000,1000,5,2,1)":          {"#NUM!", "VDB requires start_period <= end_period"},
		"=VDB(10000,1000,5,0,6)":          {"#NUM!", "VDB requires end_period <= life"},
		"=VDB(10000,1000,5,0,1,-0.2)":     {"#VALUE!", "VDB requires factor >= 0"},
		"=VDB(\"\",1000,5,0,1)":           {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=VDB(10000,\"\",5,0,1)":          {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=VDB(10000,1000,\"\",0,1)":       {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=VDB(10000,1000,5,\"\",1)":       {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=VDB(10000,1000,5,0,\"\")":       {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=VDB(10000,1000,5,0,1,\"\")":     {"#NUM!", "#NUM!"},
		"=VDB(10000,1000,5,0,1,0.2,\"\")": {"#NUM!", "#NUM!"},
		// YIELD
		"=YIELD()": {"#VALUE!", "YIELD requires 6 or 7 arguments"},
		"=YIELD(\"\",\"06/30/2015\",10%,101,100,4)":                {"#VALUE!", "#VALUE!"},
		"=YIELD(\"01/01/2010\",\"\",10%,101,100,4)":                {"#VALUE!", "#VALUE!"},
		"=YIELD(\"01/01/2010\",\"06/30/2015\",\"\",101,100,4)":     {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=YIELD(\"01/01/2010\",\"06/30/2015\",10%,\"\",100,4)":     {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=YIELD(\"01/01/2010\",\"06/30/2015\",10%,101,\"\",4)":     {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=YIELD(\"01/01/2010\",\"06/30/2015\",10%,101,100,\"\")":   {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=YIELD(\"01/01/2010\",\"06/30/2015\",10%,101,100,4,\"\")": {"#NUM!", "#NUM!"},
		"=YIELD(\"01/01/2010\",\"06/30/2015\",10%,101,100,3)":      {"#NUM!", "#NUM!"},
		"=YIELD(\"01/01/2010\",\"06/30/2015\",10%,101,100,4,5)":    {"#NUM!", "invalid basis"},
		"=YIELD(\"01/01/2010\",\"06/30/2015\",-1,101,100,4)":       {"#NUM!", "YIELD requires rate >= 0"},
		"=YIELD(\"01/01/2010\",\"06/30/2015\",10%,0,100,4)":        {"#NUM!", "YIELD requires pr > 0"},
		"=YIELD(\"01/01/2010\",\"06/30/2015\",10%,101,-1,4)":       {"#NUM!", "YIELD requires redemption >= 0"},
		// YIELDDISC
		"=YIELDDISC()": {"#VALUE!", "YIELDDISC requires 4 or 5 arguments"},
		"=YIELDDISC(\"\",\"06/30/2017\",97,100,0)":              {"#VALUE!", "#VALUE!"},
		"=YIELDDISC(\"01/01/2017\",\"\",97,100,0)":              {"#VALUE!", "#VALUE!"},
		"=YIELDDISC(\"01/01/2017\",\"06/30/2017\",\"\",100,0)":  {"#VALUE!", "#VALUE!"},
		"=YIELDDISC(\"01/01/2017\",\"06/30/2017\",97,\"\",0)":   {"#VALUE!", "#VALUE!"},
		"=YIELDDISC(\"01/01/2017\",\"06/30/2017\",97,100,\"\")": {"#NUM!", "#NUM!"},
		"=YIELDDISC(\"01/01/2017\",\"06/30/2017\",0,100)":       {"#NUM!", "YIELDDISC requires pr > 0"},
		"=YIELDDISC(\"01/01/2017\",\"06/30/2017\",97,0)":        {"#NUM!", "YIELDDISC requires redemption > 0"},
		"=YIELDDISC(\"01/01/2017\",\"06/30/2017\",97,100,5)":    {"#NUM!", "invalid basis"},
		// YIELDMAT
		"=YIELDMAT()": {"#VALUE!", "YIELDMAT requires 5 or 6 arguments"},
		"=YIELDMAT(\"\",\"06/30/2018\",\"06/01/2014\",5.5%,101,0)":            {"#VALUE!", "#VALUE!"},
		"=YIELDMAT(\"01/01/2017\",\"\",\"06/01/2014\",5.5%,101,0)":            {"#VALUE!", "#VALUE!"},
		"=YIELDMAT(\"01/01/2017\",\"06/30/2018\",\"\",5.5%,101,0)":            {"#VALUE!", "#VALUE!"},
		"=YIELDMAT(\"01/01/2017\",\"06/30/2018\",\"06/01/2014\",\"\",101,0)":  {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=YIELDMAT(\"01/01/2017\",\"06/30/2018\",\"06/01/2014\",5,\"\",0)":    {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=YIELDMAT(\"01/01/2017\",\"06/30/2018\",\"06/01/2014\",5,5.5%,\"\")": {"#NUM!", "#NUM!"},
		"=YIELDMAT(\"06/01/2014\",\"06/30/2018\",\"01/01/2017\",5.5%,101,0)":  {"#NUM!", "YIELDMAT requires settlement > issue"},
		"=YIELDMAT(\"01/01/2017\",\"06/30/2018\",\"06/01/2014\",-1,101,0)":    {"#NUM!", "YIELDMAT requires rate >= 0"},
		"=YIELDMAT(\"01/01/2017\",\"06/30/2018\",\"06/01/2014\",1,0,0)":       {"#NUM!", "YIELDMAT requires pr > 0"},
		"=YIELDMAT(\"01/01/2017\",\"06/30/2018\",\"06/01/2014\",5.5%,101,5)":  {"#NUM!", "invalid basis"},
		// DISPIMG
		"=_xlfn.DISPIMG()": {"#VALUE!", "DISPIMG requires 2 numeric arguments"},
	}
	for formula, expected := range mathCalcError {
		f := prepareCalcData(cellData)
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
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
		"=SUM(Sheet1!A1:Sheet1!A2)":       "3",
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

	referenceCalcError := map[string][]string{
		// MDETERM
		"=MDETERM(A1:B3)": {"#VALUE!", "#VALUE!"},
		// SUM
		"=1+SUM(SUM(A1+A2/A4)*(2-3),2)": {"#VALUE!", "#DIV/0!"},
	}
	for formula, expected := range referenceCalcError {
		f := prepareCalcData(cellData)
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
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

	// Test get calculated cell value on not formula cell
	f := prepareCalcData(cellData)
	result, err := f.CalcCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "1", result)
	// Test get calculated cell value on not exists worksheet
	f = prepareCalcData(cellData)
	_, err = f.CalcCellValue("SheetN", "A1")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Test get calculated cell value with invalid sheet name
	_, err = f.CalcCellValue("Sheet:1", "A1")
	assert.Equal(t, ErrSheetNameInvalid, err)
	// Test get calculated cell value with not support formula
	f = prepareCalcData(cellData)
	assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "=UNSUPPORT(A1)"))
	_, err = f.CalcCellValue("Sheet1", "A1")
	assert.EqualError(t, err, "not support UNSUPPORT function")
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestCalcCellValue.xlsx")))
}

func TestCalcWithDefinedName(t *testing.T) {
	cellData := [][]interface{}{
		{"A1_as_string", "B1_as_string", 123, nil},
	}
	f := prepareCalcData(cellData)
	assert.NoError(t, f.SetDefinedName(&DefinedName{Name: "defined_name1", RefersTo: "Sheet1!A1", Scope: "Workbook"}))
	assert.NoError(t, f.SetDefinedName(&DefinedName{Name: "defined_name1", RefersTo: "Sheet1!B1", Scope: "Sheet1"}))
	assert.NoError(t, f.SetDefinedName(&DefinedName{Name: "defined_name2", RefersTo: "Sheet1!C1", Scope: "Workbook"}))

	assert.NoError(t, f.SetCellFormula("Sheet1", "D1", "=defined_name1"))
	result, err := f.CalcCellValue("Sheet1", "D1")
	assert.NoError(t, err)
	// DefinedName with scope WorkSheet takes precedence over DefinedName with scope Workbook, so we should get B1 value
	assert.Equal(t, "B1_as_string", result, "=defined_name1")

	assert.NoError(t, f.SetCellFormula("Sheet1", "D1", "=CONCATENATE(\"<\",defined_name1,\">\")"))
	result, err = f.CalcCellValue("Sheet1", "D1")
	assert.NoError(t, err)
	assert.Equal(t, "<B1_as_string>", result, "=defined_name1")

	// comparing numeric values
	assert.NoError(t, f.SetCellFormula("Sheet1", "D1", "=123=defined_name2"))
	result, err = f.CalcCellValue("Sheet1", "D1")
	assert.NoError(t, err)
	assert.Equal(t, "TRUE", result, "=123=defined_name2")

	// comparing text values
	assert.NoError(t, f.SetCellFormula("Sheet1", "D1", "=\"B1_as_string\"=defined_name1"))
	result, err = f.CalcCellValue("Sheet1", "D1")
	assert.NoError(t, err)
	assert.Equal(t, "TRUE", result, "=\"B1_as_string\"=defined_name1")

	// comparing text values
	assert.NoError(t, f.SetCellFormula("Sheet1", "D1", "=IF(\"B1_as_string\"=defined_name1,\"YES\",\"NO\")"))
	result, err = f.CalcCellValue("Sheet1", "D1")
	assert.NoError(t, err)
	assert.Equal(t, "YES", result, "=IF(\"B1_as_string\"=defined_name1,\"YES\",\"NO\")")
}

func TestCalcISBLANK(t *testing.T) {
	argsList := list.New()
	argsList.PushBack(formulaArg{
		Type: ArgUnknown,
	})
	fn := formulaFuncs{}
	result := fn.ISBLANK(argsList)
	assert.Equal(t, "TRUE", result.Value())
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
	assert.Equal(t, result.Value(), "FALSE")
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
	assert.Equal(t, compareFormulaArg(newEmptyFormulaArg(), newEmptyFormulaArg(), newNumberFormulaArg(matchModeMaxLess), false), criteriaEq)
	lhs := newListFormulaArg([]formulaArg{newEmptyFormulaArg()})
	rhs := newListFormulaArg([]formulaArg{newEmptyFormulaArg(), newEmptyFormulaArg()})
	assert.Equal(t, compareFormulaArg(lhs, rhs, newNumberFormulaArg(matchModeMaxLess), false), criteriaL)
	assert.Equal(t, compareFormulaArg(rhs, lhs, newNumberFormulaArg(matchModeMaxLess), false), criteriaG)

	lhs = newListFormulaArg([]formulaArg{newBoolFormulaArg(true)})
	rhs = newListFormulaArg([]formulaArg{newBoolFormulaArg(true)})
	assert.Equal(t, compareFormulaArg(lhs, rhs, newNumberFormulaArg(matchModeMaxLess), false), criteriaEq)

	lhs = newListFormulaArg([]formulaArg{newNumberFormulaArg(1)})
	rhs = newListFormulaArg([]formulaArg{newNumberFormulaArg(0)})
	assert.Equal(t, compareFormulaArg(lhs, rhs, newNumberFormulaArg(matchModeMaxLess), false), criteriaG)

	assert.Equal(t, compareFormulaArg(formulaArg{Type: ArgUnknown}, formulaArg{Type: ArgUnknown}, newNumberFormulaArg(matchModeMaxLess), false), criteriaErr)
}

func TestCalcCompareFormulaArgMatrix(t *testing.T) {
	lhs := newMatrixFormulaArg([][]formulaArg{{newEmptyFormulaArg()}})
	rhs := newMatrixFormulaArg([][]formulaArg{{newEmptyFormulaArg(), newEmptyFormulaArg()}})
	assert.Equal(t, compareFormulaArgMatrix(lhs, rhs, newNumberFormulaArg(matchModeMaxLess), false), criteriaL)

	lhs = newMatrixFormulaArg([][]formulaArg{{newEmptyFormulaArg(), newEmptyFormulaArg()}})
	rhs = newMatrixFormulaArg([][]formulaArg{{newEmptyFormulaArg()}})
	assert.Equal(t, compareFormulaArgMatrix(lhs, rhs, newNumberFormulaArg(matchModeMaxLess), false), criteriaG)

	lhs = newMatrixFormulaArg([][]formulaArg{{newNumberFormulaArg(1)}})
	rhs = newMatrixFormulaArg([][]formulaArg{{newNumberFormulaArg(0)}})
	assert.Equal(t, compareFormulaArgMatrix(lhs, rhs, newNumberFormulaArg(matchModeMaxLess), false), criteriaG)
}

func TestCalcANCHORARRAY(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 1))
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", 2))
	formulaType, ref := STCellFormulaTypeArray, "B1:B2"
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "A1:A2",
		FormulaOpts{Ref: &ref, Type: &formulaType}))
	assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "SUM(_xlfn.ANCHORARRAY($B$1))"))
	result, err := f.CalcCellValue("Sheet1", "C1")
	assert.NoError(t, err)
	assert.Equal(t, "3", result)

	assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "SUM(_xlfn.ANCHORARRAY(\"\",\"\"))"))
	result, err = f.CalcCellValue("Sheet1", "C1")
	assert.EqualError(t, err, "ANCHORARRAY requires 1 numeric argument")
	assert.Equal(t, "#VALUE!", result)

	fn := &formulaFuncs{f: f, sheet: "SheetN"}
	argsList := list.New()
	argsList.PushBack(newStringFormulaArg("$B$1"))
	formulaArg := fn.ANCHORARRAY(argsList)
	assert.Equal(t, "sheet SheetN does not exist", formulaArg.Value())

	fn.sheet = "Sheet1"
	argsList = argsList.Init()
	arg := newStringFormulaArg("$A$1")
	arg.cellRefs = list.New()
	arg.cellRefs.PushBack(cellRef{Row: 1, Col: 1})
	argsList.PushBack(arg)
	formulaArg = fn.ANCHORARRAY(argsList)
	assert.Equal(t, ArgEmpty, formulaArg.Type)

	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).SheetData.Row[0].C[0].F = &xlsxF{}
	formulaArg = fn.ANCHORARRAY(argsList)
	assert.Equal(t, ArgError, formulaArg.Type)
	assert.Equal(t, ErrParameterInvalid.Error(), formulaArg.Value())

	argsList = argsList.Init()
	arg = newStringFormulaArg("$B$1")
	arg.cellRefs = list.New()
	arg.cellRefs.PushBack(cellRef{Row: 1, Col: 1, Sheet: "SheetN"})
	argsList.PushBack(arg)
	ws.(*xlsxWorksheet).SheetData.Row[0].C[0].F = &xlsxF{Ref: "A1:A1"}
	formulaArg = fn.ANCHORARRAY(argsList)
	assert.Equal(t, ArgError, formulaArg.Type)
	assert.Equal(t, "sheet SheetN does not exist", formulaArg.Value())
}

func TestCalcArrayFormula(t *testing.T) {
	t.Run("matrix_multiplication", func(t *testing.T) {
		f := NewFile()
		assert.NoError(t, f.SetSheetRow("Sheet1", "A1", &[]int{1, 2}))
		assert.NoError(t, f.SetSheetRow("Sheet1", "A2", &[]int{3, 4}))
		formulaType, ref := STCellFormulaTypeArray, "C1:C2"
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "A1:A2*B1:B2",
			FormulaOpts{Ref: &ref, Type: &formulaType}))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err)
		assert.Equal(t, "2", result)
		result, err = f.CalcCellValue("Sheet1", "C2")
		assert.NoError(t, err)
		assert.Equal(t, "12", result)
	})
	t.Run("matrix_multiplication_with_defined_name", func(t *testing.T) {
		f := NewFile()
		assert.NoError(t, f.SetSheetRow("Sheet1", "A1", &[]int{1, 2}))
		assert.NoError(t, f.SetSheetRow("Sheet1", "A2", &[]int{3, 4}))
		assert.NoError(t, f.SetDefinedName(&DefinedName{
			Name:     "matrix",
			RefersTo: "Sheet1!$A$1:$A$2",
		}))
		formulaType, ref := STCellFormulaTypeArray, "C1:C2"
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "matrix*B1:B2+\"1\"",
			FormulaOpts{Ref: &ref, Type: &formulaType}))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err)
		assert.Equal(t, "3", result)
		result, err = f.CalcCellValue("Sheet1", "C2")
		assert.NoError(t, err)
		assert.Equal(t, "13", result)
	})
	t.Run("columm_multiplication", func(t *testing.T) {
		f := NewFile()
		assert.NoError(t, f.SetSheetRow("Sheet1", "A1", &[]int{1, 2}))
		assert.NoError(t, f.SetSheetRow("Sheet1", "A2", &[]int{3, 4}))
		formulaType, ref := STCellFormulaTypeArray, "C1:C1048576"
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "A:A*B:B",
			FormulaOpts{Ref: &ref, Type: &formulaType}))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err)
		assert.Equal(t, "2", result)
		result, err = f.CalcCellValue("Sheet1", "C2")
		assert.NoError(t, err)
		assert.Equal(t, "12", result)
	})
	t.Run("row_multiplication", func(t *testing.T) {
		f := NewFile()
		assert.NoError(t, f.SetSheetRow("Sheet1", "A1", &[]int{1, 2}))
		assert.NoError(t, f.SetSheetRow("Sheet1", "A2", &[]int{3, 4}))
		formulaType, ref := STCellFormulaTypeArray, "A3:XFD3"
		assert.NoError(t, f.SetCellFormula("Sheet1", "A3", "1:1*2:2",
			FormulaOpts{Ref: &ref, Type: &formulaType}))
		result, err := f.CalcCellValue("Sheet1", "A3")
		assert.NoError(t, err)
		assert.Equal(t, "3", result)
		result, err = f.CalcCellValue("Sheet1", "B3")
		assert.NoError(t, err)
		assert.Equal(t, "8", result)
	})
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
	calcError := map[string][]string{
		"=VLOOKUP(INT(1),C3:C3,1,FALSE)": {"#N/A", "VLOOKUP no result found"},
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "F4", formula))
		result, err := f.CalcCellValue("Sheet1", "F4")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
	}
}

func TestCalcBoolean(t *testing.T) {
	cellData := [][]interface{}{{0.5, "TRUE", -0.5, "FALSE", true}}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=AVERAGEA(A1:C1)":  "0.333333333333333",
		"=MAX(0.5,B1)":      "0.5",
		"=MAX(A1:B1)":       "0.5",
		"=MAXA(A1:B1)":      "0.5",
		"=MAXA(A1:E1)":      "1",
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

func TestCalcMAXMIN(t *testing.T) {
	cellData := [][]interface{}{{"1"}, {"2"}, {true}}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=MAX(A1:A3)":  "0",
		"=MAXA(A1:A3)": "1",
		"=MIN(A1:A3)":  "0",
		"=MINA(A1:A3)": "1",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", formula))
		result, err := f.CalcCellValue("Sheet1", "B1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
}

func TestCalcAVERAGEIF(t *testing.T) {
	f := prepareCalcData([][]interface{}{
		{"Monday", 500},
		{"Tuesday", 50},
		{"Thursday", 100},
		{"Friday", 100},
		{"Thursday", 200},
		{5, 300},
		{2, 200},
		{3, 100},
		{4, 50},
		{5, 100},
		{1, 50},
		{true, 200},
		{true, 250},
		{false, 50},
	})
	for formula, expected := range map[string]string{
		"=AVERAGEIF(A1:A14,\"Thursday\",B1:B14)": "150",
		"=AVERAGEIF(A1:A14,5,B1:B14)":            "200",
		"=AVERAGEIF(A1:A14,\">2\",B1:B14)":       "137.5",
		"=AVERAGEIF(A1:A14,TRUE,B1:B14)":         "225",
		"=AVERAGEIF(A1:A14,\"<>TRUE\",B1:B14)":   "150",
	} {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
}

func TestCalcCOVAR(t *testing.T) {
	cellData := [][]interface{}{
		{"array1", "array2"},
		{2, 22.9},
		{7, 33.49},
		{8, 34.5},
		{3, 27.61},
		{4, 19.5},
		{1, 10.11},
		{6, 37.9},
		{5, 31.08},
	}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=COVAR(A1:A9,B1:B9)":        "16.633125",
		"=COVAR(A2:A9,B2:B9)":        "16.633125",
		"=COVARIANCE.P(A1:A9,B1:B9)": "16.633125",
		"=COVARIANCE.P(A2:A9,B2:B9)": "16.633125",
		"=COVARIANCE.S(A1:A9,B1:B9)": "19.0092857142857",
		"=COVARIANCE.S(A2:A9,B2:B9)": "19.0092857142857",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	calcError := map[string][]string{
		"=COVAR()":                   {"#VALUE!", "COVAR requires 2 arguments"},
		"=COVAR(A2:A9,B3:B3)":        {"#N/A", "#N/A"},
		"=COVARIANCE.P()":            {"#VALUE!", "COVARIANCE.P requires 2 arguments"},
		"=COVARIANCE.P(A2:A9,B3:B3)": {"#N/A", "#N/A"},
		"=COVARIANCE.S()":            {"#VALUE!", "COVARIANCE.S requires 2 arguments"},
		"=COVARIANCE.S(A2:A9,B3:B3)": {"#N/A", "#N/A"},
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
	}
}

func TestCalcDatabase(t *testing.T) {
	cellData := [][]interface{}{
		{"Tree", "Height", "Age", "Yield", "Profit", "Height"},
		{nil, ">1000%", nil, nil, nil, "<16"},
		{},
		{"Tree", "Height", "Age", "Yield", "Profit"},
		{"Apple", 18, 20, 14, 105},
		{"Pear", 12, 12, 10, 96},
		{"Cherry", 13, 14, 9, 105},
		{"Apple", 14, nil, 10, 75},
		{"Pear", 9, 8, 8, 77},
		{"Apple", 12, 11, 6, 45},
	}
	f := prepareCalcData(cellData)
	assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "=\"=Apple\""))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A3", "=\"=Pear\""))
	assert.NoError(t, f.SetCellFormula("Sheet1", "C8", "=NA()"))
	formulaList := map[string]string{
		"=DAVERAGE(A4:E10,\"Profit\",A1:F3)": "73.25",
		"=DCOUNT(A4:E10,\"Age\",A1:F2)":      "1",
		"=DCOUNT(A4:E10,,A1:F2)":             "2",
		"=DCOUNT(A4:E10,\"Profit\",A1:F2)":   "2",
		"=DCOUNT(A4:E10,\"Tree\",A1:F2)":     "0",
		"=DCOUNT(A4:E10,\"Age\",A2:F3)":      "0",
		"=DCOUNTA(A4:E10,\"Age\",A1:F2)":     "1",
		"=DCOUNTA(A4:E10,,A1:F2)":            "2",
		"=DCOUNTA(A4:E10,\"Profit\",A1:F2)":  "2",
		"=DCOUNTA(A4:E10,\"Tree\",A1:F2)":    "2",
		"=DCOUNTA(A4:E10,\"Age\",A2:F3)":     "0",
		"=DGET(A4:E6,\"Profit\",A1:F3)":      "96",
		"=DMAX(A4:E10,\"Tree\",A1:F3)":       "0",
		"=DMAX(A4:E10,\"Profit\",A1:F3)":     "96",
		"=DMIN(A4:E10,\"Tree\",A1:F3)":       "0",
		"=DMIN(A4:E10,\"Profit\",A1:F3)":     "45",
		"=DPRODUCT(A4:E10,\"Profit\",A1:F3)": "24948000",
		"=DSTDEV(A4:E10,\"Profit\",A1:F3)":   "21.077238908358",
		"=DSTDEVP(A4:E10,\"Profit\",A1:F3)":  "18.2534243362718",
		"=DSUM(A4:E10,\"Profit\",A1:F3)":     "293",
		"=DVAR(A4:E10,\"Profit\",A1:F3)":     "444.25",
		"=DVARP(A4:E10,\"Profit\",A1:F3)":    "333.1875",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "A11", formula))
		result, err := f.CalcCellValue("Sheet1", "A11")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	calcError := map[string][]string{
		"=DAVERAGE()":                         {"#VALUE!", "DAVERAGE requires 3 arguments"},
		"=DAVERAGE(A4:E10,\"x\",A1:F3)":       {"#VALUE!", "#VALUE!"},
		"=DAVERAGE(A4:E10,\"Tree\",A1:F3)":    {"#DIV/0!", "#DIV/0!"},
		"=DCOUNT()":                           {"#VALUE!", "DCOUNT requires at least 2 arguments"},
		"=DCOUNT(A4:E10,\"Age\",A1:F2,\"\")":  {"#VALUE!", "DCOUNT allows at most 3 arguments"},
		"=DCOUNT(A4,\"Age\",A1:F2)":           {"#VALUE!", "#VALUE!"},
		"=DCOUNT(A4:E10,NA(),A1:F2)":          {"#VALUE!", "#VALUE!"},
		"=DCOUNT(A4:E4,,A1:F2)":               {"#VALUE!", "#VALUE!"},
		"=DCOUNT(A4:E10,\"x\",A2:F3)":         {"#VALUE!", "#VALUE!"},
		"=DCOUNTA()":                          {"#VALUE!", "DCOUNTA requires at least 2 arguments"},
		"=DCOUNTA(A4:E10,\"Age\",A1:F2,\"\")": {"#VALUE!", "DCOUNTA allows at most 3 arguments"},
		"=DCOUNTA(A4,\"Age\",A1:F2)":          {"#VALUE!", "#VALUE!"},
		"=DCOUNTA(A4:E10,NA(),A1:F2)":         {"#VALUE!", "#VALUE!"},
		"=DCOUNTA(A4:E4,,A1:F2)":              {"#VALUE!", "#VALUE!"},
		"=DCOUNTA(A4:E10,\"x\",A2:F3)":        {"#VALUE!", "#VALUE!"},
		"=DGET()":                             {"#VALUE!", "DGET requires 3 arguments"},
		"=DGET(A4:E5,\"Profit\",A1:F3)":       {"#VALUE!", "#VALUE!"},
		"=DGET(A4:E10,\"Profit\",A1:F3)":      {"#NUM!", "#NUM!"},
		"=DMAX()":                             {"#VALUE!", "DMAX requires 3 arguments"},
		"=DMAX(A4:E10,\"x\",A1:F3)":           {"#VALUE!", "#VALUE!"},
		"=DMIN()":                             {"#VALUE!", "DMIN requires 3 arguments"},
		"=DMIN(A4:E10,\"x\",A1:F3)":           {"#VALUE!", "#VALUE!"},
		"=DPRODUCT()":                         {"#VALUE!", "DPRODUCT requires 3 arguments"},
		"=DPRODUCT(A4:E10,\"x\",A1:F3)":       {"#VALUE!", "#VALUE!"},
		"=DSTDEV()":                           {"#VALUE!", "DSTDEV requires 3 arguments"},
		"=DSTDEV(A4:E10,\"x\",A1:F3)":         {"#VALUE!", "#VALUE!"},
		"=DSTDEVP()":                          {"#VALUE!", "DSTDEVP requires 3 arguments"},
		"=DSTDEVP(A4:E10,\"x\",A1:F3)":        {"#VALUE!", "#VALUE!"},
		"=DSUM()":                             {"#VALUE!", "DSUM requires 3 arguments"},
		"=DSUM(A4:E10,\"x\",A1:F3)":           {"#VALUE!", "#VALUE!"},
		"=DVAR()":                             {"#VALUE!", "DVAR requires 3 arguments"},
		"=DVAR(A4:E10,\"x\",A1:F3)":           {"#VALUE!", "#VALUE!"},
		"=DVARP()":                            {"#VALUE!", "DVARP requires 3 arguments"},
		"=DVARP(A4:E10,\"x\",A1:F3)":          {"#VALUE!", "#VALUE!"},
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "A11", formula))
		result, err := f.CalcCellValue("Sheet1", "A11")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
	}
}

func TestCalcDBCS(t *testing.T) {
	f := NewFile(Options{CultureInfo: CultureNameZhCN})
	assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "=DBCS(\"`~·!@#$¥%…^&*()_-+=[]{}\\|;:'\"\"<,>.?/01234567890 abc ABC \uff65\uff9e\uff9f \uff74\uff78\uff7e\uff99\")"))
	result, err := f.CalcCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "\uff40\uff5e\u00b7\uff01\uff20\uff03\uff04\u00a5\uff05\u2026\uff3e\uff06\uff0a\uff08\uff09\uff3f\uff0d\uff0b\uff1d\uff3b\uff3d\uff5b\uff5d\uff3c\uff5c\uff1b\uff1a\uff07\uff02\uff1c\uff0c\uff1e\uff0e\uff1f\uff0f\uff10\uff11\uff12\uff13\uff14\uff15\uff16\uff17\uff18\uff19\uff10\u3000\uff41\uff42\uff43\u3000\uff21\uff22\uff23\u3000\uff65\uff9e\uff9f\u3000\uff74\uff78\uff7e\uff99", result)
}

func TestCalcFORMULATEXT(t *testing.T) {
	f, formulaText := NewFile(), "=SUM(B1:C1)"
	assert.NoError(t, f.SetCellFormula("Sheet1", "A1", formulaText))
	for _, formula := range []string{"=FORMULATEXT(A1)", "=FORMULATEXT(A:A)", "=FORMULATEXT(A1:B1)"} {
		assert.NoError(t, f.SetCellFormula("Sheet1", "D1", formula), formula)
		result, err := f.CalcCellValue("Sheet1", "D1")
		assert.NoError(t, err, formula)
		assert.Equal(t, formulaText, result, formula)
	}
}

func TestCalcGROWTHandTREND(t *testing.T) {
	cellData := [][]interface{}{
		{"known_x's", "known_y's", 0, -1},
		{1, 10, 1},
		{2, 20, 1},
		{3, 40},
		{4, 80},
		{},
		{"new_x's", "new_y's"},
		{5},
		{6},
		{7},
	}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=GROWTH(A2:B2)":                    "1",
		"=GROWTH(B2:B5,A2:A5,A8:A10)":       "160",
		"=GROWTH(B2:B5,A2:A5,A8:A10,FALSE)": "467.842838114059",
		"=GROWTH(A4:A5,A2:B3,A8:A10,FALSE)": "",
		"=GROWTH(A3:A5,A2:B4,A2:B3)":        "2",
		"=GROWTH(A4:A5,A2:B3)":              "",
		"=GROWTH(A2:B2,A2:B3)":              "",
		"=GROWTH(A2:B2,A2:B3,A2:B3,FALSE)":  "1.28402541668774",
		"=GROWTH(A2:B2,A4:B5,A4:B5,FALSE)":  "1",
		"=GROWTH(A3:C3,A2:C3,A2:B3)":        "2",
		"=TREND(A2:B2)":                     "1",
		"=TREND(B2:B5,A2:A5,A8:A10)":        "95",
		"=TREND(B2:B5,A2:A5,A8:A10,FALSE)":  "81.6666666666667",
		"=TREND(A4:A5,A2:B3,A8:A10,FALSE)":  "",
		"=TREND(A4:A5,A2:B3,A2:B3,FALSE)":   "1.5",
		"=TREND(A3:A5,A2:B4,A2:B3)":         "2",
		"=TREND(A4:A5,A2:B3)":               "",
		"=TREND(A2:B2,A2:B3)":               "",
		"=TREND(A2:B2,A2:B3,A2:B3,FALSE)":   "1",
		"=TREND(A2:B2,A4:B5,A4:B5,FALSE)":   "1",
		"=TREND(A3:C3,A2:C3,A2:B3)":         "2",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	calcError := map[string][]string{
		"=GROWTH()":                          {"#VALUE!", "GROWTH requires at least 1 argument"},
		"=GROWTH(B2:B5,A2:A5,A8:A10,TRUE,0)": {"#VALUE!", "GROWTH allows at most 4 arguments"},
		"=GROWTH(A1:B1,A2:A5,A8:A10,TRUE)":   {"#VALUE!", "#VALUE!"},
		"=GROWTH(B2:B5,A1:B1,A8:A10,TRUE)":   {"#VALUE!", "#VALUE!"},
		"=GROWTH(B2:B5,A2:A5,A1:B1,TRUE)":    {"#VALUE!", "#VALUE!"},
		"=GROWTH(B2:B5,A2:A5,A8:A10,\"\")":   {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=GROWTH(A2:B3,A4:B4)":               {"#REF!", "#REF!"},
		"=GROWTH(A4:B4,A2:A2)":               {"#REF!", "#REF!"},
		"=GROWTH(A2:A2,A4:A5)":               {"#REF!", "#REF!"},
		"=GROWTH(C1:C1,A2:A3)":               {"#VALUE!", "#VALUE!"},
		"=GROWTH(D1:D1,A2:A3)":               {"#NUM!", "#NUM!"},
		"=GROWTH(A2:A3,C1:C1)":               {"#VALUE!", "#VALUE!"},
		"=TREND()":                           {"#VALUE!", "TREND requires at least 1 argument"},
		"=TREND(B2:B5,A2:A5,A8:A10,TRUE,0)":  {"#VALUE!", "TREND allows at most 4 arguments"},
		"=TREND(A1:B1,A2:A5,A8:A10,TRUE)":    {"#VALUE!", "#VALUE!"},
		"=TREND(B2:B5,A1:B1,A8:A10,TRUE)":    {"#VALUE!", "#VALUE!"},
		"=TREND(B2:B5,A2:A5,A1:B1,TRUE)":     {"#VALUE!", "#VALUE!"},
		"=TREND(B2:B5,A2:A5,A8:A10,\"\")":    {"#VALUE!", "strconv.ParseBool: parsing \"\": invalid syntax"},
		"=TREND(A2:B3,A4:B4)":                {"#REF!", "#REF!"},
		"=TREND(A4:B4,A2:A2)":                {"#REF!", "#REF!"},
		"=TREND(A2:A2,A4:A5)":                {"#REF!", "#REF!"},
		"=TREND(C1:C1,A2:A3)":                {"#VALUE!", "#VALUE!"},
		"=TREND(D1:D1,A2:A3)":                {"#REF!", "#REF!"},
		"=TREND(A2:A3,C1:C1)":                {"#VALUE!", "#VALUE!"},
		"=TREND(C1:C1,C1:C1)":                {"#VALUE!", "#VALUE!"},
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
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
	calcError := map[string][]string{
		"=HLOOKUP(INT(1),A3:A3,1,FALSE)": {"#N/A", "HLOOKUP no result found"},
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "B10", formula))
		result, err := f.CalcCellValue("Sheet1", "B10")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
	}
}

func TestCalcCHITESTandCHISQdotTEST(t *testing.T) {
	cellData := [][]interface{}{
		{nil, "Observed Frequencies", nil, nil, "Expected Frequencies"},
		{nil, "men", "women", nil, nil, "men", "women"},
		{"answer a", 33, 39, nil, "answer a", 26.25, 31.5},
		{"answer b", 62, 62, nil, "answer b", 57.75, 61.95},
		{"answer c", 10, 4, nil, "answer c", 21, 11.55},
		{nil, -1, 0},
	}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=CHITEST(B3:C5,F3:G5)":    "0.000699102758787672",
		"=CHITEST(B3:C3,F3:G3)":    "0.0605802098655177",
		"=CHITEST(B3:B4,F3:F4)":    "0.152357748933542",
		"=CHITEST(B4:B6,F3:F5)":    "7.07076951440726E-25",
		"=CHISQ.TEST(B3:C5,F3:G5)": "0.000699102758787672",
		"=CHISQ.TEST(B3:C3,F3:G3)": "0.0605802098655177",
		"=CHISQ.TEST(B3:B4,F3:F4)": "0.152357748933542",
		"=CHISQ.TEST(B4:B6,F3:F5)": "7.07076951440726E-25",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "I1", formula))
		result, err := f.CalcCellValue("Sheet1", "I1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	calcError := map[string][]string{
		"=CHITEST()":                  {"#VALUE!", "CHITEST requires 2 arguments"},
		"=CHITEST(MUNIT(0),MUNIT(0))": {"#VALUE!", "#VALUE!"},
		"=CHITEST(B3:C5,F3:F4)":       {"#N/A", "#N/A"},
		"=CHITEST(B3:B3,F3:F3)":       {"#N/A", "#N/A"},
		"=CHITEST(F3:F5,B4:B6)":       {"#NUM!", "#NUM!"},
		"=CHITEST(F3:F5,C4:C6)":       {"#DIV/0!", "#DIV/0!"},
		"=CHISQ.TEST()":               {"#VALUE!", "CHISQ.TEST requires 2 arguments"},
		"=CHISQ.TEST(B3:C5,F3:F4)":    {"#N/A", "#N/A"},
		"=CHISQ.TEST(B3:B3,F3:F3)":    {"#N/A", "#N/A"},
		"=CHISQ.TEST(F3:F5,B4:B6)":    {"#NUM!", "#NUM!"},
		"=CHISQ.TEST(F3:F5,C4:C6)":    {"#DIV/0!", "#DIV/0!"},
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "I1", formula))
		result, err := f.CalcCellValue("Sheet1", "I1")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
	}
}

func TestCalcFTEST(t *testing.T) {
	cellData := [][]interface{}{
		{"Group 1", "Group 2"},
		{3.5, 9.2},
		{4.7, 8.2},
		{6.2, 7.3},
		{4.9, 6.1},
		{3.8, 5.4},
		{5.5, 7.8},
		{7.1, 5.9},
		{6.7, 8.4},
		{3.9, 7.7},
		{4.6, 6.6},
	}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=FTEST(A2:A11,B2:B11)":  "0.95403555939413",
		"=F.TEST(A2:A11,B2:B11)": "0.95403555939413",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	calcError := map[string][]string{
		"=FTEST()":               {"#VALUE!", "FTEST requires 2 arguments"},
		"=FTEST(A2:A2,B2:B2)":    {"#DIV/0!", "#DIV/0!"},
		"=FTEST(A12:A14,B2:B4)":  {"#DIV/0!", "#DIV/0!"},
		"=FTEST(A2:A4,B2:B2)":    {"#DIV/0!", "#DIV/0!"},
		"=FTEST(A2:A4,B12:B14)":  {"#DIV/0!", "#DIV/0!"},
		"=F.TEST()":              {"#VALUE!", "F.TEST requires 2 arguments"},
		"=F.TEST(A2:A2,B2:B2)":   {"#DIV/0!", "#DIV/0!"},
		"=F.TEST(A12:A14,B2:B4)": {"#DIV/0!", "#DIV/0!"},
		"=F.TEST(A2:A4,B2:B2)":   {"#DIV/0!", "#DIV/0!"},
		"=F.TEST(A2:A4,B12:B14)": {"#DIV/0!", "#DIV/0!"},
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
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
	calcError := map[string][]string{
		"=IRR()":       {"#VALUE!", "IRR requires at least 1 argument"},
		"=IRR(0,0,0)":  {"#VALUE!", "IRR allows at most 2 arguments"},
		"=IRR(0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=IRR(A2:A3)":  {"#NUM!", "#NUM!"},
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", formula))
		result, err := f.CalcCellValue("Sheet1", "B1")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
	}
}

func TestCalcMAXMINIFS(t *testing.T) {
	f := NewFile()
	for cell, row := range map[string][]interface{}{
		"A1": {1, -math.MaxFloat64 - 1},
		"A2": {2, -math.MaxFloat64 - 2},
		"A3": {3, math.MaxFloat64 + 1},
		"A4": {4, math.MaxFloat64 + 2},
	} {
		assert.NoError(t, f.SetSheetRow("Sheet1", cell, &row))
	}
	formulaList := map[string]string{
		"=MAX(B1:B2)":                 "0",
		"=MAXIFS(B1:B2,A1:A2,\">0\")": "0",
		"=MIN(B3:B4)":                 "0",
		"=MINIFS(B3:B4,A3:A4,\"<0\")": "0",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
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
	calcError := map[string][]string{
		"=MIRR()":             {"#VALUE!", "MIRR requires 3 arguments"},
		"=MIRR(A1:A5,\"\",0)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=MIRR(A1:A5,0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=MIRR(B1:B5,0,0)":    {"#DIV/0!", "#DIV/0!"},
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", formula))
		result, err := f.CalcCellValue("Sheet1", "B1")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
	}
}

func TestCalcSUMIFSAndAVERAGEIFS(t *testing.T) {
	cellData := [][]interface{}{
		{"Quarter", "Area", "Sales Rep.", "Sales"},
		{1, "North", "Jeff", 223000},
		{1, "North", "Chris", 125000},
		{1, "South", "Carol", 456000},
		{2, "North", "Jeff", 322000},
		{2, "North", "Chris", 340000},
		{2, "South", "Carol", 198000},
		{3, "North", "Jeff", 310000},
		{3, "North", "Chris", 250000},
		{3, "South", "Carol", 460000},
		{4, "North", "Jeff", 261000},
		{4, "North", "Chris", 389000},
		{4, "South", "Carol", 305000},
	}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=AVERAGEIFS(D2:D13,A2:A13,1,B2:B13,\"North\")":                "174000",
		"=AVERAGEIFS(D2:D13,A2:A13,\">2\",C2:C13,\"Jeff\")":            "285500",
		"=SUMIFS(D2:D13,A2:A13,1,B2:B13,\"North\")":                    "348000",
		"=SUMIFS(D2:D13,A2:A13,\">2\",C2:C13,\"Jeff\")":                "571000",
		"=SUMIFS(D2:D13,A2:A13,1,D2:D13,125000)":                       "125000",
		"=SUMIFS(D2:D13,A2:A13,1,D2:D13,\">100000\",C2:C13,\"Chris\")": "125000",
		"=SUMIFS(D2:D13,A2:A13,1,D2:D13,\"<40000\",C2:C13,\"Chris\")":  "0",
		"=SUMIFS(D2:D13,A2:A13,1,A2:A13,2)":                            "0",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "E1", formula))
		result, err := f.CalcCellValue("Sheet1", "E1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	calcError := map[string][]string{
		"=AVERAGEIFS()":                                  {"#VALUE!", "AVERAGEIFS requires at least 3 arguments"},
		"=AVERAGEIFS(H1,\"\")":                           {"#VALUE!", "AVERAGEIFS requires at least 3 arguments"},
		"=AVERAGEIFS(H1,\"\",TRUE,1)":                    {"#N/A", "#N/A"},
		"=AVERAGEIFS(H1,\"\",TRUE)":                      {"#DIV/0!", "AVERAGEIF divide by zero"},
		"=AVERAGEIFS(D2:D13,A2:A13,1,A2:A13,2)":          {"#DIV/0!", "AVERAGEIF divide by zero"},
		"=SUMIFS()":                                      {"#VALUE!", "SUMIFS requires at least 3 arguments"},
		"=SUMIFS(D2:D13,A2:A13,1,B2:B13)":                {"#N/A", "#N/A"},
		"=SUMIFS(D20:D23,A2:A13,\">2\",C2:C13,\"Jeff\")": {"#VALUE!", "#VALUE!"},
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "E1", formula))
		result, err := f.CalcCellValue("Sheet1", "E1")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
	}
}

func TestCalcXIRR(t *testing.T) {
	cellData := [][]interface{}{
		{-100.00, 42370},
		{20.00, 42461},
		{40.00, 42644},
		{25.00, 42767},
		{8.00, 42795},
		{15.00, 42887},
		{-1e-10, 42979},
	}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=XIRR(A1:A4,B1:B4)":     "-0.196743861298328",
		"=XIRR(A1:A6,B1:B6,0.5)": "0.0944390744445204",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	calcError := map[string][]string{
		"=XIRR()":                 {"#VALUE!", "XIRR requires 2 or 3 arguments"},
		"=XIRR(A1:A4,B1:B4,-1)":   {"#VALUE!", "XIRR requires guess > -1"},
		"=XIRR(\"\",B1:B4)":       {"#VALUE!", "#VALUE!"},
		"=XIRR(A1:A4,\"\")":       {"#VALUE!", "#VALUE!"},
		"=XIRR(A1:A4,B1:B4,\"\")": {"#NUM!", "#NUM!"},
		"=XIRR(A2:A6,B2:B6)":      {"#NUM!", "#NUM!"},
		"=XIRR(A2:A7,B2:B7)":      {"#NUM!", "#NUM!"},
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
	}
}

func TestCalcXLOOKUP(t *testing.T) {
	cellData := [][]interface{}{
		{},
		{nil, nil, "Quarter", "Gross Profit", "Net profit", "Profit %"},
		{nil, nil, "Qtr1", nil, 19342, 29.30},
		{},
		{nil, "Income Statement", "Qtr1", "Qtr2", "Qtr3", "Qtr4", "Total"},
		{nil, "Total sales", 50000, 78200, 89500, 91250, 308.95},
		{nil, "Cost of sales", -25000, -42050, -59450, -60450, -186950},
		{nil, "Gross Profit", 25000, 36150, 30050, 30800, 122000},
		{},
		{nil, "Depreciation", -899, -791, -202, -412, -2304},
		{nil, "Interest", -513, -853, -150, -956, -2472},
		{nil, "Earnings before Tax", 23588, 34506, 29698, 29432, 117224},
		{},
		{nil, "Tax", -4246, -6211, -5346, -5298, 21100},
		{},
		{nil, "Net profit", 19342, 28295, 24352, 24134, 96124},
		{nil, "Profit %", 0.293, 0.278, 0.234, 0.276, 0.269},
	}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=SUM(XLOOKUP($C3,$C5:$C5,$C6:$C17,NA(),0,2))":        "87272.293",
		"=SUM(XLOOKUP($C3,$C5:$C5,$C6:$G6,NA(),0,-2))":        "309258.95",
		"=SUM(XLOOKUP($C3,$C5:$C5,$C6:$C17,NA(),0,-2))":       "87272.293",
		"=SUM(XLOOKUP($C3,$C5:$G5,$C6:$G17,NA(),0,2))":        "87272.293",
		"=SUM(XLOOKUP(D2,$B6:$B17,$C6:$G17,NA(),0,2))":        "244000",
		"=XLOOKUP(D2,$B6:$B17,C6:C17)":                        "25000",
		"=XLOOKUP(D2,$B6:$B17,XLOOKUP($C3,$C5:$G5,$C6:$G17))": "25000",
		"=XLOOKUP(\"*p*\",B2:B9,C2:C9,NA(),2)":                "25000",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "D3", formula))
		result, err := f.CalcCellValue("Sheet1", "D3")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	calcError := map[string][]string{
		"=XLOOKUP()": {"#VALUE!", "XLOOKUP requires at least 3 arguments"},
		"=XLOOKUP($C3,$C5:$C5,$C6:$C17,NA(),0,2,1)":  {"#VALUE!", "XLOOKUP allows at most 6 arguments"},
		"=XLOOKUP($C3,$C5,$C6,NA(),0,2)":             {"#N/A", "#N/A"},
		"=XLOOKUP($C3,$C4:$D5,$C6:$C17,NA(),0,2)":    {"#VALUE!", "#VALUE!"},
		"=XLOOKUP($C3,$C5:$C5,$C6:$G17,NA(),0,-2)":   {"#VALUE!", "#VALUE!"},
		"=XLOOKUP($C3,$C5:$G5,$C6:$F7,NA(),0,2)":     {"#VALUE!", "#VALUE!"},
		"=XLOOKUP(D2,$B6:$B17,$C6:$G16,NA(),0,2)":    {"#VALUE!", "#VALUE!"},
		"=XLOOKUP(D2,$B6:$B17,$C6:$G17,NA(),3,2)":    {"#VALUE!", "#VALUE!"},
		"=XLOOKUP(D2,$B6:$B17,$C6:$G17,NA(),0,0)":    {"#VALUE!", "#VALUE!"},
		"=XLOOKUP(D2,$B6:$B17,$C6:$G17,NA(),\"\",2)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=XLOOKUP(D2,$B6:$B17,$C6:$G17,NA(),0,\"\")": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "D3", formula))
		result, err := f.CalcCellValue("Sheet1", "D3")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
	}

	cellData = [][]interface{}{
		{"Salesperson", "Item", "Amont"},
		{"B", "Apples", 30, 25, 15, 50, 45, 18},
		{"L", "Oranges", 25, "D3", "E3"},
		{"C", "Grapes", 15},
		{"L", "Lemons", 50},
		{"L", "Oranges", 45},
		{"C", "Peaches", 18},
		{"B", "Pears", 40},
		{"B", "Apples", 55},
	}
	f = prepareCalcData(cellData)
	formulaList = map[string]string{
		// Test match mode with partial match (wildcards)
		"=XLOOKUP(\"*p*\",B2:B9,C2:C9,NA(),2)": "30",
		// Test match mode with approximate match in vertical (next larger item)
		"=XLOOKUP(32,B2:B9,C2:C9,NA(),1)": "30",
		// Test match mode with approximate match in horizontal (next larger item)
		"=XLOOKUP(30,C2:F2,C3:F3,NA(),1)": "25",
		// Test match mode with approximate match in vertical (next smaller item)
		"=XLOOKUP(40,C2:C9,B2:B9,NA(),-1)": "Pears",
		// Test match mode with approximate match in horizontal (next smaller item)
		"=XLOOKUP(29,C2:F2,C3:F3,NA(),-1)": "D3",
		// Test search mode
		"=XLOOKUP(\"L\",A2:A9,C2:C9,NA(),0,1)":  "25",
		"=XLOOKUP(\"L\",A2:A9,C2:C9,NA(),0,-1)": "45",
		"=XLOOKUP(\"L\",A2:A9,C2:C9,NA(),0,2)":  "50",
		"=XLOOKUP(\"L\",A2:A9,C2:C9,NA(),0,-2)": "45",
		// Test match mode and search mode
		"=XLOOKUP(29,C2:H2,C3:H3,NA(),-1,-1)": "D3",
		"=XLOOKUP(29,C2:H2,C3:H3,NA(),-1,1)":  "D3",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "D4", formula))
		result, err := f.CalcCellValue("Sheet1", "D4")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	calcError = map[string][]string{
		// Test match mode with exact match
		"=XLOOKUP(\"*p*\",B2:B9,C2:C9,NA(),0)": {"#N/A", "#N/A"},
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "D3", formula))
		result, err := f.CalcCellValue("Sheet1", "D3")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
	}
}

func TestCalcXNPV(t *testing.T) {
	cellData := [][]interface{}{
		{nil, 0.05},
		{42370, -10000, nil},
		{42401, 2000},
		{42491, 2400},
		{42552, 2900},
		{42675, 3500},
		{42736, 4100},
		{},
		{42401},
		{42370},
	}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=XNPV(B1,B2:B7,A2:A7)": "4447.93800944052",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	calcError := map[string][]string{
		"=XNPV()":                 {"#VALUE!", "XNPV requires 3 arguments"},
		"=XNPV(\"\",B2:B7,A2:A7)": {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=XNPV(0,B2:B7,A2:A7)":    {"#VALUE!", "XNPV requires rate > 0"},
		"=XNPV(B1,\"\",A2:A7)":    {"#VALUE!", "#VALUE!"},
		"=XNPV(B1,B2:B7,\"\")":    {"#VALUE!", "#VALUE!"},
		"=XNPV(B1,B2:B7,C2:C7)":   {"#VALUE!", "#VALUE!"},
		"=XNPV(B1,B2,A2)":         {"#NUM!", "#NUM!"},
		"=XNPV(B1,B2:B3,A2:A5)":   {"#NUM!", "#NUM!"},
		"=XNPV(B1,B2:B3,A9:A10)":  {"#VALUE!", "#VALUE!"},
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
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
		"=MATCH(-10,D1:D6,-1)":     "6",
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
		assert.Equal(t, expected, result, formula)
	}
	assert.Equal(t, newErrorFormulaArg(formulaErrorNA, formulaErrorNA), calcMatch(2, nil, []formulaArg{}))
}

func TestCalcISFORMULA(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "=ISFORMULA(A1)"))
	for _, formula := range []string{"=NA()", "=SUM(A1:A3)"} {
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", formula))
		result, err := f.CalcCellValue("Sheet1", "B1")
		assert.NoError(t, err, formula)
		assert.Equal(t, "TRUE", result, formula)
	}
}

func TestCalcMODE(t *testing.T) {
	cellData := [][]interface{}{
		{1, 1},
		{1, 1},
		{2, 2},
		{2, 2},
		{3, 2},
		{3},
		{3},
		{4},
		{4},
		{4},
	}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=MODE(A1:A10)":      "3",
		"=MODE(B1:B6)":       "2",
		"=MODE.MULT(A1:A10)": "3",
		"=MODE.SNGL(A1:A10)": "3",
		"=MODE.SNGL(B1:B6)":  "2",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	calcError := map[string][]string{
		"=MODE()":            {"#VALUE!", "MODE requires at least 1 argument"},
		"=MODE(0,\"\")":      {"#VALUE!", "#VALUE!"},
		"=MODE(D1:D3)":       {"#N/A", "#N/A"},
		"=MODE.MULT()":       {"#VALUE!", "MODE.MULT requires at least 1 argument"},
		"=MODE.MULT(0,\"\")": {"#VALUE!", "#VALUE!"},
		"=MODE.MULT(D1:D3)":  {"#N/A", "#N/A"},
		"=MODE.SNGL()":       {"#VALUE!", "MODE.SNGL requires at least 1 argument"},
		"=MODE.SNGL(0,\"\")": {"#VALUE!", "#VALUE!"},
		"=MODE.SNGL(D1:D3)":  {"#N/A", "#N/A"},
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
	}
}

func TestCalcPEARSON(t *testing.T) {
	cellData := [][]interface{}{
		{"x", "y"},
		{1, 10.11},
		{2, 22.9},
		{2, 27.61},
		{3, 27.61},
		{4, 11.15},
		{5, 31.08},
		{6, 37.9},
		{7, 33.49},
		{8, 21.05},
		{9, 27.01},
		{10, 45.78},
		{11, 31.32},
		{12, 50.57},
		{13, 45.48},
		{14, 40.94},
		{15, 53.76},
		{16, 36.18},
		{17, 49.77},
		{18, 55.66},
		{19, 63.83},
		{20, 63.6},
	}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=PEARSON(A2:A22,B2:B22)": "0.864129542184994",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
}

func TestCalcPROB(t *testing.T) {
	cellData := [][]interface{}{
		{"x", "probability"},
		{0, 0.1},
		{1, 0.15},
		{2, 0.17},
		{3, 0.22},
		{4, 0.21},
		{5, 0.09},
		{6, 0.05},
		{7, 0.01},
	}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=PROB(A2:A9,B2:B9,3)":    "0.22",
		"=PROB(A2:A9,B2:B9,3,5)":  "0.52",
		"=PROB(A2:A9,B2:B9,8,10)": "0",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "=NA()"))
	calcError := map[string][]string{
		"=PROB(A2:A9,B2:B9,3)": {"#NUM!", "#NUM!"},
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
	}
}

func TestCalcRSQ(t *testing.T) {
	cellData := [][]interface{}{
		{"known_y's", "known_x's"},
		{2, 22.9},
		{7, 33.49},
		{8, 34.5},
		{3, 27.61},
		{4, 19.5},
		{1, 10.11},
		{6, 37.9},
		{5, 31.08},
	}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=RSQ(A2:A9,B2:B9)": "0.711666290486784",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
}

func TestCalcSLOP(t *testing.T) {
	cellData := [][]interface{}{
		{"known_x's", "known_y's"},
		{1, 3},
		{2, 7},
		{3, 17},
		{4, 20},
		{5, 20},
		{6, 27},
	}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=SLOPE(A2:A7,B2:B7)": "0.200826446280992",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
}

func TestCalcSHEET(t *testing.T) {
	f := NewFile()
	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)
	formulaList := map[string]string{
		"=SHEET(\"Sheet2\")":   "2",
		"=SHEET(Sheet2!A1)":    "2",
		"=SHEET(Sheet2!A1:A2)": "2",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", formula))
		result, err := f.CalcCellValue("Sheet1", "A1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
}

func TestCalcSHEETS(t *testing.T) {
	f := NewFile()
	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)
	formulaList := map[string]string{
		"=SHEETS(Sheet1!A1:B1)":        "1",
		"=SHEETS(Sheet1!A1:Sheet1!B1)": "1",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", formula))
		result, err := f.CalcCellValue("Sheet1", "A1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
}

func TestCalcSTEY(t *testing.T) {
	cellData := [][]interface{}{
		{"known_x's", "known_y's"},
		{1, 3},
		{2, 7.9},
		{3, 8},
		{4, 9.2},
		{4.5, 12},
		{5, 10.5},
		{6, 15},
		{7, 15.5},
		{8, 17},
	}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=STEYX(B2:B11,A2:A11)": "1.20118634668221",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	calcError := map[string][]string{
		"=STEYX()":             {"#VALUE!", "STEYX requires 2 arguments"},
		"=STEYX(B2:B11,A1:A9)": {"#N/A", "#N/A"},
		"=STEYX(B2,A2)":        {"#DIV/0!", "#DIV/0!"},
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
	}
}

func TestCalcTTEST(t *testing.T) {
	cellData := [][]interface{}{
		{4, 8, nil, 1, 1},
		{5, 3, nil, 1, 1},
		{2, 7},
		{5, 3},
		{8, 5},
		{9, 2},
		{3, 2},
		{2, 7},
		{3, 9},
		{8, 4},
		{9, 4},
		{5, 7},
	}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=TTEST(A1:A12,B1:B12,1,1)":  "0.44907068944428",
		"=TTEST(A1:A12,B1:B12,1,2)":  "0.436717306029283",
		"=TTEST(A1:A12,B1:B12,1,3)":  "0.436722015384755",
		"=TTEST(A1:A12,B1:B12,2,1)":  "0.898141378888559",
		"=TTEST(A1:A12,B1:B12,2,2)":  "0.873434612058567",
		"=TTEST(A1:A12,B1:B12,2,3)":  "0.873444030769511",
		"=T.TEST(A1:A12,B1:B12,1,1)": "0.44907068944428",
		"=T.TEST(A1:A12,B1:B12,1,2)": "0.436717306029283",
		"=T.TEST(A1:A12,B1:B12,1,3)": "0.436722015384755",
		"=T.TEST(A1:A12,B1:B12,2,1)": "0.898141378888559",
		"=T.TEST(A1:A12,B1:B12,2,2)": "0.873434612058567",
		"=T.TEST(A1:A12,B1:B12,2,3)": "0.873444030769511",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	calcError := map[string][]string{
		"=TTEST()":                      {"#VALUE!", "TTEST requires 4 arguments"},
		"=TTEST(\"\",B1:B12,1,1)":       {"#NUM!", "#NUM!"},
		"=TTEST(A1:A12,\"\",1,1)":       {"#NUM!", "#NUM!"},
		"=TTEST(A1:A12,B1:B12,\"\",1)":  {"#VALUE!", "#VALUE!"},
		"=TTEST(A1:A12,B1:B12,1,\"\")":  {"#VALUE!", "#VALUE!"},
		"=TTEST(A1:A12,B1:B12,0,1)":     {"#NUM!", "#NUM!"},
		"=TTEST(A1:A12,B1:B12,1,0)":     {"#NUM!", "#NUM!"},
		"=TTEST(A1:A2,B1:B1,1,1)":       {"#N/A", "#N/A"},
		"=TTEST(A13:A14,B13:B14,1,1)":   {"#NUM!", "#NUM!"},
		"=TTEST(A12:A13,B12:B13,1,1)":   {"#DIV/0!", "#DIV/0!"},
		"=TTEST(A13:A14,B13:B14,1,2)":   {"#NUM!", "#NUM!"},
		"=TTEST(D1:D4,E1:E4,1,3)":       {"#NUM!", "#NUM!"},
		"=T.TEST()":                     {"#VALUE!", "T.TEST requires 4 arguments"},
		"=T.TEST(\"\",B1:B12,1,1)":      {"#NUM!", "#NUM!"},
		"=T.TEST(A1:A12,\"\",1,1)":      {"#NUM!", "#NUM!"},
		"=T.TEST(A1:A12,B1:B12,\"\",1)": {"#VALUE!", "#VALUE!"},
		"=T.TEST(A1:A12,B1:B12,1,\"\")": {"#VALUE!", "#VALUE!"},
		"=T.TEST(A1:A12,B1:B12,0,1)":    {"#NUM!", "#NUM!"},
		"=T.TEST(A1:A12,B1:B12,1,0)":    {"#NUM!", "#NUM!"},
		"=T.TEST(A1:A2,B1:B1,1,1)":      {"#N/A", "#N/A"},
		"=T.TEST(A13:A14,B13:B14,1,1)":  {"#NUM!", "#NUM!"},
		"=T.TEST(A12:A13,B12:B13,1,1)":  {"#DIV/0!", "#DIV/0!"},
		"=T.TEST(A13:A14,B13:B14,1,2)":  {"#NUM!", "#NUM!"},
		"=T.TEST(D1:D4,E1:E4,1,3)":      {"#NUM!", "#NUM!"},
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
	}
}

func TestCalcNETWORKDAYSandWORKDAY(t *testing.T) {
	cellData := [][]interface{}{
		{"05/01/2019", 43586, "text1"},
		{"09/13/2019", 43721, "text2"},
		{"10/01/2019", 43739},
		{"12/25/2019", 43824},
		{"01/01/2020", 43831},
		{"01/01/2020", 43831},
		{"01/24/2020", 43854},
		{"04/04/2020", 43925},
		{"05/01/2020", 43952},
		{"06/25/2020", 44007},
	}
	f := prepareCalcData(cellData)
	formulaList := map[string]string{
		"=NETWORKDAYS(\"01/01/2020\",\"09/12/2020\")":               "183",
		"=NETWORKDAYS(\"01/01/2020\",\"09/12/2020\",2)":             "183",
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"09/12/2020\")":          "183",
		"=NETWORKDAYS.INTL(\"09/12/2020\",\"01/01/2020\")":          "-183",
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"09/12/2020\",1)":        "183",
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"09/12/2020\",2)":        "184",
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"09/12/2020\",3)":        "184",
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"09/12/2020\",4)":        "183",
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"09/12/2020\",5)":        "182",
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"09/12/2020\",6)":        "182",
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"09/12/2020\",7)":        "182",
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"09/12/2020\",11)":       "220",
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"09/12/2020\",12)":       "220",
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"09/12/2020\",13)":       "220",
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"09/12/2020\",14)":       "219",
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"09/12/2020\",15)":       "219",
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"09/12/2020\",16)":       "219",
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"09/12/2020\",17)":       "219",
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"09/12/2020\",1,A1:A12)": "178",
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"09/12/2020\",1,B1:B12)": "178",
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"09/12/2020\",1,C1:C2)":  "183",
		"=WORKDAY(\"12/01/2015\",25)":                               "42374",
		"=WORKDAY(\"01/01/2020\",123,B1:B12)":                       "44006",
		"=WORKDAY.INTL(\"12/01/2015\",0)":                           "42339",
		"=WORKDAY.INTL(\"12/01/2015\",25)":                          "42374",
		"=WORKDAY.INTL(\"12/01/2015\",-25)":                         "42304",
		"=WORKDAY.INTL(\"12/01/2015\",25,1)":                        "42374",
		"=WORKDAY.INTL(\"12/01/2015\",25,2)":                        "42374",
		"=WORKDAY.INTL(\"12/01/2015\",25,3)":                        "42372",
		"=WORKDAY.INTL(\"12/01/2015\",25,4)":                        "42373",
		"=WORKDAY.INTL(\"12/01/2015\",25,5)":                        "42374",
		"=WORKDAY.INTL(\"12/01/2015\",25,6)":                        "42374",
		"=WORKDAY.INTL(\"12/01/2015\",25,7)":                        "42374",
		"=WORKDAY.INTL(\"12/01/2015\",25,11)":                       "42368",
		"=WORKDAY.INTL(\"12/01/2015\",25,12)":                       "42368",
		"=WORKDAY.INTL(\"12/01/2015\",25,13)":                       "42368",
		"=WORKDAY.INTL(\"12/01/2015\",25,14)":                       "42369",
		"=WORKDAY.INTL(\"12/01/2015\",25,15)":                       "42368",
		"=WORKDAY.INTL(\"12/01/2015\",25,16)":                       "42368",
		"=WORKDAY.INTL(\"12/01/2015\",25,17)":                       "42368",
		"=WORKDAY.INTL(\"12/01/2015\",25,\"0001100\")":              "42374",
		"=WORKDAY.INTL(\"01/01/2020\",-123,4)":                      "43659",
		"=WORKDAY.INTL(\"01/01/2020\",123,4,44010)":                 "44002",
		"=WORKDAY.INTL(\"01/01/2020\",-123,4,43640)":                "43659",
		"=WORKDAY.INTL(\"01/01/2020\",-123,4,43660)":                "43658",
		"=WORKDAY.INTL(\"01/01/2020\",-123,7,43660)":                "43657",
		"=WORKDAY.INTL(\"01/01/2020\",123,4,A1:A12)":                "44008",
		"=WORKDAY.INTL(\"01/01/2020\",123,4,B1:B12)":                "44008",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	calcError := map[string][]string{
		"=NETWORKDAYS()": {"#VALUE!", "NETWORKDAYS requires at least 2 arguments"},
		"=NETWORKDAYS(\"01/01/2020\",\"09/12/2020\",2,\"\")":             {"#VALUE!", "NETWORKDAYS requires at most 3 arguments"},
		"=NETWORKDAYS(\"\",\"09/12/2020\",2)":                            {"#VALUE!", "#VALUE!"},
		"=NETWORKDAYS(\"01/01/2020\",\"\",2)":                            {"#VALUE!", "#VALUE!"},
		"=NETWORKDAYS.INTL()":                                            {"#VALUE!", "NETWORKDAYS.INTL requires at least 2 arguments"},
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"09/12/2020\",4,A1:A12,\"\")": {"#VALUE!", "NETWORKDAYS.INTL requires at most 4 arguments"},
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"January 25, 100\",4)":        {"#VALUE!", "#VALUE!"},
		"=NETWORKDAYS.INTL(\"\",123,4,B1:B12)":                           {"#VALUE!", "#VALUE!"},
		"=NETWORKDAYS.INTL(\"01/01/2020\",123,\"000000x\")":              {"#VALUE!", "#VALUE!"},
		"=NETWORKDAYS.INTL(\"01/01/2020\",123,\"0000002\")":              {"#VALUE!", "#VALUE!"},
		"=NETWORKDAYS.INTL(\"January 25, 100\",123)":                     {"#VALUE!", "#VALUE!"},
		"=NETWORKDAYS.INTL(\"01/01/2020\",\"09/12/2020\",8)":             {"#VALUE!", "#VALUE!"},
		"=NETWORKDAYS.INTL(-1,123)":                                      {"#NUM!", "#NUM!"},
		"=WORKDAY()":                                                     {"#VALUE!", "WORKDAY requires at least 2 arguments"},
		"=WORKDAY(\"01/01/2020\",123,A1:A12,\"\")":                       {"#VALUE!", "WORKDAY requires at most 3 arguments"},
		"=WORKDAY(\"01/01/2020\",\"\",B1:B12)":                           {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=WORKDAY(\"\",123,B1:B12)":                                      {"#VALUE!", "#VALUE!"},
		"=WORKDAY(\"January 25, 100\",123)":                              {"#VALUE!", "#VALUE!"},
		"=WORKDAY(-1,123)":                                               {"#NUM!", "#NUM!"},
		"=WORKDAY.INTL()":                                                {"#VALUE!", "WORKDAY.INTL requires at least 2 arguments"},
		"=WORKDAY.INTL(\"01/01/2020\",123,4,A1:A12,\"\")":                {"#VALUE!", "WORKDAY.INTL requires at most 4 arguments"},
		"=WORKDAY.INTL(\"01/01/2020\",\"\",4,B1:B12)":                    {"#VALUE!", "strconv.ParseFloat: parsing \"\": invalid syntax"},
		"=WORKDAY.INTL(\"\",123,4,B1:B12)":                               {"#VALUE!", "#VALUE!"},
		"=WORKDAY.INTL(\"01/01/2020\",123,\"\",B1:B12)":                  {"#VALUE!", "#VALUE!"},
		"=WORKDAY.INTL(\"01/01/2020\",123,\"000000x\")":                  {"#VALUE!", "#VALUE!"},
		"=WORKDAY.INTL(\"01/01/2020\",123,\"0000002\")":                  {"#VALUE!", "#VALUE!"},
		"=WORKDAY.INTL(\"January 25, 100\",123)":                         {"#VALUE!", "#VALUE!"},
		"=WORKDAY.INTL(-1,123)":                                          {"#NUM!", "#NUM!"},
	}
	for formula, expected := range calcError {
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", formula))
		result, err := f.CalcCellValue("Sheet1", "C1")
		assert.Equal(t, expected[0], result, formula)
		assert.EqualError(t, err, expected[1], formula)
	}
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

func TestCalcGetBetaHelperContFrac(t *testing.T) {
	assert.Equal(t, 1.0, getBetaHelperContFrac(1, 0, 1))
}

func TestCalcGetBetaDistPDF(t *testing.T) {
	assert.Equal(t, 0.0, getBetaDistPDF(0.5, 2000, 3))
	assert.Equal(t, 0.0, getBetaDistPDF(0, 1, 0))
}

func TestCalcD1mach(t *testing.T) {
	assert.Equal(t, 0.0, d1mach(6))
}

func TestCalcChebyshevInit(t *testing.T) {
	assert.Equal(t, 0, chebyshevInit(0, 0, nil))
	assert.Equal(t, 0, chebyshevInit(1, 0, []float64{0}))
}

func TestCalcChebyshevEval(t *testing.T) {
	assert.True(t, math.IsNaN(chebyshevEval(0, 0, nil)))
}

func TestCalcLgammacor(t *testing.T) {
	assert.True(t, math.IsNaN(lgammacor(9)))
	assert.Equal(t, 4.930380657631324e-32, lgammacor(3.7451940309632633e+306))
	assert.Equal(t, 8.333333333333334e-10, lgammacor(10e+07))
}

func TestCalcLgammaerr(t *testing.T) {
	assert.True(t, math.IsNaN(logrelerr(-2)))
}

func TestCalcLogBeta(t *testing.T) {
	assert.True(t, math.IsNaN(logBeta(-1, -1)))
	assert.Equal(t, math.MaxFloat64, logBeta(0, 0))
}

func TestCalcBetainvProbIterator(t *testing.T) {
	assert.Equal(t, 1.0, betainvProbIterator(1, 1, 1, 1, 1, 1, 1, 1, 1))
}

func TestNestedFunctionsWithOperators(t *testing.T) {
	f := NewFile()
	formulaList := map[string]string{
		"=LEN(\"KEEP\")":                                                   "4",
		"=LEN(\"REMOVEKEEP\") - LEN(\"REMOVE\")":                           "4",
		"=RIGHT(\"REMOVEKEEP\", 4)":                                        "KEEP",
		"=RIGHT(\"REMOVEKEEP\", 10 - 6))":                                  "KEEP",
		"=RIGHT(\"REMOVEKEEP\", LEN(\"REMOVEKEEP\") - 6)":                  "KEEP",
		"=RIGHT(\"REMOVEKEEP\", LEN(\"REMOVEKEEP\") - LEN(\"REMOV\") - 1)": "KEEP",
		"=RIGHT(\"REMOVEKEEP\", 10 - LEN(\"REMOVE\"))":                     "KEEP",
		"=RIGHT(\"REMOVEKEEP\", LEN(\"REMOVEKEEP\") - LEN(\"REMOVE\"))":    "KEEP",
	}
	for formula, expected := range formulaList {
		assert.NoError(t, f.SetCellFormula("Sheet1", "E1", formula))
		result, err := f.CalcCellValue("Sheet1", "E1")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
}

func TestFormulaRawCellValueOption(t *testing.T) {
	f := NewFile()
	rawTest := []struct {
		value    string
		raw      bool
		expected string
	}{
		{"=VALUE(\"1.0E-07\")", false, "0.00"},
		{"=VALUE(\"1.0E-07\")", true, "0.0000001"},
		{"=\"text\"", false, "$text"},
		{"=\"text\"", true, "text"},
		{"=\"10e3\"", false, "$10e3"},
		{"=\"10e3\"", true, "10e3"},
		{"=\"10\" & \"e3\"", false, "$10e3"},
		{"=\"10\" & \"e3\"", true, "10e3"},
		{"=10e3", false, "10000.00"},
		{"=10e3", true, "10000"},
		{"=\"1111111111111111\"", false, "$1111111111111111"},
		{"=\"1111111111111111\"", true, "1111111111111111"},
		{"=1111111111111111", false, "1111111111111110.00"},
		{"=1111111111111111", true, "1.11111111111111E+15"},
		{"=1444.00000000003", false, "1444.00"},
		{"=1444.00000000003", true, "1444.00000000003"},
		{"=1444.000000000003", false, "1444.00"},
		{"=1444.000000000003", true, "1444"},
		{"=ROUND(1444.00000000000003,2)", false, "1444.00"},
		{"=ROUND(1444.00000000000003,2)", true, "1444"},
	}
	exp := "0.00;0.00;;$@"
	styleID, err := f.NewStyle(&Style{CustomNumFmt: &exp})
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "A1", "A1", styleID))
	for _, test := range rawTest {
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", test.value))
		val, err := f.CalcCellValue("Sheet1", "A1", Options{RawCellValue: test.raw})
		assert.NoError(t, err)
		assert.Equal(t, test.expected, val)
	}
}

func TestFormulaArgToToken(t *testing.T) {
	assert.Equal(t,
		efp.Token{
			TType:    efp.TokenTypeOperand,
			TSubType: efp.TokenSubTypeLogical,
			TValue:   "TRUE",
		},
		formulaArgToToken(newBoolFormulaArg(true)),
	)
}

func TestPrepareTrendGrowth(t *testing.T) {
	assert.Equal(t, [][]float64(nil), prepareTrendGrowthMtxX([][]float64{{0, 0}, {0, 0}}))
	assert.Equal(t, [][]float64(nil), prepareTrendGrowthMtxY(false, [][]float64{{0, 0}, {0, 0}}))
	info, err := prepareTrendGrowth(false, [][]float64{{0, 0}, {0, 0}}, [][]float64{{0, 0}, {0, 0}})
	assert.Nil(t, info)
	assert.Equal(t, newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM), err)
}

func TestCalcColRowQRDecomposition(t *testing.T) {
	assert.False(t, calcRowQRDecomposition([][]float64{{0, 0}, {0, 0}}, []float64{0, 0}, 1, 0))
	assert.False(t, calcColQRDecomposition([][]float64{{0, 0}, {0, 0}}, []float64{0, 0}, 1, 0))
}

func TestCalcCellResolver(t *testing.T) {
	f := NewFile()
	// Test reference a cell multiple times in a formula
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "VALUE1"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "=A1"))
	for formula, expected := range map[string]string{
		"=CONCATENATE(A1,\"_\",A1)": "VALUE1_VALUE1",
		"=CONCATENATE(A1,\"_\",A2)": "VALUE1_VALUE1",
		"=CONCATENATE(A2,\"_\",A2)": "VALUE1_VALUE1",
	} {
		assert.NoError(t, f.SetCellFormula("Sheet1", "A3", formula))
		result, err := f.CalcCellValue("Sheet1", "A3")
		assert.NoError(t, err, formula)
		assert.Equal(t, expected, result, formula)
	}
	// Test calculates formula that contains a nested argument function which returns a numeric result
	f = NewFile()
	for _, cell := range []string{"A1", "B2", "B3", "B4"} {
		assert.NoError(t, f.SetCellValue("Sheet1", cell, "ABC"))
	}
	for cell, formula := range map[string]string{
		"A2": "IF(B2<>\"\",MAX(A1:A1)+1,\"\")",
		"A3": "IF(B3<>\"\",MAX(A1:A2)+1,\"\")",
		"A4": "IF(B4<>\"\",MAX(A1:A3)+1,\"\")",
	} {
		assert.NoError(t, f.SetCellFormula("Sheet1", cell, formula))
	}
	for cell, expected := range map[string]string{"A2": "1", "A3": "2", "A4": "3"} {
		result, err := f.CalcCellValue("Sheet1", cell)
		assert.NoError(t, err)
		assert.Equal(t, expected, result)
	}
	// Test calculates formula that reference date and error type cells
	assert.NoError(t, f.SetCellValue("Sheet1", "C1", "20200208T080910.123"))
	assert.NoError(t, f.SetCellValue("Sheet1", "C2", "2020-07-10 15:00:00.000"))
	assert.NoError(t, f.SetCellValue("Sheet1", "C3", formulaErrorDIV))
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).SheetData.Row[0].C[2].T = "d"
	ws.(*xlsxWorksheet).SheetData.Row[0].C[2].V = "20200208T080910.123"
	ws.(*xlsxWorksheet).SheetData.Row[1].C[2].T = "d"
	ws.(*xlsxWorksheet).SheetData.Row[1].C[2].V = "2020-07-10 15:00:00.000"
	ws.(*xlsxWorksheet).SheetData.Row[2].C[2].T = "e"
	ws.(*xlsxWorksheet).SheetData.Row[2].C[2].V = formulaErrorDIV
	for _, tbl := range [][]string{
		{"D1", "=SUM(C1,1)", "43870.3397004977"},
		{"D2", "=LEN(C2)", "23"},
		{"D3", "=IFERROR(C3,TRUE)", "TRUE"},
	} {
		assert.NoError(t, f.SetCellFormula("Sheet1", tbl[0], tbl[1]))
		result, err := f.CalcCellValue("Sheet1", tbl[0])
		assert.NoError(t, err)
		assert.Equal(t, tbl[2], result)
	}
	// Test calculates formula that reference invalid cell
	assert.NoError(t, f.SetCellValue("Sheet1", "E1", "E1"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "F1", "=LEN(E1)"))
	f.SharedStrings = nil
	f.Pkg.Store(defaultXMLPathSharedStrings, MacintoshCyrillicCharset)
	_, err := f.CalcCellValue("Sheet1", "F1")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
}

func TestEvalInfixExp(t *testing.T) {
	f := NewFile()
	arg, err := f.evalInfixExp(nil, "Sheet1", "A1", []efp.Token{
		{TSubType: efp.TokenSubTypeRange, TValue: "1A"},
	})
	assert.Equal(t, arg, newEmptyFormulaArg())
	assert.Equal(t, formulaErrorNAME, err.Error())
}

func TestParseToken(t *testing.T) {
	f := NewFile()
	assert.Equal(t, formulaErrorNAME, f.parseToken(nil, "Sheet1",
		efp.Token{TSubType: efp.TokenSubTypeRange, TValue: "1A"}, nil, nil,
	).Error())
}
