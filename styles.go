package excelize

import (
	"encoding/json"
	"encoding/xml"
	"fmt"
	"math"
	"strconv"
	"strings"
)

// Excel styles can reference number formats that are built-in, all of which
// have an id less than 164. This is a possibly incomplete list comprised of as
// many of them as I could find.
var builtInNumFmt = map[int]string{
	0:  "general",
	1:  "0",
	2:  "0.00",
	3:  "#,##0",
	4:  "#,##0.00",
	9:  "0%",
	10: "0.00%",
	11: "0.00e+00",
	12: "# ?/?",
	13: "# ??/??",
	14: "mm-dd-yy",
	15: "d-mmm-yy",
	16: "d-mmm",
	17: "mmm-yy",
	18: "h:mm am/pm",
	19: "h:mm:ss am/pm",
	20: "h:mm",
	21: "h:mm:ss",
	22: "m/d/yy h:mm",
	37: "#,##0 ;(#,##0)",
	38: "#,##0 ;[red](#,##0)",
	39: "#,##0.00;(#,##0.00)",
	40: "#,##0.00;[red](#,##0.00)",
	41: `_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)`,
	42: `_("$"* #,##0_);_("$* \(#,##0\);_("$"* "-"_);_(@_)`,
	43: `_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)`,
	44: `_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)`,
	45: "mm:ss",
	46: "[h]:mm:ss",
	47: "mmss.0",
	48: "##0.0e+0",
	49: "@",
}

// langNumFmt defined number format code with unicode in different language.
var langNumFmt = map[string]map[int]string{
	"zh-tw":         zhtwNumFmt,
	"zh-cn":         zhcnNumFmt,
	"zh-tw_unicode": zhtwUnicodeNumFmt,
	"zh-cn_unicode": zhcnUnicodeNumFmt,
	"ja-jp":         jajpNumFmt,
	"ko-kr":         kokrNumFmt,
	"ja-jp_unicode": jajpUnicodeNumFmt,
	"ko-kr_unicode": kokrUnicodeNumFmt,
	"th-th":         ththNumFmt,
	"th-th_unicode": ththUnicodeNumFmt,
}

// zh-tw format code.
var zhtwNumFmt = map[int]string{
	27: "[$-404]e/m/d",
	28: `[$-404]e"年"m"月"d"日"`,
	29: `[$-404]e"年"m"月"d"日"`,
	30: "m/d/yy",
	31: `yyyy"年"m"月"d"日"`,
	32: `hh"時"mm"分"`,
	33: `hh"時"mm"分"ss"秒"`,
	34: `上午/下午 hh"時"mm"分"`,
	35: `上午/下午 hh"時"mm"分"ss"秒"`,
	36: "[$-404]e/m/d",
	50: "[$-404]e/m/d",
	51: `[$-404]e"年"m"月"d"日"`,
	52: `上午/下午 hh"時"mm"分"`,
	53: `上午/下午 hh"時"mm"分"ss"秒"`,
	54: `[$-404]e"年"m"月"d"日"`,
	55: `上午/下午 hh"時"mm"分"`,
	56: `上午/下午 hh"時"mm"分"ss"秒"`,
	57: "[$-404]e/m/d",
	58: `[$-404]e"年"m"月"d"日"`,
}

// zh-cn format code.
var zhcnNumFmt = map[int]string{
	27: `yyyy"年"m"月"`,
	28: `m"月"d"日"`,
	29: `m"月"d"日"`,
	30: "m-d-yy",
	31: `yyyy"年"m"月"d"日"`,
	32: `h"时"mm"分"`,
	33: `h"时"mm"分"ss"秒"`,
	34: `上午/下午 h"时"mm"分"`,
	35: `上午/下午 h"时"mm"分"ss"秒"`,
	36: `yyyy"年"m"月"`,
	50: `yyyy"年"m"月"`,
	51: `m"月"d"日"`,
	52: `yyyy"年"m"月"`,
	53: `m"月"d"日"`,
	54: `m"月"d"日"`,
	55: `上午/下午 h"时"mm"分"`,
	56: `上午/下午 h"时"mm"分"ss"秒"`,
	57: `yyyy"年"m"月"`,
	58: `m"月"d"日"`,
}

// zh-tw format code (with unicode values provided for language glyphs where
// they occur).
var zhtwUnicodeNumFmt = map[int]string{
	27: "[$-404]e/m/d",
	28: `[$-404]e"5E74"m"6708"d"65E5"`,
	29: `[$-404]e"5E74"m"6708"d"65E5"`,
	30: "m/d/yy",
	31: `yyyy"5E74"m"6708"d"65E5"`,
	32: `hh"6642"mm"5206"`,
	33: `hh"6642"mm"5206"ss"79D2"`,
	34: `4E0A5348/4E0B5348hh"6642"mm"5206"`,
	35: `4E0A5348/4E0B5348hh"6642"mm"5206"ss"79D2"`,
	36: "[$-404]e/m/d",
	50: "[$-404]e/m/d",
	51: `[$-404]e"5E74"m"6708"d"65E5"`,
	52: `4E0A5348/4E0B5348hh"6642"mm"5206"`,
	53: `4E0A5348/4E0B5348hh"6642"mm"5206"ss"79D2"`,
	54: `[$-404]e"5E74"m"6708"d"65E5"`,
	55: `4E0A5348/4E0B5348hh"6642"mm"5206"`,
	56: `4E0A5348/4E0B5348hh"6642"mm"5206"ss"79D2"`,
	57: "[$-404]e/m/d",
	58: `[$-404]e"5E74"m"6708"d"65E5"`,
}

// zh-cn format code (with unicode values provided for language glyphs where
// they occur).
var zhcnUnicodeNumFmt = map[int]string{
	27: `yyyy"5E74"m"6708"`,
	28: `m"6708"d"65E5"`,
	29: `m"6708"d"65E5"`,
	30: "m-d-yy",
	31: `yyyy"5E74"m"6708"d"65E5"`,
	32: `h"65F6"mm"5206"`,
	33: `h"65F6"mm"5206"ss"79D2"`,
	34: `4E0A5348/4E0B5348h"65F6"mm"5206"`,
	35: `4E0A5348/4E0B5348h"65F6"mm"5206"ss"79D2"`,
	36: `yyyy"5E74"m"6708"`,
	50: `yyyy"5E74"m"6708"`,
	51: `m"6708"d"65E5"`,
	52: `yyyy"5E74"m"6708"`,
	53: `m"6708"d"65E5"`,
	54: `m"6708"d"65E5"`,
	55: `4E0A5348/4E0B5348h"65F6"mm"5206"`,
	56: `4E0A5348/4E0B5348h"65F6"mm"5206"ss"79D2"`,
	57: `yyyy"5E74"m"6708"`,
	58: `m"6708"d"65E5"`,
}

// ja-jp format code.
var jajpNumFmt = map[int]string{
	27: "[$-411]ge.m.d",
	28: `[$-411]ggge"年"m"月"d"日"`,
	29: `[$-411]ggge"年"m"月"d"日"`,
	30: "m/d/yy",
	31: `yyyy"年"m"月"d"日"`,
	32: `h"時"mm"分"`,
	33: `h"時"mm"分"ss"秒"`,
	34: `yyyy"年"m"月"`,
	35: `m"月"d"日"`,
	36: "[$-411]ge.m.d",
	50: "[$-411]ge.m.d",
	51: `[$-411]ggge"年"m"月"d"日"`,
	52: `yyyy"年"m"月"`,
	53: `m"月"d"日"`,
	54: `[$-411]ggge"年"m"月"d"日"`,
	55: `yyyy"年"m"月"`,
	56: `m"月"d"日"`,
	57: "[$-411]ge.m.d",
	58: `[$-411]ggge"年"m"月"d"日"`,
}

// ko-kr format code.
var kokrNumFmt = map[int]string{
	27: `yyyy"年" mm"月" dd"日"`,
	28: "mm-dd",
	29: "mm-dd",
	30: "mm-dd-yy",
	31: `yyyy"년" mm"월" dd"일"`,
	32: `h"시" mm"분"`,
	33: `h"시" mm"분" ss"초"`,
	34: `yyyy-mm-dd`,
	35: `yyyy-mm-dd`,
	36: `yyyy"年" mm"月" dd"日"`,
	50: `yyyy"年" mm"月" dd"日"`,
	51: "mm-dd",
	52: "yyyy-mm-dd",
	53: "yyyy-mm-dd",
	54: "mm-dd",
	55: "yyyy-mm-dd",
	56: "yyyy-mm-dd",
	57: `yyyy"年" mm"月" dd"日"`,
	58: "mm-dd",
}

// ja-jp format code (with unicode values provided for language glyphs where
// they occur).
var jajpUnicodeNumFmt = map[int]string{
	27: "[$-411]ge.m.d",
	28: `[$-411]ggge"5E74"m"6708"d"65E5"`,
	29: `[$-411]ggge"5E74"m"6708"d"65E5"`,
	30: "m/d/yy",
	31: `yyyy"5E74"m"6708"d"65E5"`,
	32: `h"6642"mm"5206"`,
	33: `h"6642"mm"5206"ss"79D2"`,
	34: `yyyy"5E74"m"6708"`,
	35: `m"6708"d"65E5"`,
	36: "[$-411]ge.m.d",
	50: "[$-411]ge.m.d",
	51: `[$-411]ggge"5E74"m"6708"d"65E5"`,
	52: `yyyy"5E74"m"6708"`,
	53: `m"6708"d"65E5"`,
	54: `[$-411]ggge"5E74"m"6708"d"65E5"`,
	55: `yyyy"5E74"m"6708"`,
	56: `m"6708"d"65E5"`,
	57: "[$-411]ge.m.d",
	58: `[$-411]ggge"5E74"m"6708"d"65E5"`,
}

// ko-kr format code (with unicode values provided for language glyphs where
// they occur).
var kokrUnicodeNumFmt = map[int]string{
	27: `yyyy"5E74" mm"6708" dd"65E5"`,
	28: "mm-dd",
	29: "mm-dd",
	30: "mm-dd-yy",
	31: `yyyy"B144" mm"C6D4" dd"C77C"`,
	32: `h"C2DC" mm"BD84"`,
	33: `h"C2DC" mm"BD84" ss"CD08"`,
	34: "yyyy-mm-dd",
	35: "yyyy-mm-dd",
	36: `yyyy"5E74" mm"6708" dd"65E5"`,
	50: `yyyy"5E74" mm"6708" dd"65E5"`,
	51: "mm-dd",
	52: "yyyy-mm-dd",
	53: "yyyy-mm-dd",
	54: "mm-dd",
	55: "yyyy-mm-dd",
	56: "yyyy-mm-dd",
	57: `yyyy"5E74" mm"6708" dd"65E5"`,
	58: "mm-dd",
}

// th-th format code.
var ththNumFmt = map[int]string{
	59: "t0",
	60: "t0.00",
	61: "t#,##0",
	62: "t#,##0.00",
	67: "t0%",
	68: "t0.00%",
	69: "t# ?/?",
	70: "t# ??/??",
	71: "ว/ด/ปปปป",
	72: "ว-ดดด-ปป",
	73: "ว-ดดด",
	74: "ดดด-ปป",
	75: "ช:นน",
	76: "ช:นน:ทท",
	77: "ว/ด/ปปปป ช:นน",
	78: "นน:ทท",
	79: "[ช]:นน:ทท",
	80: "นน:ทท.0",
	81: "d/m/bb",
}

// th-th format code (with unicode values provided for language glyphs where
// they occur).
var ththUnicodeNumFmt = map[int]string{
	59: "t0",
	60: "t0.00",
	61: "t#,##0",
	62: "t#,##0.00",
	67: "t0%",
	68: "t0.00%",
	69: "t# ?/?",
	70: "t# ??/??",
	71: "0E27/0E14/0E1B0E1B0E1B0E1B",
	72: "0E27-0E140E140E14-0E1B0E1B",
	73: "0E27-0E140E140E14",
	74: "0E140E140E14-0E1B0E1B",
	75: "0E0A:0E190E19",
	76: "0E0A:0E190E19:0E170E17",
	77: "0E27/0E14/0E1B0E1B0E1B0E1B 0E0A:0E190E19",
	78: "0E190E19:0E170E17",
	79: "[0E0A]:0E190E19:0E170E17",
	80: "0E190E19:0E170E17.0",
	81: "d/m/bb",
}

// currencyNumFmt defined the currency number format map.
var currencyNumFmt = map[int]string{
	164: `"CN¥",##0.00`,
	165: "[$$-409]#,##0.00",
	166: "[$$-45C]#,##0.00",
	167: "[$$-1004]#,##0.00",
	168: "[$$-404]#,##0.00",
	169: "[$$-C09]#,##0.00",
	170: "[$$-2809]#,##0.00",
	171: "[$$-1009]#,##0.00",
	172: "[$$-2009]#,##0.00",
	173: "[$$-1409]#,##0.00",
	174: "[$$-4809]#,##0.00",
	175: "[$$-2C09]#,##0.00",
	176: "[$$-2409]#,##0.00",
	177: "[$$-1000]#,##0.00",
	178: `#,##0.00\ [$$-C0C]`,
	179: "[$$-475]#,##0.00",
	180: "[$$-83E]#,##0.00",
	181: `[$$-86B]\ #,##0.00`,
	182: `[$$-340A]\ #,##0.00`,
	183: "[$$-240A]#,##0.00",
	184: `[$$-300A]\ #,##0.00`,
	185: "[$$-440A]#,##0.00",
	186: "[$$-80A]#,##0.00",
	187: "[$$-500A]#,##0.00",
	188: "[$$-540A]#,##0.00",
	189: `[$$-380A]\ #,##0.00`,
}

// builtInNumFmtFunc defined the format conversion functions map. Partial format
// code doesn't support currently and will return original string.
var builtInNumFmtFunc = map[int]func(i int, v string) string{
	0:  formatToString,
	1:  formatToInt,
	2:  formatToFloat,
	3:  formatToInt,
	4:  formatToFloat,
	9:  formatToC,
	10: formatToD,
	11: formatToE,
	12: formatToString, // Doesn't support currently
	13: formatToString, // Doesn't support currently
	14: parseTime,
	15: parseTime,
	16: parseTime,
	17: parseTime,
	18: parseTime,
	19: parseTime,
	20: parseTime,
	21: parseTime,
	22: parseTime,
	37: formatToA,
	38: formatToA,
	39: formatToB,
	40: formatToB,
	41: formatToString, // Doesn't support currently
	42: formatToString, // Doesn't support currently
	43: formatToString, // Doesn't support currently
	44: formatToString, // Doesn't support currently
	45: parseTime,
	46: parseTime,
	47: parseTime,
	48: formatToE,
	49: formatToString,
}

// formatToString provides function to return original string by given built-in
// number formats code and cell string.
func formatToString(i int, v string) string {
	return v
}

// formatToInt provides function to convert original string to integer format as
// string type by given built-in number formats code and cell string.
func formatToInt(i int, v string) string {
	f, err := strconv.ParseFloat(v, 64)
	if err != nil {
		return v
	}
	return fmt.Sprintf("%d", int(f))
}

// formatToFloat provides function to convert original string to float format as
// string type by given built-in number formats code and cell string.
func formatToFloat(i int, v string) string {
	f, err := strconv.ParseFloat(v, 64)
	if err != nil {
		return v
	}
	return fmt.Sprintf("%.2f", f)
}

// formatToA provides function to convert original string to special format as
// string type by given built-in number formats code and cell string.
func formatToA(i int, v string) string {
	f, err := strconv.ParseFloat(v, 64)
	if err != nil {
		return v
	}
	if f < 0 {
		t := int(math.Abs(f))
		return fmt.Sprintf("(%d)", t)
	}
	t := int(f)
	return fmt.Sprintf("%d", t)
}

// formatToB provides function to convert original string to special format as
// string type by given built-in number formats code and cell string.
func formatToB(i int, v string) string {
	f, err := strconv.ParseFloat(v, 64)
	if err != nil {
		return v
	}
	if f < 0 {
		return fmt.Sprintf("(%.2f)", f)
	}
	return fmt.Sprintf("%.2f", f)
}

// formatToC provides function to convert original string to special format as
// string type by given built-in number formats code and cell string.
func formatToC(i int, v string) string {
	f, err := strconv.ParseFloat(v, 64)
	if err != nil {
		return v
	}
	f = f * 100
	return fmt.Sprintf("%d%%", int(f))
}

// formatToD provides function to convert original string to special format as
// string type by given built-in number formats code and cell string.
func formatToD(i int, v string) string {
	f, err := strconv.ParseFloat(v, 64)
	if err != nil {
		return v
	}
	f = f * 100
	return fmt.Sprintf("%.2f%%", f)
}

// formatToE provides function to convert original string to special format as
// string type by given built-in number formats code and cell string.
func formatToE(i int, v string) string {
	f, err := strconv.ParseFloat(v, 64)
	if err != nil {
		return v
	}
	return fmt.Sprintf("%.e", f)
}

// parseTime provides function to returns a string parsed using time.Time.
// Replace Excel placeholders with Go time placeholders. For example, replace
// yyyy with 2006. These are in a specific order, due to the fact that m is used
// in month, minute, and am/pm. It would be easier to fix that with regular
// expressions, but if it's possible to keep this simple it would be easier to
// maintain. Full-length month and days (e.g. March, Tuesday) have letters in
// them that would be replaced by other characters below (such as the 'h' in
// March, or the 'd' in Tuesday) below. First we convert them to arbitrary
// characters unused in Excel Date formats, and then at the end, turn them to
// what they should actually be.
func parseTime(i int, v string) string {
	f, err := strconv.ParseFloat(v, 64)
	if err != nil {
		return v
	}
	val := timeFromExcelTime(f, false)
	format := builtInNumFmt[i]

	replacements := []struct{ xltime, gotime string }{
		{"yyyy", "2006"},
		{"yy", "06"},
		{"mmmm", "%%%%"},
		{"dddd", "&&&&"},
		{"dd", "02"},
		{"d", "2"},
		{"mmm", "Jan"},
		{"mmss", "0405"},
		{"ss", "05"},
		{"hh", "15"},
		{"h", "3"},
		{"mm:", "04:"},
		{":mm", ":04"},
		{"mm", "01"},
		{"am/pm", "pm"},
		{"m/", "1/"},
		{"%%%%", "January"},
		{"&&&&", "Monday"},
	}
	for _, repl := range replacements {
		format = strings.Replace(format, repl.xltime, repl.gotime, 1)
	}
	// If the hour is optional, strip it out, along with the possible dangling
	// colon that would remain.
	if val.Hour() < 1 {
		format = strings.Replace(format, "]:", "]", 1)
		format = strings.Replace(format, "[3]", "", 1)
		format = strings.Replace(format, "[15]", "", 1)
	} else {
		format = strings.Replace(format, "[3]", "3", 1)
		format = strings.Replace(format, "[15]", "15", 1)
	}
	return val.Format(format)
}

// stylesReader provides function to get the pointer to the structure after
// deserialization of xl/styles.xml.
func (f *File) stylesReader() *xlsxStyleSheet {
	if f.Styles == nil {
		var styleSheet xlsxStyleSheet
		xml.Unmarshal([]byte(f.readXML("xl/styles.xml")), &styleSheet)
		f.Styles = &styleSheet
	}
	return f.Styles
}

// styleSheetWriter provides function to save xl/styles.xml after serialize
// structure.
func (f *File) styleSheetWriter() {
	if f.Styles != nil {
		output, _ := xml.Marshal(f.Styles)
		f.saveFileList("xl/styles.xml", replaceWorkSheetsRelationshipsNameSpace(string(output)))
	}
}

// parseFormatStyleSet provides function to parse the format settings of the
// borders.
func parseFormatStyleSet(style string) (*formatCellStyle, error) {
	format := formatCellStyle{
		DecimalPlaces: 2,
	}
	err := json.Unmarshal([]byte(style), &format)
	return &format, err
}

// NewStyle provides function to create style for cells by given style format.
// Note that the color field uses RGB color code.
//
// The following shows the border styles sorted by excelize index number:
//
//    | Index | Name          | Weight | Style       |
//    +-------+---------------+--------+-------------+
//    | 0     | None          | 0      |             |
//    | 1     | Continuous    | 1      | ----------- |
//    | 2     | Continuous    | 2      | ----------- |
//    | 3     | Dash          | 1      | - - - - - - |
//    | 4     | Dot           | 1      | . . . . . . |
//    | 5     | Continuous    | 3      | ----------- |
//    | 6     | Double        | 3      | =========== |
//    | 7     | Continuous    | 0      | ----------- |
//    | 8     | Dash          | 2      | - - - - - - |
//    | 9     | Dash Dot      | 1      | - . - . - . |
//    | 10    | Dash Dot      | 2      | - . - . - . |
//    | 11    | Dash Dot Dot  | 1      | - . . - . . |
//    | 12    | Dash Dot Dot  | 2      | - . . - . . |
//    | 13    | SlantDash Dot | 2      | / - . / - . |
//
// The following shows the borders in the order shown in the Excel dialog:
//
//    | Index | Style       | Index | Style       |
//    +-------+-------------+-------+-------------+
//    | 0     | None        | 12    | - . . - . . |
//    | 7     | ----------- | 13    | / - . / - . |
//    | 4     | . . . . . . | 10    | - . - . - . |
//    | 11    | - . . - . . | 8     | - - - - - - |
//    | 9     | - . - . - . | 2     | ----------- |
//    | 3     | - - - - - - | 5     | ----------- |
//    | 1     | ----------- | 6     | =========== |
//
// The following shows the shading styles sorted by excelize index number:
//
//    | Index | Style           | Index | Style           |
//    +-------+-----------------+-------+-----------------+
//    | 0     | Horizontal      | 3     | Diagonal down   |
//    | 1     | Vertical        | 4     | From corner     |
//    | 2     | Diagonal Up     | 5     | From center     |
//
// The following shows the patterns styles sorted by excelize index number:
//
//    | Index | Style           | Index | Style           |
//    +-------+-----------------+-------+-----------------+
//    | 0     | None            | 10    | darkTrellis     |
//    | 1     | solid           | 11    | lightHorizontal |
//    | 2     | mediumGray      | 12    | lightVertical   |
//    | 3     | darkGray        | 13    | lightDown       |
//    | 4     | lightGray       | 14    | lightUp         |
//    | 5     | darkHorizontal  | 15    | lightGrid       |
//    | 6     | darkVertical    | 16    | lightTrellis    |
//    | 7     | darkDown        | 17    | gray125         |
//    | 8     | darkUp          | 18    | gray0625        |
//    | 9     | darkGrid        |       |                 |
//
// The following the type of horizontal alignment in cells:
//
//    | Style            |
//    +------------------+
//    | left             |
//    | center           |
//    | right            |
//    | fill             |
//    | justify          |
//    | centerContinuous |
//    | distributed      |
//
// The following the type of vertical alignment in cells:
//
//    | Style            |
//    +------------------+
//    | top              |
//    | center           |
//    | justify          |
//    | distributed      |
//
// The following the type of font underline style:
//
//    | Style            |
//    +------------------+
//    | single           |
//    | double           |
//
// Excel's built-in all languages formats are shown in the following table:
//
//    | Index | Format String                                      |
//    +-------+----------------------------------------------------+
//    | 0     | General                                            |
//    | 1     | 0                                                  |
//    | 2     | 0.00                                               |
//    | 3     | #,##0                                              |
//    | 4     | #,##0.00                                           |
//    | 5     | ($#,##0_);($#,##0)                                 |
//    | 6     | ($#,##0_);[Red]($#,##0)                            |
//    | 7     | ($#,##0.00_);($#,##0.00)                           |
//    | 8     | ($#,##0.00_);[Red]($#,##0.00)                      |
//    | 9     | 0%                                                 |
//    | 10    | 0.00%                                              |
//    | 11    | 0.00E+00                                           |
//    | 12    | # ?/?                                              |
//    | 13    | # ??/??                                            |
//    | 14    | m/d/yy                                             |
//    | 15    | d-mmm-yy                                           |
//    | 16    | d-mmm                                              |
//    | 17    | mmm-yy                                             |
//    | 18    | h:mm AM/PM                                         |
//    | 19    | h:mm:ss AM/PM                                      |
//    | 20    | h:mm                                               |
//    | 21    | h:mm:ss                                            |
//    | 22    | m/d/yy h:mm                                        |
//    | ...   | ...                                                |
//    | 37    | (#,##0_);(#,##0)                                   |
//    | 38    | (#,##0_);[Red](#,##0)                              |
//    | 39    | (#,##0.00_);(#,##0.00)                             |
//    | 40    | (#,##0.00_);[Red](#,##0.00)                        |
//    | 41    | _(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)            |
//    | 42    | _($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)         |
//    | 43    | _(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)    |
//    | 44    | _($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_) |
//    | 45    | mm:ss                                              |
//    | 46    | [h]:mm:ss                                          |
//    | 47    | mm:ss.0                                            |
//    | 48    | ##0.0E+0                                           |
//    | 49    | @                                                  |
//
// Excelize built-in currency formats are shown in the following table, only
// support these types in the following table (Index number is used only for
// markup and is not used inside an Excel file and you can't get formatted value
// by the function GetCellValue) currently:
//
//    | Index | Symbol                                             |
//    +-------+----------------------------------------------------+
//    | 164   | CN¥                                                |
//    | 165   | $ English (China)                                  |
//    | 166   | $ Cherokee (United States)                         |
//    | 167   | $ Chinese (Singapore)                              |
//    | 168   | $ Chinese (Taiwan)                                 |
//    | 169   | $ English (Australia)                              |
//    | 170   | $ English (Belize)                                 |
//    | 171   | $ English (Canada)                                 |
//    | 172   | $ English (Jamaica)                                |
//    | 173   | $ English (New Zealand)                            |
//    | 174   | $ English (Singapore)                              |
//    | 175   | $ English (Trinidad & Tobago)                      |
//    | 176   | $ English (U.S. Vigin Islands)                     |
//    | 177   | $ English (United States)                          |
//    | 178   | $ French (Canada)                                  |
//    | 179   | $ Hawaiian (United States)                         |
//    | 180   | $ Malay (Brunei)                                   |
//    | 181   | $ Quechua (Ecuador)                                |
//    | 182   | $ Spanish (Chile)                                  |
//    | 183   | $ Spanish (Colombia)                               |
//    | 184   | $ Spanish (Ecuador)                                |
//    | 185   | $ Spanish (El Salvador)                            |
//    | 186   | $ Spanish (Mexico)                                 |
//    | 187   | $ Spanish (Puerto Rico)                            |
//    | 188   | $ Spanish (United States)                          |
//    | 189   | $ Spanish (Uruguay)                                |
//
func (f *File) NewStyle(style string) (int, error) {
	var cellXfsID int
	styleSheet := f.stylesReader()
	formatCellStyle, err := parseFormatStyleSet(style)
	if err != nil {
		return cellXfsID, err
	}
	numFmtID := setNumFmt(styleSheet, formatCellStyle)
	fontID := setFont(styleSheet, formatCellStyle)
	borderID := setBorders(styleSheet, formatCellStyle)
	fillID := setFills(styleSheet, formatCellStyle)
	applyAlignment, alignment := setAlignment(styleSheet, formatCellStyle)
	cellXfsID = setCellXfs(styleSheet, fontID, numFmtID, fillID, borderID, applyAlignment, alignment)
	return cellXfsID, nil
}

// setFont provides function to add font style by given cell format settings.
func setFont(style *xlsxStyleSheet, formatCellStyle *formatCellStyle) int {
	if formatCellStyle.Font == nil {
		return 0
	}
	fontUnderlineType := map[string]string{"single": "single", "double": "double"}
	if formatCellStyle.Font.Family == "" {
		formatCellStyle.Font.Family = "Calibri"
	}
	if formatCellStyle.Font.Size < 1 {
		formatCellStyle.Font.Size = 11
	}
	if formatCellStyle.Font.Color == "" {
		formatCellStyle.Font.Color = "#000000"
	}
	f := font{
		B:      formatCellStyle.Font.Bold,
		I:      formatCellStyle.Font.Italic,
		Sz:     &attrValInt{Val: formatCellStyle.Font.Size},
		Color:  &xlsxColor{RGB: getPaletteColor(formatCellStyle.Font.Color)},
		Name:   &attrValString{Val: formatCellStyle.Font.Family},
		Family: &attrValInt{Val: 2},
		Scheme: &attrValString{Val: "minor"},
	}
	val, ok := fontUnderlineType[formatCellStyle.Font.Underline]
	if ok {
		f.U = &attrValString{Val: val}
	}
	font, _ := xml.Marshal(f)
	style.Fonts.Count++
	style.Fonts.Font = append(style.Fonts.Font, &xlsxFont{
		Font: string(font[6 : len(font)-7]),
	})
	return style.Fonts.Count - 1
}

// setNumFmt provides function to check if number format code in the range of
// built-in values.
func setNumFmt(style *xlsxStyleSheet, formatCellStyle *formatCellStyle) int {
	dp := "0."
	numFmtID := 164 // Default custom number format code from 164.
	if formatCellStyle.DecimalPlaces < 0 || formatCellStyle.DecimalPlaces > 30 {
		formatCellStyle.DecimalPlaces = 2
	}
	for i := 0; i < formatCellStyle.DecimalPlaces; i++ {
		dp += "0"
	}
	_, ok := builtInNumFmt[formatCellStyle.NumFmt]
	if !ok {
		fc, currency := currencyNumFmt[formatCellStyle.NumFmt]
		if !currency {
			return setLangNumFmt(style, formatCellStyle)
		}
		fc = strings.Replace(fc, "0.00", dp, -1)
		if style.NumFmts != nil {
			numFmtID = style.NumFmts.NumFmt[len(style.NumFmts.NumFmt)-1].NumFmtID + 1
			nf := xlsxNumFmt{
				FormatCode: fc,
				NumFmtID:   numFmtID,
			}
			style.NumFmts.NumFmt = append(style.NumFmts.NumFmt, &nf)
			style.NumFmts.Count++
		} else {
			nf := xlsxNumFmt{
				FormatCode: fc,
				NumFmtID:   numFmtID,
			}
			numFmts := xlsxNumFmts{
				NumFmt: []*xlsxNumFmt{&nf},
				Count:  1,
			}
			style.NumFmts = &numFmts
		}
		return numFmtID
	}
	return formatCellStyle.NumFmt
}

// setLangNumFmt provides function to set number format code with language.
func setLangNumFmt(style *xlsxStyleSheet, formatCellStyle *formatCellStyle) int {
	numFmts, ok := langNumFmt[formatCellStyle.Lang]
	if !ok {
		return 0
	}
	var fc string
	fc, ok = numFmts[formatCellStyle.NumFmt]
	if !ok {
		return 0
	}
	nf := xlsxNumFmt{
		FormatCode: fc,
		NumFmtID:   formatCellStyle.NumFmt,
	}
	if style.NumFmts != nil {
		style.NumFmts.NumFmt = append(style.NumFmts.NumFmt, &nf)
		style.NumFmts.Count++
	} else {
		numFmts := xlsxNumFmts{
			NumFmt: []*xlsxNumFmt{&nf},
			Count:  1,
		}
		style.NumFmts = &numFmts
	}
	return formatCellStyle.NumFmt
}

// setFills provides function to add fill elements in the styles.xml by given
// cell format settings.
func setFills(style *xlsxStyleSheet, formatCellStyle *formatCellStyle) int {
	var patterns = []string{
		"none",
		"solid",
		"mediumGray",
		"darkGray",
		"lightGray",
		"darkHorizontal",
		"darkVertical",
		"darkDown",
		"darkUp",
		"darkGrid",
		"darkTrellis",
		"lightHorizontal",
		"lightVertical",
		"lightDown",
		"lightUp",
		"lightGrid",
		"lightTrellis",
		"gray125",
		"gray0625",
	}

	var variants = []float64{
		90,
		0,
		45,
		135,
	}

	var fill xlsxFill
	switch formatCellStyle.Fill.Type {
	case "gradient":
		if len(formatCellStyle.Fill.Color) != 2 {
			break
		}
		var gradient xlsxGradientFill
		switch formatCellStyle.Fill.Shading {
		case 0, 1, 2, 3:
			gradient.Degree = variants[formatCellStyle.Fill.Shading]
		case 4:
			gradient.Type = "path"
		case 5:
			gradient.Type = "path"
			gradient.Bottom = 0.5
			gradient.Left = 0.5
			gradient.Right = 0.5
			gradient.Top = 0.5
		default:
			break
		}
		var stops []*xlsxGradientFillStop
		for index, color := range formatCellStyle.Fill.Color {
			var stop xlsxGradientFillStop
			stop.Position = float64(index)
			stop.Color.RGB = getPaletteColor(color)
			stops = append(stops, &stop)
		}
		gradient.Stop = stops
		fill.GradientFill = &gradient
	case "pattern":
		if formatCellStyle.Fill.Pattern > 18 || formatCellStyle.Fill.Pattern < 0 {
			break
		}
		if len(formatCellStyle.Fill.Color) < 1 {
			break
		}
		var pattern xlsxPatternFill
		pattern.PatternType = patterns[formatCellStyle.Fill.Pattern]
		pattern.FgColor.RGB = getPaletteColor(formatCellStyle.Fill.Color[0])
		fill.PatternFill = &pattern
	}
	style.Fills.Count++
	style.Fills.Fill = append(style.Fills.Fill, &fill)
	return style.Fills.Count - 1
}

// setAlignment provides function to formatting information pertaining to text
// alignment in cells. There are a variety of choices for how text is aligned
// both horizontally and vertically, as well as indentation settings, and so on.
func setAlignment(style *xlsxStyleSheet, formatCellStyle *formatCellStyle) (bool, *xlsxAlignment) {
	if formatCellStyle.Alignment == nil {
		return false, &xlsxAlignment{}
	}
	var alignment = xlsxAlignment{
		Horizontal:      formatCellStyle.Alignment.Horizontal,
		Indent:          formatCellStyle.Alignment.Indent,
		JustifyLastLine: formatCellStyle.Alignment.JustifyLastLine,
		ReadingOrder:    formatCellStyle.Alignment.ReadingOrder,
		RelativeIndent:  formatCellStyle.Alignment.RelativeIndent,
		ShrinkToFit:     formatCellStyle.Alignment.ShrinkToFit,
		TextRotation:    formatCellStyle.Alignment.TextRotation,
		Vertical:        formatCellStyle.Alignment.Vertical,
		WrapText:        formatCellStyle.Alignment.WrapText,
	}
	return true, &alignment
}

// setBorders provides function to add border elements in the styles.xml by
// given borders format settings.
func setBorders(style *xlsxStyleSheet, formatCellStyle *formatCellStyle) int {
	var styles = []string{
		"none",
		"thin",
		"medium",
		"dashed",
		"dotted",
		"thick",
		"double",
		"hair",
		"mediumDashed",
		"dashDot",
		"mediumDashDot",
		"dashDotDot",
		"mediumDashDotDot",
		"slantDashDot",
	}

	var border xlsxBorder
	for _, v := range formatCellStyle.Border {
		if v.Style > 13 || v.Style < 0 {
			continue
		}
		var color xlsxColor
		color.RGB = getPaletteColor(v.Color)
		switch v.Type {
		case "left":
			border.Left.Style = styles[v.Style]
			border.Left.Color = &color
		case "right":
			border.Right.Style = styles[v.Style]
			border.Right.Color = &color
		case "top":
			border.Top.Style = styles[v.Style]
			border.Top.Color = &color
		case "bottom":
			border.Bottom.Style = styles[v.Style]
			border.Bottom.Color = &color
		case "diagonalUp":
			border.Diagonal.Style = styles[v.Style]
			border.Diagonal.Color = &color
			border.DiagonalUp = true
		case "diagonalDown":
			border.Diagonal.Style = styles[v.Style]
			border.Diagonal.Color = &color
			border.DiagonalDown = true
		}
	}
	style.Borders.Count++
	style.Borders.Border = append(style.Borders.Border, &border)
	return style.Borders.Count - 1
}

// setCellXfs provides function to set describes all of the formatting for a
// cell.
func setCellXfs(style *xlsxStyleSheet, fontID, numFmtID, fillID, borderID int, applyAlignment bool, alignment *xlsxAlignment) int {
	var xf xlsxXf
	xf.FontID = fontID
	if fontID != 0 {
		xf.ApplyFont = true
	}
	xf.NumFmtID = numFmtID
	if numFmtID != 0 {
		xf.ApplyNumberFormat = true
	}
	xf.FillID = fillID
	xf.BorderID = borderID
	style.CellXfs.Count++
	xf.Alignment = alignment
	xf.ApplyAlignment = applyAlignment
	style.CellXfs.Xf = append(style.CellXfs.Xf, xf)
	return style.CellXfs.Count - 1
}

// SetCellStyle provides function to add style attribute for cells by given
// worksheet sheet index, coordinate area and style ID. Note that diagonalDown
// and diagonalUp type border should be use same color in the same coordinate
// area.
//
// For example create a borders of cell H9 on Sheet1:
//
//    style, err := xlsx.NewStyle(`{"border":[{"type":"left","color":"0000FF","style":3},{"type":"top","color":"00FF00","style":4},{"type":"bottom","color":"FFFF00","style":5},{"type":"right","color":"FF0000","style":6},{"type":"diagonalDown","color":"A020F0","style":7},{"type":"diagonalUp","color":"A020F0","style":8}]}`)
//    if err != nil {
//        fmt.Println(err)
//    }
//    xlsx.SetCellStyle("Sheet1", "H9", "H9", style)
//
// Set gradient fill with vertical variants shading styles for cell H9 on
// Sheet1:
//
//    style, err := xlsx.NewStyle(`{"fill":{"type":"gradient","color":["#FFFFFF","#E0EBF5"],"shading":1}}`)
//    if err != nil {
//        fmt.Println(err)
//    }
//    xlsx.SetCellStyle("Sheet1", "H9", "H9", style)
//
// Set solid style pattern fill for cell H9 on Sheet1:
//
//    style, err := xlsx.NewStyle(`{"fill":{"type":"pattern","color":["#E0EBF5"],"pattern":1}}`)
//    if err != nil {
//        fmt.Println(err)
//    }
//    xlsx.SetCellStyle("Sheet1", "H9", "H9", style)
//
// Set alignment style for cell H9 on Sheet1:
//
//    style, err := xlsx.NewStyle(`{"alignment":{"horizontal":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":45,"vertical":"","wrap_text":true}}`)
//    if err != nil {
//        fmt.Println(err)
//    }
//    xlsx.SetCellStyle("Sheet1", "H9", "H9", style)
//
// Dates and times in Excel are represented by real numbers, for example "Apr 7
// 2017 12:00 PM" is represented by the number 42920.5. Set date and time format
// for cell H9 on Sheet1:
//
//    xlsx.SetCellValue("Sheet1", "H9", 42920.5)
//    style, err := xlsx.NewStyle(`{"number_format": 22}`)
//    if err != nil {
//        fmt.Println(err)
//    }
//    xlsx.SetCellStyle("Sheet1", "H9", "H9", style)
//
// Set font style for cell H9 on Sheet1:
//
//    style, err := xlsx.NewStyle(`{"font":{"bold":true,"italic":true,"family":"Berlin Sans FB Demi","size":36,"color":"#777777"}}`)
//    if err != nil {
//        fmt.Println(err)
//    }
//    xlsx.SetCellStyle("Sheet1", "H9", "H9", style)
//
func (f *File) SetCellStyle(sheet, hcell, vcell string, styleID int) {
	hcell = strings.ToUpper(hcell)
	vcell = strings.ToUpper(vcell)

	// Coordinate conversion, convert C1:B3 to 2,0,1,2.
	hcol := string(strings.Map(letterOnlyMapF, hcell))
	hrow, _ := strconv.Atoi(strings.Map(intOnlyMapF, hcell))
	hyAxis := hrow - 1
	hxAxis := TitleToNumber(hcol)

	vcol := string(strings.Map(letterOnlyMapF, vcell))
	vrow, _ := strconv.Atoi(strings.Map(intOnlyMapF, vcell))
	vyAxis := vrow - 1
	vxAxis := TitleToNumber(vcol)

	if vxAxis < hxAxis {
		hcell, vcell = vcell, hcell
		vxAxis, hxAxis = hxAxis, vxAxis
	}

	if vyAxis < hyAxis {
		hcell, vcell = vcell, hcell
		vyAxis, hyAxis = hyAxis, vyAxis
	}

	// Correct the coordinate area, such correct C1:B3 to B1:C3.
	hcell = ToAlphaString(hxAxis) + strconv.Itoa(hyAxis+1)
	vcell = ToAlphaString(vxAxis) + strconv.Itoa(vyAxis+1)

	xlsx := f.workSheetReader(sheet)

	completeRow(xlsx, vyAxis+1, vxAxis+1)
	completeCol(xlsx, vyAxis+1, vxAxis+1)

	for r, row := range xlsx.SheetData.Row {
		for k, c := range row.C {
			if checkCellInArea(c.R, hcell+":"+vcell) {
				xlsx.SheetData.Row[r].C[k].S = styleID
			}
		}
	}
}

// getPaletteColor provides function to convert the RBG color by given string.
func getPaletteColor(color string) string {
	return "FF" + strings.Replace(strings.ToUpper(color), "#", "", -1)
}
