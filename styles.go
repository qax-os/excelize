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

// parseFormatStyleSet provides function to parse the format settings of the
// borders.
func parseFormatStyleSet(style string) (*formatCellStyle, error) {
	var format formatCellStyle
	err := json.Unmarshal([]byte(style), &format)
	return &format, err
}

// SetCellStyle provides function to set style for cells by given sheet index
// and coordinate area in XLSX file. Note that the color field uses RGB color
// code and diagonalDown and diagonalUp type border should be use same color in
// the same coordinate area.
//
// For example create a borders of cell H9 on Sheet1:
//
//    err := xlsx.SetCellStyle("Sheet1", "H9", "H9", `{"border":[{"type":"left","color":"0000FF","style":3},{"type":"top","color":"00FF00","style":4},{"type":"bottom","color":"FFFF00","style":5},{"type":"right","color":"FF0000","style":6},{"type":"diagonalDown","color":"A020F0","style":7},{"type":"diagonalUp","color":"A020F0","style":8}]}`)
//    if err != nil {
//        fmt.Println(err)
//    }
//
// Set gradient fill with vertical variants shading styles for cell H9 on
// Sheet1:
//
//    err := xlsx.SetCellStyle("Sheet1", "H9", "H9", `{"fill":{"type":"gradient","color":["#FFFFFF","#E0EBF5"],"shading":1}}`)
//    if err != nil {
//        fmt.Println(err)
//    }
//
// Set solid style pattern fill for cell H9 on Sheet1:
//
//    err := xlsx.SetCellStyle("Sheet1", "H9", "H9", `{"fill":{"type":"pattern","color":["#E0EBF5"],"pattern":1}}`)
//    if err != nil {
//        fmt.Println(err)
//    }
//
// Set alignment style for cell H9 on Sheet1:
//
//    err = xlsx.SetCellStyle("Sheet1", "H9", "H9", `{"alignment":{"horizontal":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":45,"vertical":"","wrap_text":true}}`)
//    if err != nil {
//        fmt.Println(err)
//    }
//
// Dates and times in Excel are represented by real numbers, for example "Apr 7
// 2017 12:00 PM" is represented by the number 42920.5. Set date and time format
// for cell H9 on Sheet1:
//
//    xlsx.SetCellValue("Sheet2", "H9", 42920.5)
//    err = xlsx.SetCellStyle("Sheet1", "H9", "H9", `{"number_format": 22}`)
//    if err != nil {
//        fmt.Println(err)
//    }
//
// Set font style for cell H9 on Sheet1:
//
//    err = xlsx.SetCellStyle("Sheet1", "H9", "H9", `{"font":{"bold":true,"italic":true,"family":"Berlin Sans FB Demi","size":36,"color":"#777777"}}`)
//    if err != nil {
//        fmt.Println(err)
//    }
//
// The following shows the border styles sorted by excelize index number:
//
//    +-------+---------------+--------+-----------------+
//    | Index | Name          | Weight | Style           |
//    +=======+===============+========+=================+
//    | 0     | None          | 0      |                 |
//    +-------+---------------+--------+-----------------+
//    | 1     | Continuous    | 1      | ``-----------`` |
//    +-------+---------------+--------+-----------------+
//    | 2     | Continuous    | 2      | ``-----------`` |
//    +-------+---------------+--------+-----------------+
//    | 3     | Dash          | 1      | ``- - - - - -`` |
//    +-------+---------------+--------+-----------------+
//    | 4     | Dot           | 1      | ``. . . . . .`` |
//    +-------+---------------+--------+-----------------+
//    | 5     | Continuous    | 3      | ``-----------`` |
//    +-------+---------------+--------+-----------------+
//    | 6     | Double        | 3      | ``===========`` |
//    +-------+---------------+--------+-----------------+
//    | 7     | Continuous    | 0      | ``-----------`` |
//    +-------+---------------+--------+-----------------+
//    | 8     | Dash          | 2      | ``- - - - - -`` |
//    +-------+---------------+--------+-----------------+
//    | 9     | Dash Dot      | 1      | ``- . - . - .`` |
//    +-------+---------------+--------+-----------------+
//    | 10    | Dash Dot      | 2      | ``- . - . - .`` |
//    +-------+---------------+--------+-----------------+
//    | 11    | Dash Dot Dot  | 1      | ``- . . - . .`` |
//    +-------+---------------+--------+-----------------+
//    | 12    | Dash Dot Dot  | 2      | ``- . . - . .`` |
//    +-------+---------------+--------+-----------------+
//    | 13    | SlantDash Dot | 2      | ``/ - . / - .`` |
//    +-------+---------------+--------+-----------------+
//
// The following shows the borders in the order shown in the Excel dialog:
//
//    +-------+-----------------+-------+-----------------+
//    | Index | Style           | Index | Style           |
//    +=======+=================+=======+=================+
//    | 0     | None            | 12    | ``- . . - . .`` |
//    +-------+-----------------+-------+-----------------+
//    | 7     | ``-----------`` | 13    | ``/ - . / - .`` |
//    +-------+-----------------+-------+-----------------+
//    | 4     | ``. . . . . .`` | 10    | ``- . - . - .`` |
//    +-------+-----------------+-------+-----------------+
//    | 11    | ``- . . - . .`` | 8     | ``- - - - - -`` |
//    +-------+-----------------+-------+-----------------+
//    | 9     | ``- . - . - .`` | 2     | ``-----------`` |
//    +-------+-----------------+-------+-----------------+
//    | 3     | ``- - - - - -`` | 5     | ``-----------`` |
//    +-------+-----------------+-------+-----------------+
//    | 1     | ``-----------`` | 6     | ``===========`` |
//    +-------+-----------------+-------+-----------------+
//
// The following shows the shading styles sorted by excelize index number:
//
//    +-------+-----------------+-------+-----------------+
//    | Index | Style           | Index | Style           |
//    +=======+=================+=======+=================+
//    | 0     | Horizontal      | 3     | Diagonal down   |
//    +-------+-----------------+-------+-----------------+
//    | 1     | Vertical        | 4     | From corner     |
//    +-------+-----------------+-------+-----------------+
//    | 2     | Diagonal Up     | 5     | From center     |
//    +-------+-----------------+-------+-----------------+
//
// The following shows the patterns styles sorted by excelize index number:
//
//    +-------+-----------------+-------+-----------------+
//    | Index | Style           | Index | Style           |
//    +=======+=================+=======+=================+
//    | 0     | None            | 10    | darkTrellis     |
//    +-------+-----------------+-------+-----------------+
//    | 1     | solid           | 11    | lightHorizontal |
//    +-------+-----------------+-------+-----------------+
//    | 2     | mediumGray      | 12    | lightVertical   |
//    +-------+-----------------+-------+-----------------+
//    | 3     | darkGray        | 13    | lightDown       |
//    +-------+-----------------+-------+-----------------+
//    | 4     | lightGray       | 14    | lightUp         |
//    +-------+-----------------+-------+-----------------+
//    | 5     | darkHorizontal  | 15    | lightGrid       |
//    +-------+-----------------+-------+-----------------+
//    | 6     | darkVertical    | 16    | lightTrellis    |
//    +-------+-----------------+-------+-----------------+
//    | 7     | darkDown        | 17    | gray125         |
//    +-------+-----------------+-------+-----------------+
//    | 8     | darkUp          | 18    | gray0625        |
//    +-------+-----------------+-------+-----------------+
//    | 9     | darkGrid        |       |                 |
//    +-------+-----------------+-------+-----------------+
//
// The following the type of horizontal alignment in cells:
//
//    +------------------+
//    | Style            |
//    +==================+
//    | left             |
//    +------------------+
//    | center           |
//    +------------------+
//    | right            |
//    +------------------+
//    | fill             |
//    +------------------+
//    | justify          |
//    +------------------+
//    | centerContinuous |
//    +------------------+
//    | distributed      |
//    +------------------+
//
// The following the type of vertical alignment in cells:
//
//    +------------------+
//    | Style            |
//    +==================+
//    | top              |
//    +------------------+
//    | center           |
//    +------------------+
//    | justify          |
//    +------------------+
//    | distributed      |
//    +------------------+
//
// The following the type of font underline style:
//
//    +------------------+
//    | Style            |
//    +==================+
//    | single           |
//    +------------------+
//    | double           |
//    +------------------+
//
// Excel's built-in formats are shown in the following table:
//
//    +-------+----------------------------------------------------+
//    | Index | Format String                                      |
//    +=======+====================================================+
//    | 0     | General                                            |
//    +-------+----------------------------------------------------+
//    | 1     | 0                                                  |
//    +-------+----------------------------------------------------+
//    | 2     | 0.00                                               |
//    +-------+----------------------------------------------------+
//    | 3     | #,##0                                              |
//    +-------+----------------------------------------------------+
//    | 4     | #,##0.00                                           |
//    +-------+----------------------------------------------------+
//    | 5     | ($#,##0_);($#,##0)                                 |
//    +-------+----------------------------------------------------+
//    | 6     | ($#,##0_);[Red]($#,##0)                            |
//    +-------+----------------------------------------------------+
//    | 7     | ($#,##0.00_);($#,##0.00)                           |
//    +-------+----------------------------------------------------+
//    | 8     | ($#,##0.00_);[Red]($#,##0.00)                      |
//    +-------+----------------------------------------------------+
//    | 9     | 0%                                                 |
//    +-------+----------------------------------------------------+
//    | 10    | 0.00%                                              |
//    +-------+----------------------------------------------------+
//    | 11    | 0.00E+00                                           |
//    +-------+----------------------------------------------------+
//    | 12    | # ?/?                                              |
//    +-------+----------------------------------------------------+
//    | 13    | # ??/??                                            |
//    +-------+----------------------------------------------------+
//    | 14    | m/d/yy                                             |
//    +-------+----------------------------------------------------+
//    | 15    | d-mmm-yy                                           |
//    +-------+----------------------------------------------------+
//    | 16    | d-mmm                                              |
//    +-------+----------------------------------------------------+
//    | 17    | mmm-yy                                             |
//    +-------+----------------------------------------------------+
//    | 18    | h:mm AM/PM                                         |
//    +-------+----------------------------------------------------+
//    | 19    | h:mm:ss AM/PM                                      |
//    +-------+----------------------------------------------------+
//    | 20    | h:mm                                               |
//    +-------+----------------------------------------------------+
//    | 21    | h:mm:ss                                            |
//    +-------+----------------------------------------------------+
//    | 22    | m/d/yy h:mm                                        |
//    +-------+----------------------------------------------------+
//    | ...   | ...                                                |
//    +-------+----------------------------------------------------+
//    | 37    | (#,##0_);(#,##0)                                   |
//    +-------+----------------------------------------------------+
//    | 38    | (#,##0_);[Red](#,##0)                              |
//    +-------+----------------------------------------------------+
//    | 39    | (#,##0.00_);(#,##0.00)                             |
//    +-------+----------------------------------------------------+
//    | 40    | (#,##0.00_);[Red](#,##0.00)                        |
//    +-------+----------------------------------------------------+
//    | 41    | _(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)            |
//    +-------+----------------------------------------------------+
//    | 42    | _($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)         |
//    +-------+----------------------------------------------------+
//    | 43    | _(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)    |
//    +-------+----------------------------------------------------+
//    | 44    | _($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_) |
//    +-------+----------------------------------------------------+
//    | 45    | mm:ss                                              |
//    +-------+----------------------------------------------------+
//    | 46    | [h]:mm:ss                                          |
//    +-------+----------------------------------------------------+
//    | 47    | mm:ss.0                                            |
//    +-------+----------------------------------------------------+
//    | 48    | ##0.0E+0                                           |
//    +-------+----------------------------------------------------+
//    | 49    | @                                                  |
//    +-------+----------------------------------------------------+
//
func (f *File) SetCellStyle(sheet, hcell, vcell, style string) error {
	var styleSheet xlsxStyleSheet
	xml.Unmarshal([]byte(f.readXML("xl/styles.xml")), &styleSheet)
	formatCellStyle, err := parseFormatStyleSet(style)
	if err != nil {
		return err
	}
	numFmtID := setNumFmt(&styleSheet, formatCellStyle)
	fontID := setFont(&styleSheet, formatCellStyle)
	borderID := setBorders(&styleSheet, formatCellStyle)
	fillID := setFills(&styleSheet, formatCellStyle)
	applyAlignment, alignment := setAlignment(&styleSheet, formatCellStyle)
	cellXfsID := setCellXfs(&styleSheet, fontID, numFmtID, fillID, borderID, applyAlignment, alignment)
	output, err := xml.Marshal(styleSheet)
	if err != nil {
		return err
	}
	f.saveFileList("xl/styles.xml", replaceWorkSheetsRelationshipsNameSpace(string(output)))
	f.setCellStyle(sheet, hcell, vcell, cellXfsID)
	return err
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
	_, ok := builtInNumFmt[formatCellStyle.NumFmt]
	if !ok {
		return 0
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

// setCellStyle provides function to add style attribute for cells by given
// sheet index, coordinate area and style ID.
func (f *File) setCellStyle(sheet, hcell, vcell string, styleID int) {
	hcell = strings.ToUpper(hcell)
	vcell = strings.ToUpper(vcell)

	// Coordinate conversion, convert C1:B3 to 2,0,1,2.
	hcol := string(strings.Map(letterOnlyMapF, hcell))
	hrow, _ := strconv.Atoi(strings.Map(intOnlyMapF, hcell))
	hyAxis := hrow - 1
	hxAxis := titleToNumber(hcol)

	vcol := string(strings.Map(letterOnlyMapF, vcell))
	vrow, _ := strconv.Atoi(strings.Map(intOnlyMapF, vcell))
	vyAxis := vrow - 1
	vxAxis := titleToNumber(vcol)

	if vxAxis < hxAxis {
		hcell, vcell = vcell, hcell
		vxAxis, hxAxis = hxAxis, vxAxis
	}

	if vyAxis < hyAxis {
		hcell, vcell = vcell, hcell
		vyAxis, hyAxis = hyAxis, vyAxis
	}

	// Correct the coordinate area, such correct C1:B3 to B1:C3.
	hcell = toAlphaString(hxAxis+1) + strconv.Itoa(hyAxis+1)
	vcell = toAlphaString(vxAxis+1) + strconv.Itoa(vyAxis+1)

	xlsx := f.workSheetReader(sheet)

	completeRow(xlsx, vxAxis+1, vyAxis+1)
	completeCol(xlsx, vxAxis+1, vyAxis+1)

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
