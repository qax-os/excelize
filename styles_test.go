package excelize

import (
	"fmt"
	"math"
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestStyleFill(t *testing.T) {
	cases := []struct {
		label      string
		format     string
		expectFill bool
	}{{
		label:      "no_fill",
		format:     `{"alignment":{"wrap_text":true}}`,
		expectFill: false,
	}, {
		label:      "fill",
		format:     `{"fill":{"type":"pattern","pattern":1,"color":["#000000"]}}`,
		expectFill: true,
	}}

	for _, testCase := range cases {
		xl := NewFile()
		styleID, err := xl.NewStyle(testCase.format)
		assert.NoError(t, err)

		styles := xl.stylesReader()
		style := styles.CellXfs.Xf[styleID]
		if testCase.expectFill {
			assert.NotEqual(t, *style.FillID, 0, testCase.label)
		} else {
			assert.Equal(t, *style.FillID, 0, testCase.label)
		}
	}
	f := NewFile()
	styleID1, err := f.NewStyle(`{"fill":{"type":"pattern","pattern":1,"color":["#000000"]}}`)
	assert.NoError(t, err)
	styleID2, err := f.NewStyle(`{"fill":{"type":"pattern","pattern":1,"color":["#000000"]}}`)
	assert.NoError(t, err)
	assert.Equal(t, styleID1, styleID2)
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestStyleFill.xlsx")))
}

func TestSetConditionalFormat(t *testing.T) {
	cases := []struct {
		label  string
		format string
		rules  []*xlsxCfRule
	}{{
		label: "3_color_scale",
		format: `[{
			"type":"3_color_scale",
			"criteria":"=",
			"min_type":"num",
			"mid_type":"num",
			"max_type":"num",
			"min_value": "-10",
			"mid_value": "0",
			"max_value": "10",
			"min_color":"ff0000",
			"mid_color":"00ff00",
			"max_color":"0000ff"
		}]`,
		rules: []*xlsxCfRule{{
			Priority: 1,
			Type:     "colorScale",
			ColorScale: &xlsxColorScale{
				Cfvo: []*xlsxCfvo{{
					Type: "num",
					Val:  "-10",
				}, {
					Type: "num",
					Val:  "0",
				}, {
					Type: "num",
					Val:  "10",
				}},
				Color: []*xlsxColor{{
					RGB: "FFFF0000",
				}, {
					RGB: "FF00FF00",
				}, {
					RGB: "FF0000FF",
				}},
			},
		}},
	}, {
		label: "3_color_scale default min/mid/max",
		format: `[{
			"type":"3_color_scale",
			"criteria":"=",
			"min_type":"num",
			"mid_type":"num",
			"max_type":"num",
			"min_color":"ff0000",
			"mid_color":"00ff00",
			"max_color":"0000ff"
		}]`,
		rules: []*xlsxCfRule{{
			Priority: 1,
			Type:     "colorScale",
			ColorScale: &xlsxColorScale{
				Cfvo: []*xlsxCfvo{{
					Type: "num",
					Val:  "0",
				}, {
					Type: "num",
					Val:  "50",
				}, {
					Type: "num",
					Val:  "0",
				}},
				Color: []*xlsxColor{{
					RGB: "FFFF0000",
				}, {
					RGB: "FF00FF00",
				}, {
					RGB: "FF0000FF",
				}},
			},
		}},
	}, {
		label: "2_color_scale default min/max",
		format: `[{
			"type":"2_color_scale",
			"criteria":"=",
			"min_type":"num",
			"max_type":"num",
			"min_color":"ff0000",
			"max_color":"0000ff"
		}]`,
		rules: []*xlsxCfRule{{
			Priority: 1,
			Type:     "colorScale",
			ColorScale: &xlsxColorScale{
				Cfvo: []*xlsxCfvo{{
					Type: "num",
					Val:  "0",
				}, {
					Type: "num",
					Val:  "0",
				}},
				Color: []*xlsxColor{{
					RGB: "FFFF0000",
				}, {
					RGB: "FF0000FF",
				}},
			},
		}},
	}}

	for _, testCase := range cases {
		f := NewFile()
		const sheet = "Sheet1"
		const cellRange = "A1:A1"

		err := f.SetConditionalFormat(sheet, cellRange, testCase.format)
		if err != nil {
			t.Fatalf("%s", err)
		}

		ws, err := f.workSheetReader(sheet)
		assert.NoError(t, err)
		cf := ws.ConditionalFormatting
		assert.Len(t, cf, 1, testCase.label)
		assert.Len(t, cf[0].CfRule, 1, testCase.label)
		assert.Equal(t, cellRange, cf[0].SQRef, testCase.label)
		assert.EqualValues(t, testCase.rules, cf[0].CfRule, testCase.label)
	}
}

func TestUnsetConditionalFormat(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 7))
	assert.NoError(t, f.UnsetConditionalFormat("Sheet1", "A1:A10"))
	format, err := f.NewConditionalStyle(`{"font":{"color":"#9A0511"},"fill":{"type":"pattern","color":["#FEC7CE"],"pattern":1}}`)
	assert.NoError(t, err)
	assert.NoError(t, f.SetConditionalFormat("Sheet1", "A1:A10", fmt.Sprintf(`[{"type":"cell","criteria":">","format":%d,"value":"6"}]`, format)))
	assert.NoError(t, f.UnsetConditionalFormat("Sheet1", "A1:A10"))
	// Test unset conditional format on not exists worksheet.
	assert.EqualError(t, f.UnsetConditionalFormat("SheetN", "A1:A10"), "sheet SheetN is not exist")
	// Save spreadsheet by the given path.
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestUnsetConditionalFormat.xlsx")))
}

func TestNewStyle(t *testing.T) {
	f := NewFile()
	styleID, err := f.NewStyle(`{"font":{"bold":true,"italic":true,"family":"Times New Roman","size":36,"color":"#777777"}}`)
	assert.NoError(t, err)
	styles := f.stylesReader()
	fontID := styles.CellXfs.Xf[styleID].FontID
	font := styles.Fonts.Font[*fontID]
	assert.Contains(t, *font.Name.Val, "Times New Roman", "Stored font should contain font name")
	assert.Equal(t, 2, styles.CellXfs.Count, "Should have 2 styles")
	_, err = f.NewStyle(&Style{})
	assert.NoError(t, err)
	_, err = f.NewStyle(Style{})
	assert.EqualError(t, err, ErrParameterInvalid.Error())

	var exp string
	_, err = f.NewStyle(&Style{CustomNumFmt: &exp})
	assert.EqualError(t, err, ErrCustomNumFmt.Error())
	_, err = f.NewStyle(&Style{Font: &Font{Family: strings.Repeat("s", MaxFontFamilyLength+1)}})
	assert.EqualError(t, err, ErrFontLength.Error())
	_, err = f.NewStyle(&Style{Font: &Font{Size: MaxFontSize + 1}})
	assert.EqualError(t, err, ErrFontSize.Error())

	// new numeric custom style
	fmt := "####;####"
	f.Styles.NumFmts = nil
	styleID, err = f.NewStyle(&Style{
		CustomNumFmt: &fmt,
	})
	assert.NoError(t, err)
	assert.Equal(t, 2, styleID)

	assert.NotNil(t, f.Styles)
	assert.NotNil(t, f.Styles.CellXfs)
	assert.NotNil(t, f.Styles.CellXfs.Xf)

	nf := f.Styles.CellXfs.Xf[styleID]
	assert.Equal(t, 164, *nf.NumFmtID)

	// new currency custom style
	f.Styles.NumFmts = nil
	styleID, err = f.NewStyle(&Style{
		Lang:   "ko-kr",
		NumFmt: 32, // must not be in currencyNumFmt

	})
	assert.NoError(t, err)
	assert.Equal(t, 3, styleID)

	assert.NotNil(t, f.Styles)
	assert.NotNil(t, f.Styles.CellXfs)
	assert.NotNil(t, f.Styles.CellXfs.Xf)

	nf = f.Styles.CellXfs.Xf[styleID]
	assert.Equal(t, 32, *nf.NumFmtID)

	// Test set build-in scientific number format
	styleID, err = f.NewStyle(&Style{NumFmt: 11})
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "A1", "B1", styleID))
	assert.NoError(t, f.SetSheetRow("Sheet1", "A1", &[]float64{1.23, 1.234}))
	rows, err := f.GetRows("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, [][]string{{"1.23E+00", "1.23E+00"}}, rows)

	f = NewFile()
	// Test currency number format
	customNumFmt := "[$$-409]#,##0.00"
	style1, err := f.NewStyle(&Style{CustomNumFmt: &customNumFmt})
	assert.NoError(t, err)
	style2, err := f.NewStyle(&Style{NumFmt: 165})
	assert.NoError(t, err)
	assert.Equal(t, style1, style2)

	style3, err := f.NewStyle(&Style{NumFmt: 166})
	assert.NoError(t, err)
	assert.Equal(t, 2, style3)

	f = NewFile()
	f.Styles.NumFmts = nil
	f.Styles.CellXfs.Xf = nil
	style4, err := f.NewStyle(&Style{NumFmt: 160, Lang: "unknown"})
	assert.NoError(t, err)
	assert.Equal(t, 1, style4)

	f = NewFile()
	f.Styles.NumFmts = nil
	f.Styles.CellXfs.Xf = nil
	style5, err := f.NewStyle(&Style{NumFmt: 160, Lang: "zh-cn"})
	assert.NoError(t, err)
	assert.Equal(t, 1, style5)
}

func TestGetDefaultFont(t *testing.T) {
	f := NewFile()
	s := f.GetDefaultFont()
	assert.Equal(t, s, "Calibri", "Default font should be Calibri")
}

func TestSetDefaultFont(t *testing.T) {
	f := NewFile()
	f.SetDefaultFont("Arial")
	styles := f.stylesReader()
	s := f.GetDefaultFont()
	assert.Equal(t, s, "Arial", "Default font should change to Arial")
	assert.Equal(t, *styles.CellStyles.CellStyle[0].CustomBuiltIn, true)
}

func TestStylesReader(t *testing.T) {
	f := NewFile()
	// Test read styles with unsupported charset.
	f.Styles = nil
	f.Pkg.Store(defaultXMLPathStyles, MacintoshCyrillicCharset)
	assert.EqualValues(t, new(xlsxStyleSheet), f.stylesReader())
}

func TestThemeReader(t *testing.T) {
	f := NewFile()
	// Test read theme with unsupported charset.
	f.Pkg.Store("xl/theme/theme1.xml", MacintoshCyrillicCharset)
	assert.EqualValues(t, new(xlsxTheme), f.themeReader())
}

func TestSetCellStyle(t *testing.T) {
	f := NewFile()
	// Test set cell style on not exists worksheet.
	assert.EqualError(t, f.SetCellStyle("SheetN", "A1", "A2", 1), "sheet SheetN is not exist")
}

func TestGetStyleID(t *testing.T) {
	assert.Equal(t, -1, NewFile().getStyleID(&xlsxStyleSheet{}, nil))
}

func TestGetFillID(t *testing.T) {
	assert.Equal(t, -1, getFillID(NewFile().stylesReader(), &Style{Fill: Fill{Type: "unknown"}}))
}

func TestParseTime(t *testing.T) {
	assert.Equal(t, "2019", parseTime("43528", "YYYY"))
	assert.Equal(t, "43528", parseTime("43528", ""))

	assert.Equal(t, "2019-03-04 05:05:42", parseTime("43528.2123", "YYYY-MM-DD hh:mm:ss"))
	assert.Equal(t, "2019-03-04 05:05:42", parseTime("43528.2123", "YYYY-MM-DD hh:mm:ss;YYYY-MM-DD hh:mm:ss"))
	assert.Equal(t, "3/4/2019 5:5:42", parseTime("43528.2123", "M/D/YYYY h:m:s"))
	assert.Equal(t, "3/4/2019 0:5:42", parseTime("43528.003958333335", "m/d/yyyy h:m:s"))
	assert.Equal(t, "3/4/2019 0:05:42", parseTime("43528.003958333335", "M/D/YYYY h:mm:s"))
	assert.Equal(t, "3:30:00 PM", parseTime("0.64583333333333337", "h:mm:ss am/pm"))
	assert.Equal(t, "0:05", parseTime("43528.003958333335", "h:mm"))
	assert.Equal(t, "0:0", parseTime("6.9444444444444444E-5", "h:m"))
	assert.Equal(t, "0:00", parseTime("6.9444444444444444E-5", "h:mm"))
	assert.Equal(t, "0:0", parseTime("6.9444444444444444E-5", "h:m"))
	assert.Equal(t, "12:1", parseTime("0.50070601851851848", "h:m"))
	assert.Equal(t, "23:30", parseTime("0.97952546296296295", "h:m"))
	assert.Equal(t, "March", parseTime("43528", "mmmm"))
	assert.Equal(t, "Monday", parseTime("43528", "dddd"))
}

func TestThemeColor(t *testing.T) {
	for _, clr := range [][]string{
		{"FF000000", ThemeColor("000000", -0.1)},
		{"FF000000", ThemeColor("000000", 0)},
		{"FF33FF33", ThemeColor("00FF00", 0.2)},
		{"FFFFFFFF", ThemeColor("000000", 1)},
		{"FFFFFFFF", ThemeColor(strings.Repeat(string(rune(math.MaxUint8+1)), 6), 1)},
		{"FFFFFFFF", ThemeColor(strings.Repeat(string(rune(-1)), 6), 1)},
	} {
		assert.Equal(t, clr[0], clr[1])
	}
}

func TestGetNumFmtID(t *testing.T) {
	f := NewFile()

	fs1, err := parseFormatStyleSet(`{"protection":{"hidden":false,"locked":false},"number_format":10}`)
	assert.NoError(t, err)
	id1 := getNumFmtID(&xlsxStyleSheet{}, fs1)

	fs2, err := parseFormatStyleSet(`{"protection":{"hidden":false,"locked":false},"number_format":0}`)
	assert.NoError(t, err)
	id2 := getNumFmtID(&xlsxStyleSheet{}, fs2)

	assert.NotEqual(t, id1, id2)
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestStyleNumFmt.xlsx")))
}
