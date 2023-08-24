package excelize

import (
	"math"
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestStyleFill(t *testing.T) {
	cases := []struct {
		label      string
		format     *Style
		expectFill bool
	}{{
		label:      "no_fill",
		format:     &Style{Alignment: &Alignment{WrapText: true}},
		expectFill: false,
	}, {
		label:      "fill",
		format:     &Style{Fill: Fill{Type: "pattern", Pattern: 1, Color: []string{"000000"}}},
		expectFill: true,
	}}

	for _, testCase := range cases {
		xl := NewFile()
		styleID, err := xl.NewStyle(testCase.format)
		assert.NoError(t, err)

		styles, err := xl.stylesReader()
		assert.NoError(t, err)
		style := styles.CellXfs.Xf[styleID]
		if testCase.expectFill {
			assert.NotEqual(t, *style.FillID, 0, testCase.label)
		} else {
			assert.Equal(t, *style.FillID, 0, testCase.label)
		}
	}
	f := NewFile()
	styleID1, err := f.NewStyle(&Style{Fill: Fill{Type: "pattern", Pattern: 1, Color: []string{"000000"}}})
	assert.NoError(t, err)
	styleID2, err := f.NewStyle(&Style{Fill: Fill{Type: "pattern", Pattern: 1, Color: []string{"000000"}}})
	assert.NoError(t, err)
	assert.Equal(t, styleID1, styleID2)
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestStyleFill.xlsx")))
}

func TestSetConditionalFormat(t *testing.T) {
	cases := []struct {
		label  string
		format []ConditionalFormatOptions
		rules  []*xlsxCfRule
	}{{
		label: "3_color_scale",
		format: []ConditionalFormatOptions{{
			Type:     "3_color_scale",
			Criteria: "=",
			MinType:  "num",
			MidType:  "num",
			MaxType:  "num",
			MinValue: "-10",
			MidValue: "0",
			MaxValue: "10",
			MinColor: "ff0000",
			MidColor: "00ff00",
			MaxColor: "0000ff",
		}},
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
		format: []ConditionalFormatOptions{{
			Type:     "3_color_scale",
			Criteria: "=",
			MinType:  "num",
			MidType:  "num",
			MaxType:  "num",
			MinColor: "ff0000",
			MidColor: "00ff00",
			MaxColor: "0000ff",
		}},
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
		format: []ConditionalFormatOptions{{
			Type:     "2_color_scale",
			Criteria: "=",
			MinType:  "num",
			MaxType:  "num",
			MinColor: "ff0000",
			MaxColor: "0000ff",
		}},
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
		const rangeRef = "A1:A1"
		assert.NoError(t, f.SetConditionalFormat(sheet, rangeRef, testCase.format))
		ws, err := f.workSheetReader(sheet)
		assert.NoError(t, err)
		cf := ws.ConditionalFormatting
		assert.Len(t, cf, 1, testCase.label)
		assert.Len(t, cf[0].CfRule, 1, testCase.label)
		assert.Equal(t, rangeRef, cf[0].SQRef, testCase.label)
		assert.EqualValues(t, testCase.rules, cf[0].CfRule, testCase.label)
	}
	// Test creating a conditional format with a solid color data bar style
	f := NewFile()
	condFmts := []ConditionalFormatOptions{
		{Type: "data_bar", BarColor: "#A9D08E", BarSolid: true, Format: 0, Criteria: "=", MinType: "min", MaxType: "max"},
	}
	for _, ref := range []string{"A1:A2", "B1:B2"} {
		assert.NoError(t, f.SetConditionalFormat("Sheet1", ref, condFmts))
	}
	f = NewFile()
	// Test creating a conditional format with existing extension lists
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).ExtLst = &xlsxExtLst{Ext: `
		<ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}"><x14:slicerList /></ext>
		<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"><x14:sparklineGroups /></ext>`}
	assert.NoError(t, f.SetConditionalFormat("Sheet1", "A1:A2", []ConditionalFormatOptions{{Type: "data_bar", Criteria: "=", MinType: "min", MaxType: "max", BarBorderColor: "#0000FF", BarColor: "#638EC6", BarSolid: true}}))
	f = NewFile()
	// Test creating a conditional format with invalid extension list characters
	ws, ok = f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).ExtLst = &xlsxExtLst{Ext: "<ext><x14:conditionalFormattings></x14:conditionalFormatting></x14:conditionalFormattings></ext>"}
	assert.EqualError(t, f.SetConditionalFormat("Sheet1", "A1:A2", condFmts), "XML syntax error on line 1: element <conditionalFormattings> closed by </conditionalFormatting>")
	// Test creating a conditional format with invalid icon set style
	assert.EqualError(t, f.SetConditionalFormat("Sheet1", "A1:A2", []ConditionalFormatOptions{{Type: "icon_set", IconStyle: "unknown"}}), ErrParameterInvalid.Error())
}

func TestGetConditionalFormats(t *testing.T) {
	for _, format := range [][]ConditionalFormatOptions{
		{{Type: "cell", Format: 1, Criteria: "greater than", Value: "6"}},
		{{Type: "cell", Format: 1, Criteria: "between", MinValue: "6", MaxValue: "8"}},
		{{Type: "top", Format: 1, Criteria: "=", Value: "6"}},
		{{Type: "bottom", Format: 1, Criteria: "=", Value: "6"}},
		{{Type: "average", AboveAverage: true, Format: 1, Criteria: "="}},
		{{Type: "duplicate", Format: 1, Criteria: "="}},
		{{Type: "unique", Format: 1, Criteria: "="}},
		{{Type: "3_color_scale", Criteria: "=", MinType: "num", MidType: "num", MaxType: "num", MinValue: "-10", MidValue: "50", MaxValue: "10", MinColor: "#FF0000", MidColor: "#00FF00", MaxColor: "#0000FF"}},
		{{Type: "2_color_scale", Criteria: "=", MinType: "num", MaxType: "num", MinColor: "#FF0000", MaxColor: "#0000FF"}},
		{{Type: "data_bar", Criteria: "=", MinType: "num", MaxType: "num", MinValue: "-10", MaxValue: "10", BarBorderColor: "#0000FF", BarColor: "#638EC6", BarOnly: true, BarSolid: true, StopIfTrue: true}},
		{{Type: "data_bar", Criteria: "=", MinType: "min", MaxType: "max", BarBorderColor: "#0000FF", BarColor: "#638EC6", BarDirection: "rightToLeft", BarOnly: true, BarSolid: true, StopIfTrue: true}},
		{{Type: "formula", Format: 1, Criteria: "="}},
		{{Type: "icon_set", IconStyle: "3Arrows", ReverseIcons: true, IconsOnly: true}},
	} {
		f := NewFile()
		err := f.SetConditionalFormat("Sheet1", "A1:A2", format)
		assert.NoError(t, err)
		opts, err := f.GetConditionalFormats("Sheet1")
		assert.NoError(t, err)
		assert.Equal(t, format, opts["A1:A2"])
	}
	// Test get conditional formats on no exists worksheet
	f := NewFile()
	_, err := f.GetConditionalFormats("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Test get conditional formats with invalid sheet name
	_, err = f.GetConditionalFormats("Sheet:1")
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())
}

func TestUnsetConditionalFormat(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 7))
	assert.NoError(t, f.UnsetConditionalFormat("Sheet1", "A1:A10"))
	format, err := f.NewConditionalStyle(&Style{Font: &Font{Color: "9A0511"}, Fill: Fill{Type: "pattern", Color: []string{"FEC7CE"}, Pattern: 1}})
	assert.NoError(t, err)
	assert.NoError(t, f.SetConditionalFormat("Sheet1", "A1:A10", []ConditionalFormatOptions{{Type: "cell", Criteria: ">", Format: format, Value: "6"}}))
	assert.NoError(t, f.UnsetConditionalFormat("Sheet1", "A1:A10"))
	// Test unset conditional format on not exists worksheet
	assert.EqualError(t, f.UnsetConditionalFormat("SheetN", "A1:A10"), "sheet SheetN does not exist")
	// Test unset conditional format with invalid sheet name
	assert.EqualError(t, f.UnsetConditionalFormat("Sheet:1", "A1:A10"), ErrSheetNameInvalid.Error())
	// Save spreadsheet by the given path
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestUnsetConditionalFormat.xlsx")))
}

func TestNewStyle(t *testing.T) {
	f := NewFile()
	for i := 0; i < 18; i++ {
		_, err := f.NewStyle(&Style{
			Fill: Fill{Type: "gradient", Color: []string{"FFFFFF", "4E71BE"}, Shading: i},
		})
		assert.NoError(t, err)
	}
	f = NewFile()
	styleID, err := f.NewStyle(&Style{Font: &Font{Bold: true, Italic: true, Family: "Times New Roman", Size: 36, Color: "777777"}})
	assert.NoError(t, err)
	styles, err := f.stylesReader()
	assert.NoError(t, err)
	fontID := styles.CellXfs.Xf[styleID].FontID
	font := styles.Fonts.Font[*fontID]
	assert.Contains(t, *font.Name.Val, "Times New Roman", "Stored font should contain font name")
	assert.Equal(t, 2, styles.CellXfs.Count, "Should have 2 styles")
	_, err = f.NewStyle(&Style{})
	assert.NoError(t, err)
	_, err = f.NewStyle(nil)
	assert.NoError(t, err)

	var exp string
	_, err = f.NewStyle(&Style{CustomNumFmt: &exp})
	assert.EqualError(t, err, ErrCustomNumFmt.Error())
	_, err = f.NewStyle(&Style{Font: &Font{Family: strings.Repeat("s", MaxFontFamilyLength+1)}})
	assert.EqualError(t, err, ErrFontLength.Error())
	_, err = f.NewStyle(&Style{Font: &Font{Size: MaxFontSize + 1}})
	assert.EqualError(t, err, ErrFontSize.Error())

	// Test create numeric custom style
	numFmt := "####;####"
	f.Styles.NumFmts = nil
	styleID, err = f.NewStyle(&Style{
		CustomNumFmt: &numFmt,
	})
	assert.NoError(t, err)
	assert.Equal(t, 2, styleID)

	assert.NotNil(t, f.Styles)
	assert.NotNil(t, f.Styles.CellXfs)
	assert.NotNil(t, f.Styles.CellXfs.Xf)

	nf := f.Styles.CellXfs.Xf[styleID]
	assert.Equal(t, 164, *nf.NumFmtID)

	// Test create currency custom style
	f.Styles.NumFmts = nil
	styleID, err = f.NewStyle(&Style{
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
	style4, err := f.NewStyle(&Style{NumFmt: 160})
	assert.NoError(t, err)
	assert.Equal(t, 0, style4)

	f = NewFile()
	f.Styles.NumFmts = nil
	f.Styles.CellXfs.Xf = nil
	style5, err := f.NewStyle(&Style{NumFmt: 160})
	assert.NoError(t, err)
	assert.Equal(t, 0, style5)

	// Test create style with unsupported charset style sheet
	f.Styles = nil
	f.Pkg.Store(defaultXMLPathStyles, MacintoshCyrillicCharset)
	_, err = f.NewStyle(&Style{NumFmt: 165})
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")

	// Test create cell styles reach maximum
	f = NewFile()
	f.Styles.CellXfs.Xf = make([]xlsxXf, MaxCellStyles)
	f.Styles.CellXfs.Count = MaxCellStyles
	_, err = f.NewStyle(&Style{NumFmt: 0})
	assert.Equal(t, ErrCellStyles, err)
}

func TestNewConditionalStyle(t *testing.T) {
	f := NewFile()
	_, err := f.NewConditionalStyle(&Style{Protection: &Protection{Hidden: true, Locked: true}})
	assert.NoError(t, err)
	_, err = f.NewConditionalStyle(&Style{DecimalPlaces: intPtr(4), NumFmt: 165, NegRed: true})
	assert.NoError(t, err)
	_, err = f.NewConditionalStyle(&Style{DecimalPlaces: intPtr(-1)})
	assert.NoError(t, err)
	_, err = f.NewConditionalStyle(&Style{NumFmt: 1})
	assert.NoError(t, err)
	_, err = f.NewConditionalStyle(&Style{NumFmt: 27})
	assert.NoError(t, err)
	numFmt := "general"
	_, err = f.NewConditionalStyle(&Style{CustomNumFmt: &numFmt})
	assert.NoError(t, err)
	numFmt1 := "0.00"
	_, err = f.NewConditionalStyle(&Style{CustomNumFmt: &numFmt1})
	assert.NoError(t, err)
	// Test create conditional style with unsupported charset style sheet
	f.Styles = nil
	f.Pkg.Store(defaultXMLPathStyles, MacintoshCyrillicCharset)
	_, err = f.NewConditionalStyle(&Style{Font: &Font{Color: "9A0511"}, Fill: Fill{Type: "pattern", Color: []string{"FEC7CE"}, Pattern: 1}})
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
}

func TestGetDefaultFont(t *testing.T) {
	f := NewFile()
	s, err := f.GetDefaultFont()
	assert.NoError(t, err)
	assert.Equal(t, s, "Calibri", "Default font should be Calibri")
	// Test get default font with unsupported charset style sheet
	f.Styles = nil
	f.Pkg.Store(defaultXMLPathStyles, MacintoshCyrillicCharset)
	_, err = f.GetDefaultFont()
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
}

func TestSetDefaultFont(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetDefaultFont("Arial"))
	styles, err := f.stylesReader()
	assert.NoError(t, err)
	s, err := f.GetDefaultFont()
	assert.NoError(t, err)
	assert.Equal(t, s, "Arial", "Default font should change to Arial")
	assert.Equal(t, *styles.CellStyles.CellStyle[0].CustomBuiltIn, true)
	// Test set default font with unsupported charset style sheet
	f.Styles = nil
	f.Pkg.Store(defaultXMLPathStyles, MacintoshCyrillicCharset)
	assert.EqualError(t, f.SetDefaultFont("Arial"), "XML syntax error on line 1: invalid UTF-8")
}

func TestStylesReader(t *testing.T) {
	f := NewFile()
	// Test read styles with unsupported charset
	f.Styles = nil
	f.Pkg.Store(defaultXMLPathStyles, MacintoshCyrillicCharset)
	styles, err := f.stylesReader()
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	assert.EqualValues(t, new(xlsxStyleSheet), styles)
}

func TestThemeReader(t *testing.T) {
	f := NewFile()
	// Test read theme with unsupported charset
	f.Pkg.Store(defaultXMLPathTheme, MacintoshCyrillicCharset)
	theme, err := f.themeReader()
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	assert.EqualValues(t, &xlsxTheme{XMLNSa: NameSpaceDrawingML.Value, XMLNSr: SourceRelationship.Value}, theme)
}

func TestSetCellStyle(t *testing.T) {
	f := NewFile()
	// Test set cell style on not exists worksheet
	assert.EqualError(t, f.SetCellStyle("SheetN", "A1", "A2", 1), "sheet SheetN does not exist")
	// Test set cell style with invalid style ID
	assert.EqualError(t, f.SetCellStyle("Sheet1", "A1", "A2", -1), newInvalidStyleID(-1).Error())
	// Test set cell style with not exists style ID
	assert.EqualError(t, f.SetCellStyle("Sheet1", "A1", "A2", 10), newInvalidStyleID(10).Error())
	// Test set cell style with unsupported charset style sheet
	f.Styles = nil
	f.Pkg.Store(defaultXMLPathStyles, MacintoshCyrillicCharset)
	assert.EqualError(t, f.SetCellStyle("Sheet1", "A1", "A2", 1), "XML syntax error on line 1: invalid UTF-8")
}

func TestGetStyleID(t *testing.T) {
	f := NewFile()
	styleID, err := f.getStyleID(&xlsxStyleSheet{}, nil)
	assert.NoError(t, err)
	assert.Equal(t, -1, styleID)
	// Test get style ID with unsupported charset style sheet
	f.Styles = nil
	f.Pkg.Store(defaultXMLPathStyles, MacintoshCyrillicCharset)
	_, err = f.getStyleID(&xlsxStyleSheet{
		CellXfs: &xlsxCellXfs{},
		Fonts: &xlsxFonts{
			Font: []*xlsxFont{{}},
		},
	}, &Style{NumFmt: 0, Font: &Font{}})
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
}

func TestGetFillID(t *testing.T) {
	styles, err := NewFile().stylesReader()
	assert.NoError(t, err)
	assert.Equal(t, -1, getFillID(styles, &Style{Fill: Fill{Type: "unknown"}}))
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

	fs1, err := parseFormatStyleSet(&Style{Protection: &Protection{Hidden: false, Locked: false}, NumFmt: 10})
	assert.NoError(t, err)
	id1 := getNumFmtID(&xlsxStyleSheet{}, fs1)

	fs2, err := parseFormatStyleSet(&Style{Protection: &Protection{Hidden: false, Locked: false}, NumFmt: 0})
	assert.NoError(t, err)
	id2 := getNumFmtID(&xlsxStyleSheet{}, fs2)

	assert.NotEqual(t, id1, id2)
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestStyleNumFmt.xlsx")))
}

func TestGetThemeColor(t *testing.T) {
	assert.Empty(t, (&File{}).getThemeColor(&xlsxColor{}))
	f := NewFile()
	assert.Empty(t, f.getThemeColor(nil))
	var theme int
	assert.Equal(t, "FFFFFF", f.getThemeColor(&xlsxColor{Theme: &theme}))
	assert.Equal(t, "FFFFFF", f.getThemeColor(&xlsxColor{RGB: "FFFFFF"}))
	assert.Equal(t, "FF8080", f.getThemeColor(&xlsxColor{Indexed: 2, Tint: 0.5}))
	assert.Empty(t, f.getThemeColor(&xlsxColor{Indexed: len(IndexedColorMapping), Tint: 0.5}))
}

func TestGetStyle(t *testing.T) {
	f := NewFile()
	expected := &Style{
		Border: []Border{
			{Type: "left", Color: "0000FF", Style: 3},
			{Type: "right", Color: "FF0000", Style: 6},
			{Type: "top", Color: "00FF00", Style: 4},
			{Type: "bottom", Color: "FFFF00", Style: 5},
			{Type: "diagonalUp", Color: "A020F0", Style: 7},
			{Type: "diagonalDown", Color: "A020F0", Style: 7},
		},
		Fill: Fill{Type: "gradient", Shading: 16, Color: []string{"0000FF", "00FF00"}},
		Font: &Font{
			Bold: true, Italic: true, Underline: "single", Family: "Arial",
			Size: 8.5, Strike: true, Color: "777777", ColorIndexed: 1, ColorTint: 0.1,
		},
		Alignment: &Alignment{
			Horizontal:      "center",
			Indent:          1,
			JustifyLastLine: true,
			ReadingOrder:    1,
			RelativeIndent:  1,
			ShrinkToFit:     true,
			TextRotation:    180,
			Vertical:        "center",
			WrapText:        true,
		},
		Protection: &Protection{Hidden: true, Locked: true},
		NumFmt:     49,
	}
	styleID, err := f.NewStyle(expected)
	assert.NoError(t, err)
	style, err := f.GetStyle(styleID)
	assert.NoError(t, err)
	assert.Equal(t, expected.Border, style.Border)
	assert.Equal(t, expected.Fill, style.Fill)
	assert.Equal(t, expected.Font, style.Font)
	assert.Equal(t, expected.Alignment, style.Alignment)
	assert.Equal(t, expected.Protection, style.Protection)
	assert.Equal(t, expected.NumFmt, style.NumFmt)

	expected = &Style{
		Fill: Fill{Type: "pattern", Pattern: 1, Color: []string{"0000FF"}},
	}
	styleID, err = f.NewStyle(expected)
	assert.NoError(t, err)
	style, err = f.GetStyle(styleID)
	assert.NoError(t, err)
	assert.Equal(t, expected.Fill, style.Fill)

	expected = &Style{NumFmt: 27}
	styleID, err = f.NewStyle(expected)
	assert.NoError(t, err)
	style, err = f.GetStyle(styleID)
	assert.NoError(t, err)
	assert.Equal(t, expected.NumFmt, style.NumFmt)

	expected = &Style{NumFmt: 165}
	styleID, err = f.NewStyle(expected)
	assert.NoError(t, err)
	style, err = f.GetStyle(styleID)
	assert.NoError(t, err)
	assert.Equal(t, expected.NumFmt, style.NumFmt)

	expected = &Style{NumFmt: 165, NegRed: true}
	styleID, err = f.NewStyle(expected)
	assert.NoError(t, err)
	style, err = f.GetStyle(styleID)
	assert.NoError(t, err)
	assert.Equal(t, expected.NumFmt, style.NumFmt)

	// Test get style with custom color index
	f.Styles.Colors = &xlsxStyleColors{
		IndexedColors: xlsxIndexedColors{
			RgbColor: []xlsxColor{{RGB: "FF012345"}},
		},
	}
	assert.Equal(t, "012345", f.getThemeColor(&xlsxColor{Indexed: 0}))

	// Test get style with invalid style index
	style, err = f.GetStyle(-1)
	assert.Nil(t, style)
	assert.Equal(t, err, newInvalidStyleID(-1))
	// Test get style with unsupported charset style sheet
	f.Styles = nil
	f.Pkg.Store(defaultXMLPathStyles, MacintoshCyrillicCharset)
	style, err = f.GetStyle(1)
	assert.Nil(t, style)
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
}
