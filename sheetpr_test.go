package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestSetPageMargins(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetPageMargins("Sheet1", nil))
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).PageMargins = nil
	ws.(*xlsxWorksheet).PrintOptions = nil
	expected := PageLayoutMarginsOptions{
		Bottom:       float64Ptr(1.0),
		Footer:       float64Ptr(1.0),
		Header:       float64Ptr(1.0),
		Left:         float64Ptr(1.0),
		Right:        float64Ptr(1.0),
		Top:          float64Ptr(1.0),
		Horizontally: boolPtr(true),
		Vertically:   boolPtr(true),
	}
	assert.NoError(t, f.SetPageMargins("Sheet1", &expected))
	opts, err := f.GetPageMargins("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, expected, opts)
	// Test set page margins on not exists worksheet
	assert.EqualError(t, f.SetPageMargins("SheetN", nil), "sheet SheetN does not exist")
	// Test set page margins with invalid sheet name
	assert.Equal(t, ErrSheetNameInvalid, f.SetPageMargins("Sheet:1", nil))
}

func TestGetPageMargins(t *testing.T) {
	f := NewFile()
	// Test get page margins on not exists worksheet
	_, err := f.GetPageMargins("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Test get page margins with invalid sheet name
	_, err = f.GetPageMargins("Sheet:1")
	assert.Equal(t, ErrSheetNameInvalid, err)
}

func TestSetSheetProps(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetSheetProps("Sheet1", nil))
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).SheetPr = nil
	ws.(*xlsxWorksheet).SheetFormatPr = nil
	baseColWidth, enable := uint8(8), boolPtr(true)
	expected := SheetPropsOptions{
		CodeName:                          stringPtr("code"),
		EnableFormatConditionsCalculation: enable,
		Published:                         enable,
		AutoPageBreaks:                    enable,
		FitToPage:                         enable,
		TabColorIndexed:                   intPtr(1),
		TabColorRGB:                       stringPtr("FFFF00"),
		TabColorTheme:                     intPtr(1),
		TabColorTint:                      float64Ptr(1),
		OutlineSummaryBelow:               enable,
		OutlineSummaryRight:               enable,
		BaseColWidth:                      &baseColWidth,
		DefaultColWidth:                   float64Ptr(10),
		DefaultRowHeight:                  float64Ptr(10),
		CustomHeight:                      enable,
		ZeroHeight:                        enable,
		ThickTop:                          enable,
		ThickBottom:                       enable,
	}
	assert.NoError(t, f.SetSheetProps("Sheet1", &expected))
	opts, err := f.GetSheetProps("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, expected, opts)

	ws.(*xlsxWorksheet).SheetPr = nil
	assert.NoError(t, f.SetSheetProps("Sheet1", &SheetPropsOptions{FitToPage: enable}))
	ws.(*xlsxWorksheet).SheetPr = nil
	assert.NoError(t, f.SetSheetProps("Sheet1", &SheetPropsOptions{TabColorRGB: stringPtr("FFFF00")}))
	ws.(*xlsxWorksheet).SheetPr = nil
	assert.NoError(t, f.SetSheetProps("Sheet1", &SheetPropsOptions{TabColorTheme: intPtr(1)}))
	ws.(*xlsxWorksheet).SheetPr = nil
	assert.NoError(t, f.SetSheetProps("Sheet1", &SheetPropsOptions{TabColorTint: float64Ptr(1)}))

	// Test set worksheet properties on not exists worksheet
	assert.EqualError(t, f.SetSheetProps("SheetN", nil), "sheet SheetN does not exist")
	// Test set worksheet properties with invalid sheet name
	assert.Equal(t, ErrSheetNameInvalid, f.SetSheetProps("Sheet:1", nil))
}

func TestGetSheetProps(t *testing.T) {
	f := NewFile()
	// Test get worksheet properties on not exists worksheet
	_, err := f.GetSheetProps("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Test get worksheet properties with invalid sheet name
	_, err = f.GetSheetProps("Sheet:1")
	assert.Equal(t, ErrSheetNameInvalid, err)
}
