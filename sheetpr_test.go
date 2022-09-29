package excelize

import (
	"path/filepath"
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
	// Test set page margins on not exists worksheet.
	assert.EqualError(t, f.SetPageMargins("SheetN", nil), "sheet SheetN does not exist")
}

func TestGetPageMargins(t *testing.T) {
	f := NewFile()
	// Test get page margins on not exists worksheet.
	_, err := f.GetPageMargins("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
}

func TestDebug(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetSheetProps("Sheet1", nil))
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).PageMargins = nil
	ws.(*xlsxWorksheet).PrintOptions = nil
	ws.(*xlsxWorksheet).SheetPr = nil
	ws.(*xlsxWorksheet).SheetFormatPr = nil
	// w := uint8(10)
	// f.SetSheetProps("Sheet1", &SheetPropsOptions{BaseColWidth: &w})
	f.SetPageMargins("Sheet1", &PageLayoutMarginsOptions{Horizontally: boolPtr(true)})
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestDebug.xlsx")))
}

func TestSetSheetProps(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetSheetProps("Sheet1", nil))
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).SheetPr = nil
	ws.(*xlsxWorksheet).SheetFormatPr = nil
	baseColWidth := uint8(8)
	expected := SheetPropsOptions{
		CodeName:                          stringPtr("code"),
		EnableFormatConditionsCalculation: boolPtr(true),
		Published:                         boolPtr(true),
		AutoPageBreaks:                    boolPtr(true),
		FitToPage:                         boolPtr(true),
		TabColorIndexed:                   intPtr(1),
		TabColorRGB:                       stringPtr("#FFFF00"),
		TabColorTheme:                     intPtr(1),
		TabColorTint:                      float64Ptr(1),
		OutlineSummaryBelow:               boolPtr(true),
		BaseColWidth:                      &baseColWidth,
		DefaultColWidth:                   float64Ptr(10),
		DefaultRowHeight:                  float64Ptr(10),
		CustomHeight:                      boolPtr(true),
		ZeroHeight:                        boolPtr(true),
		ThickTop:                          boolPtr(true),
		ThickBottom:                       boolPtr(true),
	}
	assert.NoError(t, f.SetSheetProps("Sheet1", &expected))
	opts, err := f.GetSheetProps("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, expected, opts)

	ws.(*xlsxWorksheet).SheetPr = nil
	assert.NoError(t, f.SetSheetProps("Sheet1", &SheetPropsOptions{FitToPage: boolPtr(true)}))
	ws.(*xlsxWorksheet).SheetPr = nil
	assert.NoError(t, f.SetSheetProps("Sheet1", &SheetPropsOptions{TabColorRGB: stringPtr("#FFFF00")}))
	ws.(*xlsxWorksheet).SheetPr = nil
	assert.NoError(t, f.SetSheetProps("Sheet1", &SheetPropsOptions{TabColorTheme: intPtr(1)}))
	ws.(*xlsxWorksheet).SheetPr = nil
	assert.NoError(t, f.SetSheetProps("Sheet1", &SheetPropsOptions{TabColorTint: float64Ptr(1)}))

	// Test SetSheetProps on not exists worksheet.
	assert.EqualError(t, f.SetSheetProps("SheetN", nil), "sheet SheetN does not exist")
}

func TestGetSheetProps(t *testing.T) {
	f := NewFile()
	// Test GetSheetProps on not exists worksheet.
	_, err := f.GetSheetProps("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
}
