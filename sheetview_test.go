package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestSetView(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetSheetView("Sheet1", -1, nil))
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).SheetViews = nil
	expected := ViewOptions{
		DefaultGridColor:  boolPtr(false),
		RightToLeft:       boolPtr(false),
		ShowFormulas:      boolPtr(false),
		ShowGridLines:     boolPtr(false),
		ShowRowColHeaders: boolPtr(false),
		ShowRuler:         boolPtr(false),
		ShowZeros:         boolPtr(false),
		TopLeftCell:       stringPtr("A1"),
		View:              stringPtr("normal"),
		ZoomScale:         float64Ptr(120),
	}
	assert.NoError(t, f.SetSheetView("Sheet1", 0, &expected))
	opts, err := f.GetSheetView("Sheet1", 0)
	assert.NoError(t, err)
	assert.Equal(t, expected, opts)
	// Test set sheet view options with invalid view index
	assert.EqualError(t, f.SetSheetView("Sheet1", 1, nil), "view index 1 out of range")
	assert.EqualError(t, f.SetSheetView("Sheet1", -2, nil), "view index -2 out of range")
	// Test set sheet view options on not exists worksheet
	assert.EqualError(t, f.SetSheetView("SheetN", 0, nil), "sheet SheetN does not exist")
}

func TestGetView(t *testing.T) {
	f := NewFile()
	_, err := f.getSheetView("SheetN", 0)
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Test get sheet view options with invalid view index
	_, err = f.GetSheetView("Sheet1", 1)
	assert.EqualError(t, err, "view index 1 out of range")
	_, err = f.GetSheetView("Sheet1", -2)
	assert.EqualError(t, err, "view index -2 out of range")
	// Test get sheet view options on not exists worksheet
	_, err = f.GetSheetView("SheetN", 0)
	assert.EqualError(t, err, "sheet SheetN does not exist")
}
