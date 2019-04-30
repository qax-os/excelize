package excelize_test

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

var _ = []excelize.SheetViewOption{
	excelize.DefaultGridColor(true),
	excelize.RightToLeft(false),
	excelize.ShowFormulas(false),
	excelize.ShowGridLines(true),
	excelize.ShowRowColHeaders(true),
	excelize.TopLeftCell("B2"),
	// SheetViewOptionPtr are also SheetViewOption
	new(excelize.DefaultGridColor),
	new(excelize.RightToLeft),
	new(excelize.ShowFormulas),
	new(excelize.ShowGridLines),
	new(excelize.ShowRowColHeaders),
	new(excelize.TopLeftCell),
}

var _ = []excelize.SheetViewOptionPtr{
	(*excelize.DefaultGridColor)(nil),
	(*excelize.RightToLeft)(nil),
	(*excelize.ShowFormulas)(nil),
	(*excelize.ShowGridLines)(nil),
	(*excelize.ShowRowColHeaders)(nil),
	(*excelize.TopLeftCell)(nil),
}

func ExampleFile_SetSheetViewOptions() {
	f := excelize.NewFile()
	const sheet = "Sheet1"

	if err := f.SetSheetViewOptions(sheet, 0,
		excelize.DefaultGridColor(false),
		excelize.RightToLeft(false),
		excelize.ShowFormulas(true),
		excelize.ShowGridLines(true),
		excelize.ShowRowColHeaders(true),
		excelize.ZoomScale(80),
		excelize.TopLeftCell("C3"),
	); err != nil {
		panic(err)
	}

	var zoomScale excelize.ZoomScale
	fmt.Println("Default:")
	fmt.Println("- zoomScale: 80")

	if err := f.SetSheetViewOptions(sheet, 0, excelize.ZoomScale(500)); err != nil {
		panic(err)
	}

	if err := f.GetSheetViewOptions(sheet, 0, &zoomScale); err != nil {
		panic(err)
	}

	fmt.Println("Used out of range value:")
	fmt.Println("- zoomScale:", zoomScale)

	if err := f.SetSheetViewOptions(sheet, 0, excelize.ZoomScale(123)); err != nil {
		panic(err)
	}

	if err := f.GetSheetViewOptions(sheet, 0, &zoomScale); err != nil {
		panic(err)
	}

	fmt.Println("Used correct value:")
	fmt.Println("- zoomScale:", zoomScale)

	// Output:
	// Default:
	// - zoomScale: 80
	// Used out of range value:
	// - zoomScale: 80
	// Used correct value:
	// - zoomScale: 123

}

func ExampleFile_GetSheetViewOptions() {
	f := excelize.NewFile()
	const sheet = "Sheet1"

	var (
		defaultGridColor  excelize.DefaultGridColor
		rightToLeft       excelize.RightToLeft
		showFormulas      excelize.ShowFormulas
		showGridLines     excelize.ShowGridLines
		showRowColHeaders excelize.ShowRowColHeaders
		zoomScale         excelize.ZoomScale
		topLeftCell       excelize.TopLeftCell
	)

	if err := f.GetSheetViewOptions(sheet, 0,
		&defaultGridColor,
		&rightToLeft,
		&showFormulas,
		&showGridLines,
		&showRowColHeaders,
		&zoomScale,
		&topLeftCell,
	); err != nil {
		panic(err)
	}

	fmt.Println("Default:")
	fmt.Println("- defaultGridColor:", defaultGridColor)
	fmt.Println("- rightToLeft:", rightToLeft)
	fmt.Println("- showFormulas:", showFormulas)
	fmt.Println("- showGridLines:", showGridLines)
	fmt.Println("- showRowColHeaders:", showRowColHeaders)
	fmt.Println("- zoomScale:", zoomScale)
	fmt.Println("- topLeftCell:", `"`+topLeftCell+`"`)

	if err := f.SetSheetViewOptions(sheet, 0, excelize.TopLeftCell("B2")); err != nil {
		panic(err)
	}

	if err := f.GetSheetViewOptions(sheet, 0, &topLeftCell); err != nil {
		panic(err)
	}

	if err := f.SetSheetViewOptions(sheet, 0, excelize.ShowGridLines(false)); err != nil {
		panic(err)
	}

	if err := f.GetSheetViewOptions(sheet, 0, &showGridLines); err != nil {
		panic(err)
	}

	fmt.Println("After change:")
	fmt.Println("- showGridLines:", showGridLines)
	fmt.Println("- topLeftCell:", topLeftCell)

	// Output:
	// Default:
	// - defaultGridColor: true
	// - rightToLeft: false
	// - showFormulas: false
	// - showGridLines: true
	// - showRowColHeaders: true
	// - zoomScale: 0
	// - topLeftCell: ""
	// After change:
	// - showGridLines: false
	// - topLeftCell: B2
}

func TestSheetViewOptionsErrors(t *testing.T) {
	f := excelize.NewFile()
	const sheet = "Sheet1"

	assert.NoError(t, f.GetSheetViewOptions(sheet, 0))
	assert.NoError(t, f.GetSheetViewOptions(sheet, -1))
	assert.Error(t, f.GetSheetViewOptions(sheet, 1))
	assert.Error(t, f.GetSheetViewOptions(sheet, -2))
	assert.NoError(t, f.SetSheetViewOptions(sheet, 0))
	assert.NoError(t, f.SetSheetViewOptions(sheet, -1))
	assert.Error(t, f.SetSheetViewOptions(sheet, 1))
	assert.Error(t, f.SetSheetViewOptions(sheet, -2))
}
