package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

var _ = []SheetViewOption{
	DefaultGridColor(true),
	ShowFormulas(false),
	ShowGridLines(true),
	ShowRowColHeaders(true),
	ShowZeros(true),
	RightToLeft(false),
	ShowRuler(false),
	View("pageLayout"),
	TopLeftCell("B2"),
	ZoomScale(100),
	// SheetViewOptionPtr are also SheetViewOption
	new(DefaultGridColor),
	new(ShowFormulas),
	new(ShowGridLines),
	new(ShowRowColHeaders),
	new(ShowZeros),
	new(RightToLeft),
	new(ShowRuler),
	new(View),
	new(TopLeftCell),
	new(ZoomScale),
}

var _ = []SheetViewOptionPtr{
	(*DefaultGridColor)(nil),
	(*ShowFormulas)(nil),
	(*ShowGridLines)(nil),
	(*ShowRowColHeaders)(nil),
	(*ShowZeros)(nil),
	(*RightToLeft)(nil),
	(*ShowRuler)(nil),
	(*View)(nil),
	(*TopLeftCell)(nil),
	(*ZoomScale)(nil),
}

func ExampleFile_SetSheetViewOptions() {
	f := NewFile()
	const sheet = "Sheet1"

	if err := f.SetSheetViewOptions(sheet, 0,
		DefaultGridColor(false),
		ShowFormulas(true),
		ShowGridLines(true),
		ShowRowColHeaders(true),
		RightToLeft(false),
		ShowRuler(false),
		View("pageLayout"),
		TopLeftCell("C3"),
		ZoomScale(80),
	); err != nil {
		fmt.Println(err)
	}

	var zoomScale ZoomScale
	fmt.Println("Default:")
	fmt.Println("- zoomScale: 80")

	if err := f.SetSheetViewOptions(sheet, 0, ZoomScale(500)); err != nil {
		fmt.Println(err)
	}

	if err := f.GetSheetViewOptions(sheet, 0, &zoomScale); err != nil {
		fmt.Println(err)
	}

	fmt.Println("Used out of range value:")
	fmt.Println("- zoomScale:", zoomScale)

	if err := f.SetSheetViewOptions(sheet, 0, ZoomScale(123)); err != nil {
		fmt.Println(err)
	}

	if err := f.GetSheetViewOptions(sheet, 0, &zoomScale); err != nil {
		fmt.Println(err)
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
	f := NewFile()
	const sheet = "Sheet1"

	var (
		defaultGridColor  DefaultGridColor
		showFormulas      ShowFormulas
		showGridLines     ShowGridLines
		showRowColHeaders ShowRowColHeaders
		showZeros         ShowZeros
		rightToLeft       RightToLeft
		showRuler         ShowRuler
		view              View
		topLeftCell       TopLeftCell
		zoomScale         ZoomScale
	)

	if err := f.GetSheetViewOptions(sheet, 0,
		&defaultGridColor,
		&showFormulas,
		&showGridLines,
		&showRowColHeaders,
		&showZeros,
		&rightToLeft,
		&showRuler,
		&view,
		&topLeftCell,
		&zoomScale,
	); err != nil {
		fmt.Println(err)
	}

	fmt.Println("Default:")
	fmt.Println("- defaultGridColor:", defaultGridColor)
	fmt.Println("- showFormulas:", showFormulas)
	fmt.Println("- showGridLines:", showGridLines)
	fmt.Println("- showRowColHeaders:", showRowColHeaders)
	fmt.Println("- showZeros:", showZeros)
	fmt.Println("- rightToLeft:", rightToLeft)
	fmt.Println("- showRuler:", showRuler)
	fmt.Println("- view:", view)
	fmt.Println("- topLeftCell:", `"`+topLeftCell+`"`)
	fmt.Println("- zoomScale:", zoomScale)

	if err := f.SetSheetViewOptions(sheet, 0, ShowGridLines(false)); err != nil {
		fmt.Println(err)
	}

	if err := f.GetSheetViewOptions(sheet, 0, &showGridLines); err != nil {
		fmt.Println(err)
	}

	if err := f.SetSheetViewOptions(sheet, 0, ShowZeros(false)); err != nil {
		fmt.Println(err)
	}

	if err := f.GetSheetViewOptions(sheet, 0, &showZeros); err != nil {
		fmt.Println(err)
	}

	if err := f.SetSheetViewOptions(sheet, 0, View("pageLayout")); err != nil {
		fmt.Println(err)
	}

	if err := f.GetSheetViewOptions(sheet, 0, &view); err != nil {
		fmt.Println(err)
	}

	if err := f.SetSheetViewOptions(sheet, 0, TopLeftCell("B2")); err != nil {
		fmt.Println(err)
	}

	if err := f.GetSheetViewOptions(sheet, 0, &topLeftCell); err != nil {
		fmt.Println(err)
	}

	fmt.Println("After change:")
	fmt.Println("- showGridLines:", showGridLines)
	fmt.Println("- showZeros:", showZeros)
	fmt.Println("- view:", view)
	fmt.Println("- topLeftCell:", topLeftCell)

	// Output:
	// Default:
	// - defaultGridColor: true
	// - showFormulas: false
	// - showGridLines: true
	// - showRowColHeaders: true
	// - showZeros: true
	// - rightToLeft: false
	// - showRuler: true
	// - view: normal
	// - topLeftCell: ""
	// - zoomScale: 0
	// After change:
	// - showGridLines: false
	// - showZeros: false
	// - view: pageLayout
	// - topLeftCell: B2
}

func TestSheetViewOptionsErrors(t *testing.T) {
	f := NewFile()
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
