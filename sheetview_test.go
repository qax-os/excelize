package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

var _ = []SheetViewOption{
	DefaultGridColor(true),
	RightToLeft(false),
	ShowFormulas(false),
	ShowGridLines(true),
	ShowRowColHeaders(true),
	TopLeftCell("B2"),
	// SheetViewOptionPtr are also SheetViewOption
	new(DefaultGridColor),
	new(RightToLeft),
	new(ShowFormulas),
	new(ShowGridLines),
	new(ShowRowColHeaders),
	new(TopLeftCell),
}

var _ = []SheetViewOptionPtr{
	(*DefaultGridColor)(nil),
	(*RightToLeft)(nil),
	(*ShowFormulas)(nil),
	(*ShowGridLines)(nil),
	(*ShowRowColHeaders)(nil),
	(*TopLeftCell)(nil),
}

func ExampleFile_SetSheetViewOptions() {
	f := NewFile()
	const sheet = "Sheet1"

	if err := f.SetSheetViewOptions(sheet, 0,
		DefaultGridColor(false),
		RightToLeft(false),
		ShowFormulas(true),
		ShowGridLines(true),
		ShowRowColHeaders(true),
		ZoomScale(80),
		TopLeftCell("C3"),
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
		rightToLeft       RightToLeft
		showFormulas      ShowFormulas
		showGridLines     ShowGridLines
		showZeros         ShowZeros
		showRowColHeaders ShowRowColHeaders
		zoomScale         ZoomScale
		topLeftCell       TopLeftCell
	)

	if err := f.GetSheetViewOptions(sheet, 0,
		&defaultGridColor,
		&rightToLeft,
		&showFormulas,
		&showGridLines,
		&showZeros,
		&showRowColHeaders,
		&zoomScale,
		&topLeftCell,
	); err != nil {
		fmt.Println(err)
	}

	fmt.Println("Default:")
	fmt.Println("- defaultGridColor:", defaultGridColor)
	fmt.Println("- rightToLeft:", rightToLeft)
	fmt.Println("- showFormulas:", showFormulas)
	fmt.Println("- showGridLines:", showGridLines)
	fmt.Println("- showZeros:", showZeros)
	fmt.Println("- showRowColHeaders:", showRowColHeaders)
	fmt.Println("- zoomScale:", zoomScale)
	fmt.Println("- topLeftCell:", `"`+topLeftCell+`"`)

	if err := f.SetSheetViewOptions(sheet, 0, TopLeftCell("B2")); err != nil {
		fmt.Println(err)
	}

	if err := f.GetSheetViewOptions(sheet, 0, &topLeftCell); err != nil {
		fmt.Println(err)
	}

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

	fmt.Println("After change:")
	fmt.Println("- showGridLines:", showGridLines)
	fmt.Println("- showZeros:", showZeros)
	fmt.Println("- topLeftCell:", topLeftCell)

	// Output:
	// Default:
	// - defaultGridColor: true
	// - rightToLeft: false
	// - showFormulas: false
	// - showGridLines: true
	// - showZeros: true
	// - showRowColHeaders: true
	// - zoomScale: 0
	// - topLeftCell: ""
	// After change:
	// - showGridLines: false
	// - showZeros: false
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
