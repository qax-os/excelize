package excelize_test

import (
	"fmt"
	"testing"

	"github.com/360EntSecGroup-Skylar/excelize"
)

var _ = []excelize.SheetViewOption{
	excelize.DefaultGridColor(true),
	excelize.RightToLeft(false),
	excelize.ShowFormulas(false),
	excelize.ShowGridLines(true),
	excelize.ShowRowColHeaders(true),
	// SheetViewOptionPtr are also SheetViewOption
	new(excelize.DefaultGridColor),
	new(excelize.RightToLeft),
	new(excelize.ShowFormulas),
	new(excelize.ShowGridLines),
	new(excelize.ShowRowColHeaders),
}

var _ = []excelize.SheetViewOptionPtr{
	(*excelize.DefaultGridColor)(nil),
	(*excelize.RightToLeft)(nil),
	(*excelize.ShowFormulas)(nil),
	(*excelize.ShowGridLines)(nil),
	(*excelize.ShowRowColHeaders)(nil),
}

func ExampleFile_SetSheetViewOptions() {
	xl := excelize.NewFile()
	const sheet = "Sheet1"

	if err := xl.SetSheetViewOptions(sheet, 0,
		excelize.DefaultGridColor(false),
		excelize.RightToLeft(false),
		excelize.ShowFormulas(true),
		excelize.ShowGridLines(true),
		excelize.ShowRowColHeaders(true),
		excelize.ZoomScale(80),
	); err != nil {
		panic(err)
	}

	var zoomScale excelize.ZoomScale
	fmt.Println("Default:")
	fmt.Println("- zoomScale: 80")

	if err := xl.SetSheetViewOptions(sheet, 0, excelize.ZoomScale(500)); err != nil {
		panic(err)
	}

	if err := xl.GetSheetViewOptions(sheet, 0, &zoomScale); err != nil {
		panic(err)
	}

	fmt.Println("Used out of range value:")
	fmt.Println("- zoomScale:", zoomScale)

	if err := xl.SetSheetViewOptions(sheet, 0, excelize.ZoomScale(123)); err != nil {
		panic(err)
	}

	if err := xl.GetSheetViewOptions(sheet, 0, &zoomScale); err != nil {
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
	xl := excelize.NewFile()
	const sheet = "Sheet1"

	var (
		defaultGridColor  excelize.DefaultGridColor
		rightToLeft       excelize.RightToLeft
		showFormulas      excelize.ShowFormulas
		showGridLines     excelize.ShowGridLines
		showRowColHeaders excelize.ShowRowColHeaders
		zoomScale         excelize.ZoomScale
	)

	if err := xl.GetSheetViewOptions(sheet, 0,
		&defaultGridColor,
		&rightToLeft,
		&showFormulas,
		&showGridLines,
		&showRowColHeaders,
		&zoomScale,
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

	if err := xl.SetSheetViewOptions(sheet, 0, excelize.ShowGridLines(false)); err != nil {
		panic(err)
	}

	if err := xl.GetSheetViewOptions(sheet, 0, &showGridLines); err != nil {
		panic(err)
	}

	fmt.Println("After change:")
	fmt.Println("- showGridLines:", showGridLines)

	// Output:
	// Default:
	// - defaultGridColor: true
	// - rightToLeft: false
	// - showFormulas: false
	// - showGridLines: true
	// - showRowColHeaders: true
	// - zoomScale: 0
	// After change:
	// - showGridLines: false
}

func TestSheetViewOptionsErrors(t *testing.T) {
	xl := excelize.NewFile()
	const sheet = "Sheet1"

	if err := xl.GetSheetViewOptions(sheet, 0); err != nil {
		t.Errorf("Unexpected error: %s", err)
	}
	if err := xl.GetSheetViewOptions(sheet, -1); err != nil {
		t.Errorf("Unexpected error: %s", err)
	}
	if err := xl.GetSheetViewOptions(sheet, 1); err == nil {
		t.Error("Error expected but got nil")
	}
	if err := xl.GetSheetViewOptions(sheet, -2); err == nil {
		t.Error("Error expected but got nil")
	}

	if err := xl.SetSheetViewOptions(sheet, 0); err != nil {
		t.Errorf("Unexpected error: %s", err)
	}
	if err := xl.SetSheetViewOptions(sheet, -1); err != nil {
		t.Errorf("Unexpected error: %s", err)
	}
	if err := xl.SetSheetViewOptions(sheet, 1); err == nil {
		t.Error("Error expected but got nil")
	}
	if err := xl.SetSheetViewOptions(sheet, -2); err == nil {
		t.Error("Error expected but got nil")
	}
}
