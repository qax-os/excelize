package excelize_test

import (
	"fmt"
	"testing"

	"github.com/mohae/deepcopy"
	"github.com/stretchr/testify/assert"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

var _ = []excelize.SheetPrOption{
	excelize.CodeName("hello"),
	excelize.EnableFormatConditionsCalculation(false),
	excelize.Published(false),
	excelize.FitToPage(true),
	excelize.AutoPageBreaks(true),
	excelize.OutlineSummaryBelow(true),
}

var _ = []excelize.SheetPrOptionPtr{
	(*excelize.CodeName)(nil),
	(*excelize.EnableFormatConditionsCalculation)(nil),
	(*excelize.Published)(nil),
	(*excelize.FitToPage)(nil),
	(*excelize.AutoPageBreaks)(nil),
	(*excelize.OutlineSummaryBelow)(nil),
}

func ExampleFile_SetSheetPrOptions() {
	f := excelize.NewFile()
	const sheet = "Sheet1"

	if err := f.SetSheetPrOptions(sheet,
		excelize.CodeName("code"),
		excelize.EnableFormatConditionsCalculation(false),
		excelize.Published(false),
		excelize.FitToPage(true),
		excelize.AutoPageBreaks(true),
		excelize.OutlineSummaryBelow(false),
	); err != nil {
		fmt.Println(err)
	}
	// Output:
}

func ExampleFile_GetSheetPrOptions() {
	f := excelize.NewFile()
	const sheet = "Sheet1"

	var (
		codeName                          excelize.CodeName
		enableFormatConditionsCalculation excelize.EnableFormatConditionsCalculation
		published                         excelize.Published
		fitToPage                         excelize.FitToPage
		autoPageBreaks                    excelize.AutoPageBreaks
		outlineSummaryBelow               excelize.OutlineSummaryBelow
	)

	if err := f.GetSheetPrOptions(sheet,
		&codeName,
		&enableFormatConditionsCalculation,
		&published,
		&fitToPage,
		&autoPageBreaks,
		&outlineSummaryBelow,
	); err != nil {
		fmt.Println(err)
	}
	fmt.Println("Defaults:")
	fmt.Printf("- codeName: %q\n", codeName)
	fmt.Println("- enableFormatConditionsCalculation:", enableFormatConditionsCalculation)
	fmt.Println("- published:", published)
	fmt.Println("- fitToPage:", fitToPage)
	fmt.Println("- autoPageBreaks:", autoPageBreaks)
	fmt.Println("- outlineSummaryBelow:", outlineSummaryBelow)
	// Output:
	// Defaults:
	// - codeName: ""
	// - enableFormatConditionsCalculation: true
	// - published: true
	// - fitToPage: false
	// - autoPageBreaks: false
	// - outlineSummaryBelow: true
}

func TestSheetPrOptions(t *testing.T) {
	const sheet = "Sheet1"

	testData := []struct {
		container  excelize.SheetPrOptionPtr
		nonDefault excelize.SheetPrOption
	}{
		{new(excelize.CodeName), excelize.CodeName("xx")},
		{new(excelize.EnableFormatConditionsCalculation), excelize.EnableFormatConditionsCalculation(false)},
		{new(excelize.Published), excelize.Published(false)},
		{new(excelize.FitToPage), excelize.FitToPage(true)},
		{new(excelize.AutoPageBreaks), excelize.AutoPageBreaks(true)},
		{new(excelize.OutlineSummaryBelow), excelize.OutlineSummaryBelow(false)},
	}

	for i, test := range testData {
		t.Run(fmt.Sprintf("TestData%d", i), func(t *testing.T) {

			opt := test.nonDefault
			t.Logf("option %T", opt)

			def := deepcopy.Copy(test.container).(excelize.SheetPrOptionPtr)
			val1 := deepcopy.Copy(def).(excelize.SheetPrOptionPtr)
			val2 := deepcopy.Copy(def).(excelize.SheetPrOptionPtr)

			f := excelize.NewFile()
			// Get the default value
			assert.NoError(t, f.GetSheetPrOptions(sheet, def), opt)
			// Get again and check
			assert.NoError(t, f.GetSheetPrOptions(sheet, val1), opt)
			if !assert.Equal(t, val1, def, opt) {
				t.FailNow()
			}
			// Set the same value
			assert.NoError(t, f.SetSheetPrOptions(sheet, val1), opt)
			// Get again and check
			assert.NoError(t, f.GetSheetPrOptions(sheet, val1), opt)
			if !assert.Equal(t, val1, def, "%T: value should not have changed", opt) {
				t.FailNow()
			}
			// Set a different value
			assert.NoError(t, f.SetSheetPrOptions(sheet, test.nonDefault), opt)
			assert.NoError(t, f.GetSheetPrOptions(sheet, val1), opt)
			// Get again and compare
			assert.NoError(t, f.GetSheetPrOptions(sheet, val2), opt)
			if !assert.Equal(t, val1, val2, "%T: value should not have changed", opt) {
				t.FailNow()
			}
			// Value should not be the same as the default
			if !assert.NotEqual(t, def, val1, "%T: value should have changed from default", opt) {
				t.FailNow()
			}
			// Restore the default value
			assert.NoError(t, f.SetSheetPrOptions(sheet, def), opt)
			assert.NoError(t, f.GetSheetPrOptions(sheet, val1), opt)
			if !assert.Equal(t, def, val1) {
				t.FailNow()
			}
		})
	}
}

func TestSetSheetrOptions(t *testing.T) {
	f := excelize.NewFile()
	// Test SetSheetrOptions on not exists worksheet.
	assert.EqualError(t, f.SetSheetPrOptions("SheetN"), "sheet SheetN is not exist")
}

func TestGetSheetPrOptions(t *testing.T) {
	f := excelize.NewFile()
	// Test GetSheetPrOptions on not exists worksheet.
	assert.EqualError(t, f.GetSheetPrOptions("SheetN"), "sheet SheetN is not exist")
}

var _ = []excelize.PageMarginsOptions{
	excelize.PageMarginBottom(1.0),
	excelize.PageMarginFooter(1.0),
	excelize.PageMarginHeader(1.0),
	excelize.PageMarginLeft(1.0),
	excelize.PageMarginRight(1.0),
	excelize.PageMarginTop(1.0),
}

var _ = []excelize.PageMarginsOptionsPtr{
	(*excelize.PageMarginBottom)(nil),
	(*excelize.PageMarginFooter)(nil),
	(*excelize.PageMarginHeader)(nil),
	(*excelize.PageMarginLeft)(nil),
	(*excelize.PageMarginRight)(nil),
	(*excelize.PageMarginTop)(nil),
}

func ExampleFile_SetPageMargins() {
	f := excelize.NewFile()
	const sheet = "Sheet1"

	if err := f.SetPageMargins(sheet,
		excelize.PageMarginBottom(1.0),
		excelize.PageMarginFooter(1.0),
		excelize.PageMarginHeader(1.0),
		excelize.PageMarginLeft(1.0),
		excelize.PageMarginRight(1.0),
		excelize.PageMarginTop(1.0),
	); err != nil {
		fmt.Println(err)
	}
	// Output:
}

func ExampleFile_GetPageMargins() {
	f := excelize.NewFile()
	const sheet = "Sheet1"

	var (
		marginBottom excelize.PageMarginBottom
		marginFooter excelize.PageMarginFooter
		marginHeader excelize.PageMarginHeader
		marginLeft   excelize.PageMarginLeft
		marginRight  excelize.PageMarginRight
		marginTop    excelize.PageMarginTop
	)

	if err := f.GetPageMargins(sheet,
		&marginBottom,
		&marginFooter,
		&marginHeader,
		&marginLeft,
		&marginRight,
		&marginTop,
	); err != nil {
		fmt.Println(err)
	}
	fmt.Println("Defaults:")
	fmt.Println("- marginBottom:", marginBottom)
	fmt.Println("- marginFooter:", marginFooter)
	fmt.Println("- marginHeader:", marginHeader)
	fmt.Println("- marginLeft:", marginLeft)
	fmt.Println("- marginRight:", marginRight)
	fmt.Println("- marginTop:", marginTop)
	// Output:
	// Defaults:
	// - marginBottom: 0.75
	// - marginFooter: 0.3
	// - marginHeader: 0.3
	// - marginLeft: 0.7
	// - marginRight: 0.7
	// - marginTop: 0.75
}

func TestPageMarginsOption(t *testing.T) {
	const sheet = "Sheet1"

	testData := []struct {
		container  excelize.PageMarginsOptionsPtr
		nonDefault excelize.PageMarginsOptions
	}{
		{new(excelize.PageMarginTop), excelize.PageMarginTop(1.0)},
		{new(excelize.PageMarginBottom), excelize.PageMarginBottom(1.0)},
		{new(excelize.PageMarginLeft), excelize.PageMarginLeft(1.0)},
		{new(excelize.PageMarginRight), excelize.PageMarginRight(1.0)},
		{new(excelize.PageMarginHeader), excelize.PageMarginHeader(1.0)},
		{new(excelize.PageMarginFooter), excelize.PageMarginFooter(1.0)},
	}

	for i, test := range testData {
		t.Run(fmt.Sprintf("TestData%d", i), func(t *testing.T) {

			opt := test.nonDefault
			t.Logf("option %T", opt)

			def := deepcopy.Copy(test.container).(excelize.PageMarginsOptionsPtr)
			val1 := deepcopy.Copy(def).(excelize.PageMarginsOptionsPtr)
			val2 := deepcopy.Copy(def).(excelize.PageMarginsOptionsPtr)

			f := excelize.NewFile()
			// Get the default value
			assert.NoError(t, f.GetPageMargins(sheet, def), opt)
			// Get again and check
			assert.NoError(t, f.GetPageMargins(sheet, val1), opt)
			if !assert.Equal(t, val1, def, opt) {
				t.FailNow()
			}
			// Set the same value
			assert.NoError(t, f.SetPageMargins(sheet, val1), opt)
			// Get again and check
			assert.NoError(t, f.GetPageMargins(sheet, val1), opt)
			if !assert.Equal(t, val1, def, "%T: value should not have changed", opt) {
				t.FailNow()
			}
			// Set a different value
			assert.NoError(t, f.SetPageMargins(sheet, test.nonDefault), opt)
			assert.NoError(t, f.GetPageMargins(sheet, val1), opt)
			// Get again and compare
			assert.NoError(t, f.GetPageMargins(sheet, val2), opt)
			if !assert.Equal(t, val1, val2, "%T: value should not have changed", opt) {
				t.FailNow()
			}
			// Value should not be the same as the default
			if !assert.NotEqual(t, def, val1, "%T: value should have changed from default", opt) {
				t.FailNow()
			}
			// Restore the default value
			assert.NoError(t, f.SetPageMargins(sheet, def), opt)
			assert.NoError(t, f.GetPageMargins(sheet, val1), opt)
			if !assert.Equal(t, def, val1) {
				t.FailNow()
			}
		})
	}
}

func TestSetPageMargins(t *testing.T) {
	f := excelize.NewFile()
	// Test set page margins on not exists worksheet.
	assert.EqualError(t, f.SetPageMargins("SheetN"), "sheet SheetN is not exist")
}

func TestGetPageMargins(t *testing.T) {
	f := excelize.NewFile()
	// Test get page margins on not exists worksheet.
	assert.EqualError(t, f.GetPageMargins("SheetN"), "sheet SheetN is not exist")
}

func ExampleFile_SetSheetFormatPr() {
	f := excelize.NewFile()
	const sheet = "Sheet1"

	if err := f.SetSheetFormatPr(sheet,
		excelize.BaseColWidth(1.0),
		excelize.DefaultColWidth(1.0),
		excelize.DefaultRowHeight(1.0),
		excelize.CustomHeight(true),
		excelize.ZeroHeight(true),
		excelize.ThickTop(true),
		excelize.ThickBottom(true),
	); err != nil {
		fmt.Println(err)
	}
	// Output:
}

func ExampleFile_GetSheetFormatPr() {
	f := excelize.NewFile()
	const sheet = "Sheet1"

	var (
		baseColWidth     excelize.BaseColWidth
		defaultColWidth  excelize.DefaultColWidth
		defaultRowHeight excelize.DefaultRowHeight
		customHeight     excelize.CustomHeight
		zeroHeight       excelize.ZeroHeight
		thickTop         excelize.ThickTop
		thickBottom      excelize.ThickBottom
	)

	if err := f.GetSheetFormatPr(sheet,
		&baseColWidth,
		&defaultColWidth,
		&defaultRowHeight,
		&customHeight,
		&zeroHeight,
		&thickTop,
		&thickBottom,
	); err != nil {
		fmt.Println(err)
	}
	fmt.Println("Defaults:")
	fmt.Println("- baseColWidth:", baseColWidth)
	fmt.Println("- defaultColWidth:", defaultColWidth)
	fmt.Println("- defaultRowHeight:", defaultRowHeight)
	fmt.Println("- customHeight:", customHeight)
	fmt.Println("- zeroHeight:", zeroHeight)
	fmt.Println("- thickTop:", thickTop)
	fmt.Println("- thickBottom:", thickBottom)
	// Output:
	// Defaults:
	// - baseColWidth: 0
	// - defaultColWidth: 0
	// - defaultRowHeight: 15
	// - customHeight: false
	// - zeroHeight: false
	// - thickTop: false
	// - thickBottom: false
}

func TestSheetFormatPrOptions(t *testing.T) {
	const sheet = "Sheet1"

	testData := []struct {
		container  excelize.SheetFormatPrOptionsPtr
		nonDefault excelize.SheetFormatPrOptions
	}{
		{new(excelize.BaseColWidth), excelize.BaseColWidth(1.0)},
		{new(excelize.DefaultColWidth), excelize.DefaultColWidth(1.0)},
		{new(excelize.DefaultRowHeight), excelize.DefaultRowHeight(1.0)},
		{new(excelize.CustomHeight), excelize.CustomHeight(true)},
		{new(excelize.ZeroHeight), excelize.ZeroHeight(true)},
		{new(excelize.ThickTop), excelize.ThickTop(true)},
		{new(excelize.ThickBottom), excelize.ThickBottom(true)},
	}

	for i, test := range testData {
		t.Run(fmt.Sprintf("TestData%d", i), func(t *testing.T) {

			opt := test.nonDefault
			t.Logf("option %T", opt)

			def := deepcopy.Copy(test.container).(excelize.SheetFormatPrOptionsPtr)
			val1 := deepcopy.Copy(def).(excelize.SheetFormatPrOptionsPtr)
			val2 := deepcopy.Copy(def).(excelize.SheetFormatPrOptionsPtr)

			f := excelize.NewFile()
			// Get the default value
			assert.NoError(t, f.GetSheetFormatPr(sheet, def), opt)
			// Get again and check
			assert.NoError(t, f.GetSheetFormatPr(sheet, val1), opt)
			if !assert.Equal(t, val1, def, opt) {
				t.FailNow()
			}
			// Set the same value
			assert.NoError(t, f.SetSheetFormatPr(sheet, val1), opt)
			// Get again and check
			assert.NoError(t, f.GetSheetFormatPr(sheet, val1), opt)
			if !assert.Equal(t, val1, def, "%T: value should not have changed", opt) {
				t.FailNow()
			}
			// Set a different value
			assert.NoError(t, f.SetSheetFormatPr(sheet, test.nonDefault), opt)
			assert.NoError(t, f.GetSheetFormatPr(sheet, val1), opt)
			// Get again and compare
			assert.NoError(t, f.GetSheetFormatPr(sheet, val2), opt)
			if !assert.Equal(t, val1, val2, "%T: value should not have changed", opt) {
				t.FailNow()
			}
			// Value should not be the same as the default
			if !assert.NotEqual(t, def, val1, "%T: value should have changed from default", opt) {
				t.FailNow()
			}
			// Restore the default value
			assert.NoError(t, f.SetSheetFormatPr(sheet, def), opt)
			assert.NoError(t, f.GetSheetFormatPr(sheet, val1), opt)
			if !assert.Equal(t, def, val1) {
				t.FailNow()
			}
		})
	}
}

func TestSetSheetFormatPr(t *testing.T) {
	f := excelize.NewFile()
	assert.NoError(t, f.GetSheetFormatPr("Sheet1"))
	f.Sheet["xl/worksheets/sheet1.xml"].SheetFormatPr = nil
	assert.NoError(t, f.SetSheetFormatPr("Sheet1", excelize.BaseColWidth(1.0)))
	// Test set formatting properties on not exists worksheet.
	assert.EqualError(t, f.SetSheetFormatPr("SheetN"), "sheet SheetN is not exist")
}

func TestGetSheetFormatPr(t *testing.T) {
	f := excelize.NewFile()
	assert.NoError(t, f.GetSheetFormatPr("Sheet1"))
	f.Sheet["xl/worksheets/sheet1.xml"].SheetFormatPr = nil
	var (
		baseColWidth     excelize.BaseColWidth
		defaultColWidth  excelize.DefaultColWidth
		defaultRowHeight excelize.DefaultRowHeight
		customHeight     excelize.CustomHeight
		zeroHeight       excelize.ZeroHeight
		thickTop         excelize.ThickTop
		thickBottom      excelize.ThickBottom
	)
	assert.NoError(t, f.GetSheetFormatPr("Sheet1",
		&baseColWidth,
		&defaultColWidth,
		&defaultRowHeight,
		&customHeight,
		&zeroHeight,
		&thickTop,
		&thickBottom,
	))
	// Test get formatting properties on not exists worksheet.
	assert.EqualError(t, f.GetSheetFormatPr("SheetN"), "sheet SheetN is not exist")
}
