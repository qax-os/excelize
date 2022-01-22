package excelize

import (
	"fmt"
	"testing"

	"github.com/mohae/deepcopy"
	"github.com/stretchr/testify/assert"
)

var _ = []SheetPrOption{
	CodeName("hello"),
	EnableFormatConditionsCalculation(false),
	Published(false),
	FitToPage(true),
	TabColor("#FFFF00"),
	AutoPageBreaks(true),
	OutlineSummaryBelow(true),
}

var _ = []SheetPrOptionPtr{
	(*CodeName)(nil),
	(*EnableFormatConditionsCalculation)(nil),
	(*Published)(nil),
	(*FitToPage)(nil),
	(*TabColor)(nil),
	(*AutoPageBreaks)(nil),
	(*OutlineSummaryBelow)(nil),
}

func ExampleFile_SetSheetPrOptions() {
	f := NewFile()
	const sheet = "Sheet1"

	if err := f.SetSheetPrOptions(sheet,
		CodeName("code"),
		EnableFormatConditionsCalculation(false),
		Published(false),
		FitToPage(true),
		TabColor("#FFFF00"),
		AutoPageBreaks(true),
		OutlineSummaryBelow(false),
	); err != nil {
		fmt.Println(err)
	}
	// Output:
}

func ExampleFile_GetSheetPrOptions() {
	f := NewFile()
	const sheet = "Sheet1"

	var (
		codeName                          CodeName
		enableFormatConditionsCalculation EnableFormatConditionsCalculation
		published                         Published
		fitToPage                         FitToPage
		tabColor                          TabColor
		autoPageBreaks                    AutoPageBreaks
		outlineSummaryBelow               OutlineSummaryBelow
	)

	if err := f.GetSheetPrOptions(sheet,
		&codeName,
		&enableFormatConditionsCalculation,
		&published,
		&fitToPage,
		&tabColor,
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
	fmt.Printf("- tabColor: %q\n", tabColor)
	fmt.Println("- autoPageBreaks:", autoPageBreaks)
	fmt.Println("- outlineSummaryBelow:", outlineSummaryBelow)
	// Output:
	// Defaults:
	// - codeName: ""
	// - enableFormatConditionsCalculation: true
	// - published: true
	// - fitToPage: false
	// - tabColor: ""
	// - autoPageBreaks: false
	// - outlineSummaryBelow: true
}

func TestSheetPrOptions(t *testing.T) {
	const sheet = "Sheet1"

	testData := []struct {
		container  SheetPrOptionPtr
		nonDefault SheetPrOption
	}{
		{new(CodeName), CodeName("xx")},
		{new(EnableFormatConditionsCalculation), EnableFormatConditionsCalculation(false)},
		{new(Published), Published(false)},
		{new(FitToPage), FitToPage(true)},
		{new(TabColor), TabColor("FFFF00")},
		{new(AutoPageBreaks), AutoPageBreaks(true)},
		{new(OutlineSummaryBelow), OutlineSummaryBelow(false)},
	}

	for i, test := range testData {
		t.Run(fmt.Sprintf("TestData%d", i), func(t *testing.T) {
			opt := test.nonDefault
			t.Logf("option %T", opt)

			def := deepcopy.Copy(test.container).(SheetPrOptionPtr)
			val1 := deepcopy.Copy(def).(SheetPrOptionPtr)
			val2 := deepcopy.Copy(def).(SheetPrOptionPtr)

			f := NewFile()
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

func TestSetSheetPrOptions(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetSheetPrOptions("Sheet1", TabColor("")))
	// Test SetSheetPrOptions on not exists worksheet.
	assert.EqualError(t, f.SetSheetPrOptions("SheetN"), "sheet SheetN is not exist")
}

func TestGetSheetPrOptions(t *testing.T) {
	f := NewFile()
	// Test GetSheetPrOptions on not exists worksheet.
	assert.EqualError(t, f.GetSheetPrOptions("SheetN"), "sheet SheetN is not exist")
}

var _ = []PageMarginsOptions{
	PageMarginBottom(1.0),
	PageMarginFooter(1.0),
	PageMarginHeader(1.0),
	PageMarginLeft(1.0),
	PageMarginRight(1.0),
	PageMarginTop(1.0),
}

var _ = []PageMarginsOptionsPtr{
	(*PageMarginBottom)(nil),
	(*PageMarginFooter)(nil),
	(*PageMarginHeader)(nil),
	(*PageMarginLeft)(nil),
	(*PageMarginRight)(nil),
	(*PageMarginTop)(nil),
}

func ExampleFile_SetPageMargins() {
	f := NewFile()
	const sheet = "Sheet1"

	if err := f.SetPageMargins(sheet,
		PageMarginBottom(1.0),
		PageMarginFooter(1.0),
		PageMarginHeader(1.0),
		PageMarginLeft(1.0),
		PageMarginRight(1.0),
		PageMarginTop(1.0),
	); err != nil {
		fmt.Println(err)
	}
	// Output:
}

func ExampleFile_GetPageMargins() {
	f := NewFile()
	const sheet = "Sheet1"

	var (
		marginBottom PageMarginBottom
		marginFooter PageMarginFooter
		marginHeader PageMarginHeader
		marginLeft   PageMarginLeft
		marginRight  PageMarginRight
		marginTop    PageMarginTop
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
		container  PageMarginsOptionsPtr
		nonDefault PageMarginsOptions
	}{
		{new(PageMarginTop), PageMarginTop(1.0)},
		{new(PageMarginBottom), PageMarginBottom(1.0)},
		{new(PageMarginLeft), PageMarginLeft(1.0)},
		{new(PageMarginRight), PageMarginRight(1.0)},
		{new(PageMarginHeader), PageMarginHeader(1.0)},
		{new(PageMarginFooter), PageMarginFooter(1.0)},
	}

	for i, test := range testData {
		t.Run(fmt.Sprintf("TestData%d", i), func(t *testing.T) {
			opt := test.nonDefault
			t.Logf("option %T", opt)

			def := deepcopy.Copy(test.container).(PageMarginsOptionsPtr)
			val1 := deepcopy.Copy(def).(PageMarginsOptionsPtr)
			val2 := deepcopy.Copy(def).(PageMarginsOptionsPtr)

			f := NewFile()
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
	f := NewFile()
	// Test set page margins on not exists worksheet.
	assert.EqualError(t, f.SetPageMargins("SheetN"), "sheet SheetN is not exist")
}

func TestGetPageMargins(t *testing.T) {
	f := NewFile()
	// Test get page margins on not exists worksheet.
	assert.EqualError(t, f.GetPageMargins("SheetN"), "sheet SheetN is not exist")
}

func ExampleFile_SetSheetFormatPr() {
	f := NewFile()
	const sheet = "Sheet1"

	if err := f.SetSheetFormatPr(sheet,
		BaseColWidth(1.0),
		DefaultColWidth(1.0),
		DefaultRowHeight(1.0),
		CustomHeight(true),
		ZeroHeight(true),
		ThickTop(true),
		ThickBottom(true),
	); err != nil {
		fmt.Println(err)
	}
	// Output:
}

func ExampleFile_GetSheetFormatPr() {
	f := NewFile()
	const sheet = "Sheet1"

	var (
		baseColWidth     BaseColWidth
		defaultColWidth  DefaultColWidth
		defaultRowHeight DefaultRowHeight
		customHeight     CustomHeight
		zeroHeight       ZeroHeight
		thickTop         ThickTop
		thickBottom      ThickBottom
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
		container  SheetFormatPrOptionsPtr
		nonDefault SheetFormatPrOptions
	}{
		{new(BaseColWidth), BaseColWidth(1.0)},
		{new(DefaultColWidth), DefaultColWidth(1.0)},
		{new(DefaultRowHeight), DefaultRowHeight(1.0)},
		{new(CustomHeight), CustomHeight(true)},
		{new(ZeroHeight), ZeroHeight(true)},
		{new(ThickTop), ThickTop(true)},
		{new(ThickBottom), ThickBottom(true)},
	}

	for i, test := range testData {
		t.Run(fmt.Sprintf("TestData%d", i), func(t *testing.T) {
			opt := test.nonDefault
			t.Logf("option %T", opt)

			def := deepcopy.Copy(test.container).(SheetFormatPrOptionsPtr)
			val1 := deepcopy.Copy(def).(SheetFormatPrOptionsPtr)
			val2 := deepcopy.Copy(def).(SheetFormatPrOptionsPtr)

			f := NewFile()
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
	f := NewFile()
	assert.NoError(t, f.GetSheetFormatPr("Sheet1"))
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).SheetFormatPr = nil
	assert.NoError(t, f.SetSheetFormatPr("Sheet1", BaseColWidth(1.0)))
	// Test set formatting properties on not exists worksheet.
	assert.EqualError(t, f.SetSheetFormatPr("SheetN"), "sheet SheetN is not exist")
}

func TestGetSheetFormatPr(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.GetSheetFormatPr("Sheet1"))
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).SheetFormatPr = nil
	var (
		baseColWidth     BaseColWidth
		defaultColWidth  DefaultColWidth
		defaultRowHeight DefaultRowHeight
		customHeight     CustomHeight
		zeroHeight       ZeroHeight
		thickTop         ThickTop
		thickBottom      ThickBottom
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
