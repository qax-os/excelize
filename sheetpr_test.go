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
		panic(err)
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
		panic(err)
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
