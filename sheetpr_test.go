package excelize_test

import (
	"fmt"
	"reflect"
	"testing"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/mohae/deepcopy"
)

var _ = []excelize.SheetPrOption{
	excelize.CodeName("hello"),
	excelize.EnableFormatConditionsCalculation(false),
	excelize.Published(false),
	excelize.FitToPage(true),
	excelize.AutoPageBreaks(true),
}

var _ = []excelize.SheetPrOptionPtr{
	(*excelize.CodeName)(nil),
	(*excelize.EnableFormatConditionsCalculation)(nil),
	(*excelize.Published)(nil),
	(*excelize.FitToPage)(nil),
	(*excelize.AutoPageBreaks)(nil),
}

func ExampleFile_SetSheetPrOptions() {
	xl := excelize.NewFile()
	const sheet = "Sheet1"

	if err := xl.SetSheetPrOptions(sheet,
		excelize.CodeName("code"),
		excelize.EnableFormatConditionsCalculation(false),
		excelize.Published(false),
		excelize.FitToPage(true),
		excelize.AutoPageBreaks(true),
	); err != nil {
		panic(err)
	}
	// Output:
}

func ExampleFile_GetSheetPrOptions() {
	xl := excelize.NewFile()
	const sheet = "Sheet1"

	var (
		codeName                          excelize.CodeName
		enableFormatConditionsCalculation excelize.EnableFormatConditionsCalculation
		published                         excelize.Published
		fitToPage                         excelize.FitToPage
		autoPageBreaks                    excelize.AutoPageBreaks
	)

	if err := xl.GetSheetPrOptions(sheet,
		&codeName,
		&enableFormatConditionsCalculation,
		&published,
		&fitToPage,
		&autoPageBreaks,
	); err != nil {
		panic(err)
	}
	fmt.Println("Defaults:")
	fmt.Printf("- codeName: %q\n", codeName)
	fmt.Println("- enableFormatConditionsCalculation:", enableFormatConditionsCalculation)
	fmt.Println("- published:", published)
	fmt.Println("- fitToPage:", fitToPage)
	fmt.Println("- autoPageBreaks:", autoPageBreaks)
	// Output:
	// Defaults:
	// - codeName: ""
	// - enableFormatConditionsCalculation: true
	// - published: true
	// - fitToPage: false
	// - autoPageBreaks: false
}

func TestSheetPrOptions(t *testing.T) {
	const sheet = "Sheet1"
	for _, test := range []struct {
		container  excelize.SheetPrOptionPtr
		nonDefault excelize.SheetPrOption
	}{
		{new(excelize.CodeName), excelize.CodeName("xx")},
		{new(excelize.EnableFormatConditionsCalculation), excelize.EnableFormatConditionsCalculation(false)},
		{new(excelize.Published), excelize.Published(false)},
		{new(excelize.FitToPage), excelize.FitToPage(true)},
		{new(excelize.AutoPageBreaks), excelize.AutoPageBreaks(true)},
	} {
		opt := test.nonDefault
		t.Logf("option %T", opt)

		def := deepcopy.Copy(test.container).(excelize.SheetPrOptionPtr)
		val1 := deepcopy.Copy(def).(excelize.SheetPrOptionPtr)
		val2 := deepcopy.Copy(def).(excelize.SheetPrOptionPtr)

		xl := excelize.NewFile()
		// Get the default value
		if err := xl.GetSheetPrOptions(sheet, def); err != nil {
			t.Fatalf("%T: %s", opt, err)
		}
		// Get again and check
		if err := xl.GetSheetPrOptions(sheet, val1); err != nil {
			t.Fatalf("%T: %s", opt, err)
		}
		if !reflect.DeepEqual(val1, def) {
			t.Fatalf("%T: value should not have changed", opt)
		}
		// Set the same value
		if err := xl.SetSheetPrOptions(sheet, val1); err != nil {
			t.Fatalf("%T: %s", opt, err)
		}
		// Get again and check
		if err := xl.GetSheetPrOptions(sheet, val1); err != nil {
			t.Fatalf("%T: %s", opt, err)
		}
		if !reflect.DeepEqual(val1, def) {
			t.Fatalf("%T: value should not have changed", opt)
		}

		// Set a different value
		if err := xl.SetSheetPrOptions(sheet, test.nonDefault); err != nil {
			t.Fatalf("%T: %s", opt, err)
		}
		if err := xl.GetSheetPrOptions(sheet, val1); err != nil {
			t.Fatalf("%T: %s", opt, err)
		}
		// Get again and compare
		if err := xl.GetSheetPrOptions(sheet, val2); err != nil {
			t.Fatalf("%T: %s", opt, err)
		}
		if !reflect.DeepEqual(val2, val1) {
			t.Fatalf("%T: value should not have changed", opt)
		}
		// Value should not be the same as the default
		if reflect.DeepEqual(val1, def) {
			t.Fatalf("%T: value should have changed from default", opt)
		}

		// Restore the default value
		if err := xl.SetSheetPrOptions(sheet, def); err != nil {
			t.Fatalf("%T: %s", opt, err)
		}
		if err := xl.GetSheetPrOptions(sheet, val1); err != nil {
			t.Fatalf("%T: %s", opt, err)
		}
		if !reflect.DeepEqual(val1, def) {
			t.Fatalf("%T: value should now be the same as default", opt)
		}
	}
}
