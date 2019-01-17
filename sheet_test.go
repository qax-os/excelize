package excelize_test

import (
	"fmt"
	"testing"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/mohae/deepcopy"
	"github.com/stretchr/testify/assert"
)

func ExampleFile_SetPageLayout() {
	xl := excelize.NewFile()
	const sheet = "Sheet1"

	if err := xl.SetPageLayout(
		"Sheet1",
		excelize.PageLayoutOrientation(excelize.OrientationLandscape),
	); err != nil {
		panic(err)
	}
	if err := xl.SetPageLayout(
		"Sheet1",
		excelize.PageLayoutPaperSize(10),
	); err != nil {
		panic(err)
	}
	// Output:
}

func ExampleFile_GetPageLayout() {
	xl := excelize.NewFile()
	const sheet = "Sheet1"
	var (
		orientation excelize.PageLayoutOrientation
		paperSize   excelize.PageLayoutPaperSize
	)
	if err := xl.GetPageLayout("Sheet1", &orientation); err != nil {
		panic(err)
	}
	if err := xl.GetPageLayout("Sheet1", &paperSize); err != nil {
		panic(err)
	}
	fmt.Println("Defaults:")
	fmt.Printf("- orientation: %q\n", orientation)
	fmt.Printf("- paper size: %d\n", paperSize)
	// Output:
	// Defaults:
	// - orientation: "portrait"
	// - paper size: 1
}

func TestPageLayoutOption(t *testing.T) {
	const sheet = "Sheet1"

	testData := []struct {
		container  excelize.PageLayoutOptionPtr
		nonDefault excelize.PageLayoutOption
	}{
		{new(excelize.PageLayoutOrientation), excelize.PageLayoutOrientation(excelize.OrientationLandscape)},
		{new(excelize.PageLayoutPaperSize), excelize.PageLayoutPaperSize(10)},
	}

	for i, test := range testData {
		t.Run(fmt.Sprintf("TestData%d", i), func(t *testing.T) {

			opt := test.nonDefault
			t.Logf("option %T", opt)

			def := deepcopy.Copy(test.container).(excelize.PageLayoutOptionPtr)
			val1 := deepcopy.Copy(def).(excelize.PageLayoutOptionPtr)
			val2 := deepcopy.Copy(def).(excelize.PageLayoutOptionPtr)

			xl := excelize.NewFile()
			// Get the default value
			if !assert.NoError(t, xl.GetPageLayout(sheet, def), opt) {
				t.FailNow()
			}
			// Get again and check
			if !assert.NoError(t, xl.GetPageLayout(sheet, val1), opt) {
				t.FailNow()
			}
			if !assert.Equal(t, val1, def, opt) {
				t.FailNow()
			}
			// Set the same value
			if !assert.NoError(t, xl.SetPageLayout(sheet, val1), opt) {
				t.FailNow()
			}
			// Get again and check
			if !assert.NoError(t, xl.GetPageLayout(sheet, val1), opt) {
				t.FailNow()
			}
			if !assert.Equal(t, val1, def, "%T: value should not have changed", opt) {
				t.FailNow()
			}
			// Set a different value
			if !assert.NoError(t, xl.SetPageLayout(sheet, test.nonDefault), opt) {
				t.FailNow()
			}
			if !assert.NoError(t, xl.GetPageLayout(sheet, val1), opt) {
				t.FailNow()
			}
			// Get again and compare
			if !assert.NoError(t, xl.GetPageLayout(sheet, val2), opt) {
				t.FailNow()
			}
			if !assert.Equal(t, val1, val2, "%T: value should not have changed", opt) {
				t.FailNow()
			}
			// Value should not be the same as the default
			if !assert.NotEqual(t, def, val1, "%T: value should have changed from default", opt) {
				t.FailNow()
			}
			// Restore the default value
			if !assert.NoError(t, xl.SetPageLayout(sheet, def), opt) {
				t.FailNow()
			}
			if !assert.NoError(t, xl.GetPageLayout(sheet, val1), opt) {
				t.FailNow()
			}
			if !assert.Equal(t, def, val1) {
				t.FailNow()
			}
		})
	}
}
