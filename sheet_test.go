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
			assert.NoError(t, xl.GetPageLayout(sheet, def), opt)
			// Get again and check
			assert.NoError(t, xl.GetPageLayout(sheet, val1), opt)
			if !assert.Equal(t, val1, def, opt) {
				t.FailNow()
			}
			// Set the same value
			assert.NoError(t, xl.SetPageLayout(sheet, val1), opt)
			// Get again and check
			assert.NoError(t, xl.GetPageLayout(sheet, val1), opt)
			if !assert.Equal(t, val1, def, "%T: value should not have changed", opt) {
				t.FailNow()
			}
			// Set a different value
			assert.NoError(t, xl.SetPageLayout(sheet, test.nonDefault), opt)
			assert.NoError(t, xl.GetPageLayout(sheet, val1), opt)
			// Get again and compare
			assert.NoError(t, xl.GetPageLayout(sheet, val2), opt)
			if !assert.Equal(t, val1, val2, "%T: value should not have changed", opt) {
				t.FailNow()
			}
			// Value should not be the same as the default
			if !assert.NotEqual(t, def, val1, "%T: value should have changed from default", opt) {
				t.FailNow()
			}
			// Restore the default value
			assert.NoError(t, xl.SetPageLayout(sheet, def), opt)
			assert.NoError(t, xl.GetPageLayout(sheet, val1), opt)
			if !assert.Equal(t, def, val1) {
				t.FailNow()
			}
		})
	}
}

func TestSetPageLayout(t *testing.T) {
	f := excelize.NewFile()
	// Test set page layout on not exists worksheet.
	assert.EqualError(t, f.SetPageLayout("SheetN"), "sheet SheetN is not exist")
}

func TestGetPageLayout(t *testing.T) {
	f := excelize.NewFile()
	// Test get page layout on not exists worksheet.
	assert.EqualError(t, f.GetPageLayout("SheetN"), "sheet SheetN is not exist")
}
