package excelize_test

import (
	"fmt"
	"path/filepath"
	"strings"
	"testing"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/mohae/deepcopy"
	"github.com/stretchr/testify/assert"
)

func ExampleFile_SetPageLayout() {
	f := excelize.NewFile()

	if err := f.SetPageLayout(
		"Sheet1",
		excelize.PageLayoutOrientation(excelize.OrientationLandscape),
	); err != nil {
		panic(err)
	}
	if err := f.SetPageLayout(
		"Sheet1",
		excelize.PageLayoutPaperSize(10),
	); err != nil {
		panic(err)
	}
	// Output:
}

func ExampleFile_GetPageLayout() {
	f := excelize.NewFile()
	var (
		orientation excelize.PageLayoutOrientation
		paperSize   excelize.PageLayoutPaperSize
	)
	if err := f.GetPageLayout("Sheet1", &orientation); err != nil {
		panic(err)
	}
	if err := f.GetPageLayout("Sheet1", &paperSize); err != nil {
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

			f := excelize.NewFile()
			// Get the default value
			assert.NoError(t, f.GetPageLayout(sheet, def), opt)
			// Get again and check
			assert.NoError(t, f.GetPageLayout(sheet, val1), opt)
			if !assert.Equal(t, val1, def, opt) {
				t.FailNow()
			}
			// Set the same value
			assert.NoError(t, f.SetPageLayout(sheet, val1), opt)
			// Get again and check
			assert.NoError(t, f.GetPageLayout(sheet, val1), opt)
			if !assert.Equal(t, val1, def, "%T: value should not have changed", opt) {
				t.FailNow()
			}
			// Set a different value
			assert.NoError(t, f.SetPageLayout(sheet, test.nonDefault), opt)
			assert.NoError(t, f.GetPageLayout(sheet, val1), opt)
			// Get again and compare
			assert.NoError(t, f.GetPageLayout(sheet, val2), opt)
			if !assert.Equal(t, val1, val2, "%T: value should not have changed", opt) {
				t.FailNow()
			}
			// Value should not be the same as the default
			if !assert.NotEqual(t, def, val1, "%T: value should have changed from default", opt) {
				t.FailNow()
			}
			// Restore the default value
			assert.NoError(t, f.SetPageLayout(sheet, def), opt)
			assert.NoError(t, f.GetPageLayout(sheet, val1), opt)
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

func TestSetHeaderFooter(t *testing.T) {
	f := excelize.NewFile()
	f.SetCellStr("Sheet1", "A1", "Test SetHeaderFooter")
	// Test set header and footer on not exists worksheet.
	assert.EqualError(t, f.SetHeaderFooter("SheetN", nil), "sheet SheetN is not exist")
	// Test set header and footer with illegal setting.
	assert.EqualError(t, f.SetHeaderFooter("Sheet1", &excelize.FormatHeaderFooter{
		OddHeader: strings.Repeat("c", 256),
	}), "field OddHeader must be less than 255 characters")

	assert.NoError(t, f.SetHeaderFooter("Sheet1", nil))
	assert.NoError(t, f.SetHeaderFooter("Sheet1", &excelize.FormatHeaderFooter{
		DifferentFirst:   true,
		DifferentOddEven: true,
		OddHeader:        "&R&P",
		OddFooter:        "&C&F",
		EvenHeader:       "&L&P",
		EvenFooter:       "&L&D&R&T",
		FirstHeader:      `&CCenter &"-,Bold"Bold&"-,Regular"HeaderU+000A&D`,
	}))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetHeaderFooter.xlsx")))
}
