package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

func ExampleFile_SetWorkbookPrOptions() {
	f := NewFile()
	if err := f.SetWorkbookPrOptions(
		Date1904(false),
		FilterPrivacy(false),
		CodeName("code"),
	); err != nil {
		fmt.Println(err)
	}
	// Output:
}

func ExampleFile_GetWorkbookPrOptions() {
	f := NewFile()
	var (
		date1904      Date1904
		filterPrivacy FilterPrivacy
		codeName      CodeName
	)
	if err := f.GetWorkbookPrOptions(&date1904); err != nil {
		fmt.Println(err)
	}
	if err := f.GetWorkbookPrOptions(&filterPrivacy); err != nil {
		fmt.Println(err)
	}
	if err := f.GetWorkbookPrOptions(&codeName); err != nil {
		fmt.Println(err)
	}
	fmt.Println("Defaults:")
	fmt.Printf("- date1904: %t\n", date1904)
	fmt.Printf("- filterPrivacy: %t\n", filterPrivacy)
	fmt.Printf("- codeName: %q\n", codeName)
	// Output:
	// Defaults:
	// - date1904: false
	// - filterPrivacy: true
	// - codeName: ""
}

func TestWorkbookPr(t *testing.T) {
	f := NewFile()
	wb := f.workbookReader()
	wb.WorkbookPr = nil
	var date1904 Date1904
	assert.NoError(t, f.GetWorkbookPrOptions(&date1904))
	assert.Equal(t, false, bool(date1904))

	wb.WorkbookPr = nil
	var codeName CodeName
	assert.NoError(t, f.GetWorkbookPrOptions(&codeName))
	assert.Equal(t, "", string(codeName))
	assert.NoError(t, f.SetWorkbookPrOptions(CodeName("code")))
	assert.NoError(t, f.GetWorkbookPrOptions(&codeName))
	assert.Equal(t, "code", string(codeName))

	wb.WorkbookPr = nil
	var filterPrivacy FilterPrivacy
	assert.NoError(t, f.GetWorkbookPrOptions(&filterPrivacy))
	assert.Equal(t, false, bool(filterPrivacy))
}
