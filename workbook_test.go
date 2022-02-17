package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

func ExampleFile_SetWorkbookPrOptions() {
	f := NewFile()
	if err := f.SetWorkbookPrOptions(
		CodeName("code"),
		FilterPrivacy(false),
	); err != nil {
		fmt.Println(err)
	}
	// Output:
}

func ExampleFile_GetWorkbookPrOptions() {
	f := NewFile()
	var codeName CodeName
	if err := f.GetWorkbookPrOptions(&codeName); err != nil {
		fmt.Println(err)
	}
	var filterPrivacy FilterPrivacy
	if err := f.GetWorkbookPrOptions(&filterPrivacy); err != nil {
		fmt.Println(err)
	}
	fmt.Println("Defaults:")
	fmt.Printf("- codeName: %q\n", codeName)
	fmt.Printf("- filterPrivacy: %t\n", filterPrivacy)
	// Output:
	// Defaults:
	// - codeName: ""
	// - filterPrivacy: true
}

func TestWorkbookPr(t *testing.T) {
	f := NewFile()
	wb := f.workbookReader()
	wb.WorkbookPr = nil
	var codeName CodeName
	assert.NoError(t, f.GetWorkbookPrOptions(&codeName))
	assert.Equal(t, "", string(codeName))
	assert.NoError(t, f.SetWorkbookPrOptions(CodeName("code")))
	assert.NoError(t, f.GetWorkbookPrOptions(&codeName))
	assert.Equal(t, "code", string(codeName))

	var filterPrivacy FilterPrivacy
	assert.NoError(t, f.GetWorkbookPrOptions(&filterPrivacy))
	assert.Equal(t, true, bool(filterPrivacy))
	assert.NoError(t, f.SetWorkbookPrOptions(FilterPrivacy(false)))
	assert.NoError(t, f.GetWorkbookPrOptions(&filterPrivacy))
	assert.Equal(t, false, bool(filterPrivacy))
}
