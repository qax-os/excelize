package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

func ExampleFile_SetWorkbookPrOptions() {
	f := NewFile()
	if err := f.SetWorkbookPrOptions(
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
		filterPrivacy FilterPrivacy
		codeName      CodeName
	)
	if err := f.GetWorkbookPrOptions(&filterPrivacy); err != nil {
		fmt.Println(err)
	}
	if err := f.GetWorkbookPrOptions(&codeName); err != nil {
		fmt.Println(err)
	}
	fmt.Println("Defaults:")
	fmt.Printf("- filterPrivacy: %t\n", filterPrivacy)
	fmt.Printf("- codeName: %q\n", codeName)
	// Output:
	// Defaults:
	// - filterPrivacy: true
	// - codeName: ""
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

	wb.WorkbookPr = nil
	var filterPrivacy FilterPrivacy
	assert.NoError(t, f.GetWorkbookPrOptions(&filterPrivacy))
	assert.Equal(t, false, bool(filterPrivacy))
}
