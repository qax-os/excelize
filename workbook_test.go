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
	fmt.Println("Defaults:")
	fmt.Printf("- codeName: %q\n", codeName)
	// Output:
	// Defaults:
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
}
