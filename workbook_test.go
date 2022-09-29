package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestWorkbookProps(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetWorkbookProps(nil))
	wb := f.workbookReader()
	wb.WorkbookPr = nil
	expected := WorkbookPropsOptions{
		Date1904:      boolPtr(true),
		FilterPrivacy: boolPtr(true),
		CodeName:      stringPtr("code"),
	}
	assert.NoError(t, f.SetWorkbookProps(&expected))
	opts, err := f.GetWorkbookProps()
	assert.NoError(t, err)
	assert.Equal(t, expected, opts)
}
