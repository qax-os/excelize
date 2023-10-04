package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestWorkbookProps(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetWorkbookProps(nil))
	wb, err := f.workbookReader()
	assert.NoError(t, err)
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
	// Test set workbook properties with unsupported charset workbook
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	assert.EqualError(t, f.SetWorkbookProps(&expected), "XML syntax error on line 1: invalid UTF-8")
	// Test get workbook properties with unsupported charset workbook
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	_, err = f.GetWorkbookProps()
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
}

func TestDeleteWorkbookRels(t *testing.T) {
	f := NewFile()
	// Test delete pivot table without worksheet relationships
	f.Relationships.Delete("xl/_rels/workbook.xml.rels")
	f.Pkg.Delete("xl/_rels/workbook.xml.rels")
	rID, err := f.deleteWorkbookRels("", "")
	assert.Empty(t, rID)
	assert.NoError(t, err)
}
