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
	wb.WorkbookPr = nil
	opts, err = f.GetWorkbookProps()
	assert.NoError(t, err)
	assert.Equal(t, WorkbookPropsOptions{}, opts)
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

func TestCalcProps(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCalcProps(nil))
	wb, err := f.workbookReader()
	assert.NoError(t, err)
	wb.CalcPr = nil
	expected := CalcPropsOptions{
		FullCalcOnLoad:        boolPtr(true),
		CalcID:                uintPtr(122211),
		ConcurrentManualCount: uintPtr(5),
		IterateCount:          uintPtr(10),
		ConcurrentCalc:        boolPtr(true),
	}
	assert.NoError(t, f.SetCalcProps(&expected))
	opts, err := f.GetCalcProps()
	assert.NoError(t, err)
	assert.Equal(t, expected, opts)
	wb.CalcPr = nil
	opts, err = f.GetCalcProps()
	assert.NoError(t, err)
	assert.Equal(t, CalcPropsOptions{}, opts)
	// Test set calculation properties with unsupported optional value
	assert.Equal(t, newInvalidOptionalValue("CalcMode", "AUTO", supportedCalcMode), f.SetCalcProps(&CalcPropsOptions{CalcMode: stringPtr("AUTO")}))
	assert.Equal(t, newInvalidOptionalValue("RefMode", "a1", supportedRefMode), f.SetCalcProps(&CalcPropsOptions{RefMode: stringPtr("a1")}))
	// Test set calculation properties with unsupported charset workbook
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	assert.EqualError(t, f.SetCalcProps(&expected), "XML syntax error on line 1: invalid UTF-8")
	// Test get calculation properties with unsupported charset workbook
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	_, err = f.GetCalcProps()
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
