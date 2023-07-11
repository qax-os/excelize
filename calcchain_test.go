package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestCalcChainReader(t *testing.T) {
	f := NewFile()
	// Test read calculation chain with unsupported charset
	f.CalcChain = nil
	f.Pkg.Store(defaultXMLPathCalcChain, MacintoshCyrillicCharset)
	_, err := f.calcChainReader()
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
}

func TestDeleteCalcChain(t *testing.T) {
	f := NewFile()
	f.CalcChain = &xlsxCalcChain{C: []xlsxCalcChainC{}}
	f.ContentTypes.Overrides = append(f.ContentTypes.Overrides, xlsxOverride{
		PartName: "/xl/calcChain.xml",
	})
	assert.NoError(t, f.deleteCalcChain(1, "A1"))

	f.CalcChain = nil
	f.Pkg.Store(defaultXMLPathCalcChain, MacintoshCyrillicCharset)
	assert.EqualError(t, f.deleteCalcChain(1, "A1"), "XML syntax error on line 1: invalid UTF-8")

	f.CalcChain = nil
	f.Pkg.Store(defaultXMLPathCalcChain, MacintoshCyrillicCharset)
	assert.EqualError(t, f.SetCellFormula("Sheet1", "A1", ""), "XML syntax error on line 1: invalid UTF-8")

	formulaType, ref := STCellFormulaTypeShared, "C1:C5"
	assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "=A1+B1", FormulaOpts{Ref: &ref, Type: &formulaType}))

	// Test delete calculation chain with unsupported charset calculation chain
	f.CalcChain = nil
	f.Pkg.Store(defaultXMLPathCalcChain, MacintoshCyrillicCharset)
	assert.EqualError(t, f.SetCellValue("Sheet1", "C1", true), "XML syntax error on line 1: invalid UTF-8")

	// Test delete calculation chain with unsupported charset content types
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, f.deleteCalcChain(1, "A1"), "XML syntax error on line 1: invalid UTF-8")
}
