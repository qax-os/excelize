package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestCalcChainReader(t *testing.T) {
	f, err := NewFile()
	assert.NoError(t, err)
	f.CalcChain = nil
	f.Pkg.Store(defaultXMLPathCalcChain, MacintoshCyrillicCharset)
	f.calcChainReader()
}

func TestDeleteCalcChain(t *testing.T) {
	f, err := NewFile()
	assert.NoError(t, err)
	f.CalcChain = &xlsxCalcChain{C: []xlsxCalcChainC{}}
	f.ContentTypes.Overrides = append(f.ContentTypes.Overrides, xlsxOverride{
		PartName: "/xl/calcChain.xml",
	})
	f.deleteCalcChain(1, "A1")
}
