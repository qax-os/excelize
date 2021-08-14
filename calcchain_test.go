package excelize

import "testing"

func TestCalcChainReader(t *testing.T) {
	f := NewFile()
	f.CalcChain = nil
	f.Pkg.Store("xl/calcChain.xml", MacintoshCyrillicCharset)
	f.calcChainReader()
}

func TestDeleteCalcChain(t *testing.T) {
	f := NewFile()
	f.CalcChain = &xlsxCalcChain{C: []xlsxCalcChainC{}}
	f.ContentTypes.Overrides = append(f.ContentTypes.Overrides, xlsxOverride{
		PartName: "/xl/calcChain.xml",
	})
	f.deleteCalcChain(1, "A1")
}
