package excelize

import (
	"fmt"
	"testing"
)

func TestRichDataInsert(t *testing.T) {
	f := NewFile()
	f.Pkg.Store(defaultXMLMetadata, []byte(`<metadata><valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>`))
	f.Pkg.Store(defaultXMLRdRichValuePart, []byte(`<rvData count="1"><rv s="0"><v>0</v><v>5</v></rv></rvData>`))
	f.Pkg.Store(defaultXMLRdRichValueRel, []byte(`<richValueRels><rel r:id="rId1"/></richValueRels>`))
	f.Pkg.Store(defaultXMLRdRichValueRelRels, []byte(fmt.Sprintf(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="%s" Target="../media/image1.png"/></Relationships>`, SourceRelationshipImage)))
	f.Sheet.Store("xl/worksheets/sheet1.xml", &xlsxWorksheet{
		SheetData: xlsxSheetData{Row: []xlsxRow{
			{R: 1, C: []xlsxC{{R: "A1", T: "e", V: formulaErrorVALUE, Vm: uintPtr(1)}}},
		}},
	})
}
