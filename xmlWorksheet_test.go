package excelize

import (
	"encoding/xml"
	"testing"

	"github.com/stretchr/testify/require"
)

func TestXmlWorksheet_MarshalXML(t *testing.T) {
	worksheet := xlsxWorksheet{
		PageSetUp: &xlsxPageSetUp{
			RID:           "aaa",
			BlackAndWhite: true,
		},
		Hyperlinks: &xlsxHyperlinks{
			Hyperlink: []xlsxHyperlink{
				{
					Ref: "google.com",
					RID: "bbbbb",
				},
			},
		},
		TableParts: &xlsxTableParts{
			TableParts: []*xlsxTablePart{
				{
					RID: "ccccc",
				},
			},
		},
		Picture: &xlsxPicture{
			RID: "dddd",
		},
		LegacyDrawing: &xlsxLegacyDrawing{
			RID: "eeee",
		},
	}
	b, err := xml.Marshal(worksheet)
	require.NoError(t, err)

	b2 := replaceRelationshipsBytes(replaceWorkSheetsRelationshipsNameSpaceBytes(b))
	t.Log(string(b))
	require.Equal(t, string(b), string(b2))
}
