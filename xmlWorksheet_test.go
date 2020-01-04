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
	// check integration with old version
	b, err := xml.Marshal(worksheet)
	require.NoError(t, err)
	b = replaceRelationshipsBytes(replaceWorkSheetsRelationshipsNameSpaceBytes(b))

	worksheet.BaseAtrr = getDefaultAttrs()
	b2, err := xml.Marshal(worksheet)
	require.NoError(t, err)

	t.Log(string(b))
	require.Equal(t, string(b), string(b2))
}
