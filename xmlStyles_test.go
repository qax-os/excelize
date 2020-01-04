package excelize

import (
	"encoding/xml"
	"testing"

	"github.com/stretchr/testify/require"
)

func TestXlsxStyleSheet_MarshalXML(t *testing.T) {
	relationships := xlsxStyleSheet{}
	// get xml by replace bytes
	b, err := xml.Marshal(relationships)
	require.NoError(t, err)
	b = replaceStyleRelationshipsNameSpaceBytes(b)
	// get xml by create a default attr array
	relationships.BaseAttr = getDefaultAttrs()
	b2, err := xml.Marshal(relationships)
	require.NoError(t, err)
	require.Equal(t, string(b), string(b2))
	t.Log(string(b2))
}
