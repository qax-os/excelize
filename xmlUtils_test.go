package excelize

import (
	"encoding/xml"
	"testing"

	"github.com/stretchr/testify/require"
)

func TestRelationship_MarshalXMLAttr(t *testing.T) {
	data2 := []byte(`<xlsxPivotCache cacheId="5" xmlns:relationships="http://schemas.openxmlformats.org/officeDocument/2006/relationships" relationships:id="6"></xlsxPivotCache>`)
	var pivotCache xlsxPivotCache
	require.NoError(t, xml.Unmarshal(data2, &pivotCache))
	require.Equal(t, 5, pivotCache.CacheID)
	require.Equal(t, relationship("6"), pivotCache.RID)
	data := &xlsxPivotCache{
		CacheID: 5,
		RID:     "6",
	}
	b, err := xml.Marshal(data)
	require.NoError(t, err)
	t.Log(string(b))
}
