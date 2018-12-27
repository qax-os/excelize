package excelize

import (
	"bytes"
	"encoding/xml"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestChartSize(t *testing.T) {

	var buffer bytes.Buffer

	categories := map[string]string{"A2": "Small", "A3": "Normal", "A4": "Large", "B1": "Apple", "C1": "Orange", "D1": "Pear"}
	values := map[string]int{"B2": 2, "C2": 3, "D2": 3, "B3": 5, "C3": 2, "D3": 4, "B4": 6, "C4": 7, "D4": 8}
	xlsx := NewFile()
	for k, v := range categories {
		xlsx.SetCellValue("Sheet1", k, v)
	}
	for k, v := range values {
		xlsx.SetCellValue("Sheet1", k, v)
	}
	xlsx.AddChart("Sheet1", "E4", `{"type":"col3DClustered","dimension":{"width":640, "height":480},"series":[{"name":"Sheet1!$A$2","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$2:$D$2"},{"name":"Sheet1!$A$3","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$3:$D$3"},{"name":"Sheet1!$A$4","categories":"Sheet1!$B$1:$D$1","values":"Sheet1!$B$4:$D$4"}],"title":{"name":"Fruit 3D Clustered Column Chart"}}`)
	// Save xlsx file by the given path.
	err := xlsx.Write(&buffer)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	newFile, err := OpenReader(&buffer)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	chartsNum := newFile.countCharts()
	if !assert.Equal(t, 1, chartsNum, "Expected 1 chart, actual %d", chartsNum) {
		t.FailNow()
	}

	var (
		workdir decodeWsDr
		anchor  decodeTwoCellAnchor
	)

	content, ok := newFile.XLSX["xl/drawings/drawing1.xml"]
	assert.True(t, ok, "Can't open the chart")

	err = xml.Unmarshal([]byte(content), &workdir)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	err = xml.Unmarshal([]byte("<decodeTwoCellAnchor>"+workdir.TwoCellAnchor[0].Content+"</decodeTwoCellAnchor>"), &anchor)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	if !assert.Equal(t, 4, anchor.From.Col, "Expected 'from' column 4") ||
		!assert.Equal(t, 3, anchor.From.Row, "Expected 'from' row 3") {

		t.FailNow()
	}

	if !assert.Equal(t, 14, anchor.To.Col, "Expected 'to' column 14") ||
		!assert.Equal(t, 27, anchor.To.Row, "Expected 'to' row 27") {

		t.FailNow()
	}
}
