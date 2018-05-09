package excelize

import (
	"bytes"
	"encoding/xml"
	"testing"
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
	if err != nil {
		t.Fatal(err)
	}

	newFile, err := OpenReader(&buffer)
	if err != nil {
		t.Fatal(err)
	}

	chartsNum := newFile.countCharts()
	if chartsNum != 1 {
		t.Fatalf("Expected 1 chart, actual %d", chartsNum)
	}

	var (
		workdir decodeWsDr
		anchor  decodeTwoCellAnchor
	)

	content, ok := newFile.XLSX["xl/drawings/drawing1.xml"]
	if !ok {
		t.Fatal("Can't open the chart")
	}

	err = xml.Unmarshal([]byte(content), &workdir)
	if err != nil {
		t.Fatal(err)
	}

	err = xml.Unmarshal([]byte("<decodeTwoCellAnchor>"+workdir.TwoCellAnchor[0].Content+"</decodeTwoCellAnchor>"), &anchor)
	if err != nil {
		t.Fatal(err)
	}

	if anchor.From.Col != 4 || anchor.From.Row != 3 {
		t.Fatalf("From: Expected column 4, row 3, actual column %d, row %d", anchor.From.Col, anchor.From.Row)
	}
	if anchor.To.Col != 14 || anchor.To.Row != 27 {
		t.Fatalf("To: Expected column 14, row 27, actual column %d, row %d", anchor.To.Col, anchor.To.Row)
	}

}
