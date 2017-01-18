package excelize

import (
	"encoding/xml"
	"strings"
)

// SetColWidth provides function to set the width of a single column or multiple columns.
// For example:
//
//    xlsx := excelize.CreateFile()
//    xlsx.SetColWidth("Sheet1", "A", "H", 20)
//    err := xlsx.Save()
//    if err != nil {
//        fmt.Println(err)
//        os.Exit(1)
//    }
//
func (f *File) SetColWidth(sheet, startcol, endcol string, width float64) {
	min := titleToNumber(strings.ToUpper(startcol)) + 1
	max := titleToNumber(strings.ToUpper(endcol)) + 1
	if min > max {
		min, max = max, min
	}
	var xlsx xlsxWorksheet
	name := "xl/worksheets/" + strings.ToLower(sheet) + ".xml"
	xml.Unmarshal([]byte(f.readXML(name)), &xlsx)
	col := xlsxCol{
		Min:         min,
		Max:         max,
		Width:       width,
		CustomWidth: true,
	}
	if xlsx.Cols != nil {
		xlsx.Cols.Col = append(xlsx.Cols.Col, col)
	} else {
		cols := xlsxCols{}
		cols.Col = append(cols.Col, col)
		xlsx.Cols = &cols
	}
	output, _ := xml.Marshal(xlsx)
	f.saveFileList(name, replaceWorkSheetsRelationshipsNameSpace(string(output)))
}
