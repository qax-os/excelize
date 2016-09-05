package excelize

import (
	"encoding/xml"
	"strconv"
	"strings"
)

// GetCellValue provide function get value from cell by given sheet index and axis in XLSX file
func (f *File) GetCellValue(sheet string, axis string) string {
	axis = strings.ToUpper(axis)
	var xlsx xlsxWorksheet
	row := getRowIndex(axis)
	xAxis := row - 1
	name := `xl/worksheets/` + strings.ToLower(sheet) + `.xml`
	xml.Unmarshal([]byte(f.readXML(name)), &xlsx)
	rows := len(xlsx.SheetData.Row)
	if rows <= xAxis {
		return ``
	}
	for _, v := range xlsx.SheetData.Row[xAxis].C {
		if xlsx.SheetData.Row[xAxis].R == row {
			if axis == v.R {
				switch v.T {
				case "s":
					shardStrings := xlsxSST{}
					xlsxSI := 0
					xlsxSI, _ = strconv.Atoi(v.V)
					xml.Unmarshal([]byte(f.readXML(`xl/sharedStrings.xml`)), &shardStrings)
					return shardStrings.SI[xlsxSI].T
				case "str":
					return v.V
				default:
					return v.V
				}
			}
		}
	}
	return ``
}
