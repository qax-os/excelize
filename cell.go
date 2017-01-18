package excelize

import (
	"encoding/xml"
	"strconv"
	"strings"
)

// GetCellValue provide function get value from cell by given sheet index and
// axis in XLSX file. The value of the merged cell is not available currently.
func (f *File) GetCellValue(sheet string, axis string) string {
	axis = strings.ToUpper(axis)
	var xlsx xlsxWorksheet
	row, _ := strconv.Atoi(strings.Map(intOnlyMapF, axis))
	xAxis := row - 1
	name := "xl/worksheets/" + strings.ToLower(sheet) + ".xml"
	xml.Unmarshal([]byte(f.readXML(name)), &xlsx)
	rows := len(xlsx.SheetData.Row)
	if rows > 1 {
		lastRow := xlsx.SheetData.Row[rows-1].R
		if lastRow >= rows {
			rows = lastRow
		}
	}
	if rows <= xAxis {
		return ""
	}
	for _, v := range xlsx.SheetData.Row {
		if v.R != row {
			continue
		}
		for _, r := range v.C {
			if axis != r.R {
				continue
			}
			switch r.T {
			case "s":
				shardStrings := xlsxSST{}
				xlsxSI := 0
				xlsxSI, _ = strconv.Atoi(r.V)
				xml.Unmarshal([]byte(f.readXML("xl/sharedStrings.xml")), &shardStrings)
				return shardStrings.SI[xlsxSI].T
			case "str":
				return r.V
			default:
				return r.V
			}
		}
	}
	return ""
}

// GetCellFormula provide function get formula from cell by given sheet index
// and axis in XLSX file.
func (f *File) GetCellFormula(sheet string, axis string) string {
	axis = strings.ToUpper(axis)
	var xlsx xlsxWorksheet
	row, _ := strconv.Atoi(strings.Map(intOnlyMapF, axis))
	xAxis := row - 1
	name := "xl/worksheets/" + strings.ToLower(sheet) + ".xml"
	xml.Unmarshal([]byte(f.readXML(name)), &xlsx)
	rows := len(xlsx.SheetData.Row)
	if rows > 1 {
		lastRow := xlsx.SheetData.Row[rows-1].R
		if lastRow >= rows {
			rows = lastRow
		}
	}
	if rows <= xAxis {
		return ""
	}
	for _, v := range xlsx.SheetData.Row {
		if v.R != row {
			continue
		}
		for _, f := range v.C {
			if axis != f.R {
				continue
			}
			if f.F != nil {
				return f.F.Content
			}
		}
	}
	return ""
}
