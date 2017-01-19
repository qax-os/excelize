package excelize

import (
	"encoding/xml"
	"strconv"
	"strings"
)

// GetCellValue provides function to get value from cell by given sheet index
// and axis in XLSX file. The value of the merged cell is not available
// currently.
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

// GetCellFormula provides function to get formula from cell by given sheet
// index and axis in XLSX file.
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

// SetCellHyperLink provides function to set cell hyperlink by given sheet index
// and link URL address. Only support external link currently.
func (f *File) SetCellHyperLink(sheet, axis, link string) {
	axis = strings.ToUpper(axis)
	var xlsx xlsxWorksheet
	name := "xl/worksheets/" + strings.ToLower(sheet) + ".xml"
	xml.Unmarshal([]byte(f.readXML(name)), &xlsx)
	rID := f.addSheetRelationships(sheet, SourceRelationshipHyperLink, link, "External")
	hyperlink := xlsxHyperlink{
		Ref: axis,
		RID: "rId" + strconv.Itoa(rID),
	}
	if xlsx.Hyperlinks != nil {
		xlsx.Hyperlinks.Hyperlink = append(xlsx.Hyperlinks.Hyperlink, hyperlink)
	} else {
		hyperlinks := xlsxHyperlinks{}
		hyperlinks.Hyperlink = append(hyperlinks.Hyperlink, hyperlink)
		xlsx.Hyperlinks = &hyperlinks
	}
	output, _ := xml.Marshal(xlsx)
	f.saveFileList(name, replaceWorkSheetsRelationshipsNameSpace(string(output)))
}

// SetCellFormula provides function to set cell formula by given string and
// sheet index.
func (f *File) SetCellFormula(sheet, axis, formula string) {
	axis = strings.ToUpper(axis)
	var xlsx xlsxWorksheet
	col := string(strings.Map(letterOnlyMapF, axis))
	row, _ := strconv.Atoi(strings.Map(intOnlyMapF, axis))
	xAxis := row - 1
	yAxis := titleToNumber(col)

	name := "xl/worksheets/" + strings.ToLower(sheet) + ".xml"
	xml.Unmarshal([]byte(f.readXML(name)), &xlsx)

	rows := xAxis + 1
	cell := yAxis + 1

	xlsx = completeRow(xlsx, rows, cell)
	xlsx = completeCol(xlsx, rows, cell)

	if xlsx.SheetData.Row[xAxis].C[yAxis].F != nil {
		xlsx.SheetData.Row[xAxis].C[yAxis].F.Content = formula
	} else {
		f := xlsxF{
			Content: formula,
		}
		xlsx.SheetData.Row[xAxis].C[yAxis].F = &f
	}
	output, _ := xml.Marshal(xlsx)
	f.saveFileList(name, replaceWorkSheetsRelationshipsNameSpace(string(output)))
}
