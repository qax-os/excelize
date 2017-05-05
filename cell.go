package excelize

import (
	"encoding/xml"
	"strconv"
	"strings"
)

// GetCellValue provides function to get formatted value from cell by given
// sheet index and axis in XLSX file. If it is possible to apply a format to the
// cell value, it will do so, if not then an error will be returned, along with
// the raw value of the cell.
func (f *File) GetCellValue(sheet, axis string) string {
	xlsx := f.workSheetReader(sheet)
	axis = strings.ToUpper(axis)
	if xlsx.MergeCells != nil {
		for i := 0; i < len(xlsx.MergeCells.Cells); i++ {
			if checkCellInArea(axis, xlsx.MergeCells.Cells[i].Ref) {
				axis = strings.Split(xlsx.MergeCells.Cells[i].Ref, ":")[0]
			}
		}
	}
	row, _ := strconv.Atoi(strings.Map(intOnlyMapF, axis))
	xAxis := row - 1
	rows := len(xlsx.SheetData.Row)
	if rows > 1 {
		lastRow := xlsx.SheetData.Row[rows-1].R
		if lastRow >= rows {
			rows = lastRow
		}
	}
	if rows < xAxis {
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
				return f.formattedValue(r.S, shardStrings.SI[xlsxSI].T)
			case "str":
				return f.formattedValue(r.S, r.V)
			default:
				return f.formattedValue(r.S, r.V)
			}
		}
	}
	return ""
}

// formattedValue provides function to returns a value after formatted. If it is
// possible to apply a format to the cell value, it will do so, if not then an
// error will be returned, along with the raw value of the cell.
func (f *File) formattedValue(s int, v string) string {
	if s == 0 {
		return v
	}
	var styleSheet xlsxStyleSheet
	xml.Unmarshal([]byte(f.readXML("xl/styles.xml")), &styleSheet)
	ok := builtInNumFmtFunc[styleSheet.CellXfs.Xf[s].NumFmtID]
	if ok != nil {
		return ok(styleSheet.CellXfs.Xf[s].NumFmtID, v)
	}
	return v
}

// GetCellFormula provides function to get formula from cell by given sheet
// index and axis in XLSX file.
func (f *File) GetCellFormula(sheet, axis string) string {
	xlsx := f.workSheetReader(sheet)
	axis = strings.ToUpper(axis)
	if xlsx.MergeCells != nil {
		for i := 0; i < len(xlsx.MergeCells.Cells); i++ {
			if checkCellInArea(axis, xlsx.MergeCells.Cells[i].Ref) {
				axis = strings.Split(xlsx.MergeCells.Cells[i].Ref, ":")[0]
			}
		}
	}
	row, _ := strconv.Atoi(strings.Map(intOnlyMapF, axis))
	xAxis := row - 1
	rows := len(xlsx.SheetData.Row)
	if rows > 1 {
		lastRow := xlsx.SheetData.Row[rows-1].R
		if lastRow >= rows {
			rows = lastRow
		}
	}
	if rows < xAxis {
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

// SetCellFormula provides function to set cell formula by given string and
// sheet index.
func (f *File) SetCellFormula(sheet, axis, formula string) {
	xlsx := f.workSheetReader(sheet)
	axis = strings.ToUpper(axis)
	if xlsx.MergeCells != nil {
		for i := 0; i < len(xlsx.MergeCells.Cells); i++ {
			if checkCellInArea(axis, xlsx.MergeCells.Cells[i].Ref) {
				axis = strings.Split(xlsx.MergeCells.Cells[i].Ref, ":")[0]
			}
		}
	}
	col := string(strings.Map(letterOnlyMapF, axis))
	row, _ := strconv.Atoi(strings.Map(intOnlyMapF, axis))
	xAxis := row - 1
	yAxis := titleToNumber(col)

	rows := xAxis + 1
	cell := yAxis + 1

	completeRow(xlsx, rows, cell)
	completeCol(xlsx, rows, cell)

	if xlsx.SheetData.Row[xAxis].C[yAxis].F != nil {
		xlsx.SheetData.Row[xAxis].C[yAxis].F.Content = formula
	} else {
		f := xlsxF{
			Content: formula,
		}
		xlsx.SheetData.Row[xAxis].C[yAxis].F = &f
	}
}

// SetCellHyperLink provides function to set cell hyperlink by given sheet index
// and link URL address. Only support external link currently.
func (f *File) SetCellHyperLink(sheet, axis, link string) {
	xlsx := f.workSheetReader(sheet)
	axis = strings.ToUpper(axis)
	if xlsx.MergeCells != nil {
		for i := 0; i < len(xlsx.MergeCells.Cells); i++ {
			if checkCellInArea(axis, xlsx.MergeCells.Cells[i].Ref) {
				axis = strings.Split(xlsx.MergeCells.Cells[i].Ref, ":")[0]
			}
		}
	}
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
}

// MergeCell provides function to merge cells by given coordinate area and sheet
// name. For example create a merged cell of D3:E9 on Sheet1:
//
//    xlsx.MergeCell("sheet1", "D3", "E9")
//
// If you create a merged cell that overlaps with another existing merged cell,
// those merged cells that already exist will be removed.
func (f *File) MergeCell(sheet, hcell, vcell string) {
	if hcell == vcell {
		return
	}

	hcell = strings.ToUpper(hcell)
	vcell = strings.ToUpper(vcell)

	// Coordinate conversion, convert C1:B3 to 2,0,1,2.
	hcol := string(strings.Map(letterOnlyMapF, hcell))
	hrow, _ := strconv.Atoi(strings.Map(intOnlyMapF, hcell))
	hyAxis := hrow - 1
	hxAxis := titleToNumber(hcol)

	vcol := string(strings.Map(letterOnlyMapF, vcell))
	vrow, _ := strconv.Atoi(strings.Map(intOnlyMapF, vcell))
	vyAxis := vrow - 1
	vxAxis := titleToNumber(vcol)

	if vxAxis < hxAxis {
		hcell, vcell = vcell, hcell
		vxAxis, hxAxis = hxAxis, vxAxis
	}

	if vyAxis < hyAxis {
		hcell, vcell = vcell, hcell
		vyAxis, hyAxis = hyAxis, vyAxis
	}

	xlsx := f.workSheetReader(sheet)
	if xlsx.MergeCells != nil {
		mergeCell := xlsxMergeCell{}
		// Correct the coordinate area, such correct C1:B3 to B1:C3.
		mergeCell.Ref = toAlphaString(hxAxis+1) + strconv.Itoa(hyAxis+1) + ":" + toAlphaString(vxAxis+1) + strconv.Itoa(vyAxis+1)
		// Delete the merged cells of the overlapping area.
		for i := 0; i < len(xlsx.MergeCells.Cells); i++ {
			if checkCellInArea(hcell, xlsx.MergeCells.Cells[i].Ref) || checkCellInArea(strings.Split(xlsx.MergeCells.Cells[i].Ref, ":")[0], mergeCell.Ref) {
				xlsx.MergeCells.Cells = append(xlsx.MergeCells.Cells[:i], xlsx.MergeCells.Cells[i+1:]...)
			} else if checkCellInArea(vcell, xlsx.MergeCells.Cells[i].Ref) || checkCellInArea(strings.Split(xlsx.MergeCells.Cells[i].Ref, ":")[1], mergeCell.Ref) {
				xlsx.MergeCells.Cells = append(xlsx.MergeCells.Cells[:i], xlsx.MergeCells.Cells[i+1:]...)
			}
		}
		xlsx.MergeCells.Cells = append(xlsx.MergeCells.Cells, &mergeCell)
	} else {
		mergeCell := xlsxMergeCell{}
		// Correct the coordinate area, such correct C1:B3 to B1:C3.
		mergeCell.Ref = toAlphaString(hxAxis+1) + strconv.Itoa(hyAxis+1) + ":" + toAlphaString(vxAxis+1) + strconv.Itoa(vyAxis+1)
		mergeCells := xlsxMergeCells{}
		mergeCells.Cells = append(mergeCells.Cells, &mergeCell)
		xlsx.MergeCells = &mergeCells
	}
}

// checkCellInArea provides function to determine if a given coordinate is
// within an area.
func checkCellInArea(cell, area string) bool {
	result := false
	cell = strings.ToUpper(cell)
	col := string(strings.Map(letterOnlyMapF, cell))
	row, _ := strconv.Atoi(strings.Map(intOnlyMapF, cell))
	xAxis := row - 1
	yAxis := titleToNumber(col)

	ref := strings.Split(area, ":")
	hCol := string(strings.Map(letterOnlyMapF, ref[0]))
	hRow, _ := strconv.Atoi(strings.Map(intOnlyMapF, ref[0]))
	hyAxis := hRow - 1
	hxAxis := titleToNumber(hCol)

	vCol := string(strings.Map(letterOnlyMapF, ref[1]))
	vRow, _ := strconv.Atoi(strings.Map(intOnlyMapF, ref[1]))
	vyAxis := vRow - 1
	vxAxis := titleToNumber(vCol)

	if hxAxis <= yAxis && yAxis <= vxAxis && hyAxis <= xAxis && xAxis <= vyAxis {
		result = true
	}

	return result
}
