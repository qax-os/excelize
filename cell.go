package excelize

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
	"time"
)

// mergeCellsParser provides function to check merged cells in worksheet by
// given axis.
func (f *File) mergeCellsParser(xlsx *xlsxWorksheet, axis string) string {
	axis = strings.ToUpper(axis)
	if xlsx.MergeCells != nil {
		for i := 0; i < len(xlsx.MergeCells.Cells); i++ {
			if checkCellInArea(axis, xlsx.MergeCells.Cells[i].Ref) {
				axis = strings.Split(xlsx.MergeCells.Cells[i].Ref, ":")[0]
			}
		}
	}
	return axis
}

// SetCellValue provides function to set value of a cell. The following shows
// the supported data types:
//
//    int
//    int8
//    int16
//    int32
//    int64
//    uint
//    uint8
//    uint16
//    uint32
//    uint64
//    float32
//    float64
//    string
//    []byte
//    time.Duration
//    time.Time
//    nil
//
// Note that default date format is m/d/yy h:mm of time.Time type value. You can
// set numbers format by SetCellStyle() method.
func (f *File) SetCellValue(sheet, axis string, value interface{}) {
	switch t := value.(type) {
	case float32:
		f.SetCellDefault(sheet, axis, strconv.FormatFloat(float64(value.(float32)), 'f', -1, 32))
	case float64:
		f.SetCellDefault(sheet, axis, strconv.FormatFloat(float64(value.(float64)), 'f', -1, 64))
	case string:
		f.SetCellStr(sheet, axis, t)
	case []byte:
		f.SetCellStr(sheet, axis, string(t))
	case time.Duration:
		f.SetCellDefault(sheet, axis, strconv.FormatFloat(float64(value.(time.Duration).Seconds()/86400), 'f', -1, 32))
		f.setDefaultTimeStyle(sheet, axis, 21)
	case time.Time:
		f.SetCellDefault(sheet, axis, strconv.FormatFloat(float64(timeToExcelTime(timeToUTCTime(value.(time.Time)))), 'f', -1, 64))
		f.setDefaultTimeStyle(sheet, axis, 22)
	case nil:
		f.SetCellStr(sheet, axis, "")
	default:
		f.setCellIntValue(sheet, axis, value)
	}
}

// setCellIntValue provides function to set int value of a cell.
func (f *File) setCellIntValue(sheet, axis string, value interface{}) {
	switch value.(type) {
	case int:
		f.SetCellInt(sheet, axis, value.(int))
	case int8:
		f.SetCellInt(sheet, axis, int(value.(int8)))
	case int16:
		f.SetCellInt(sheet, axis, int(value.(int16)))
	case int32:
		f.SetCellInt(sheet, axis, int(value.(int32)))
	case int64:
		f.SetCellInt(sheet, axis, int(value.(int64)))
	case uint:
		f.SetCellInt(sheet, axis, int(value.(uint)))
	case uint8:
		f.SetCellInt(sheet, axis, int(value.(uint8)))
	case uint16:
		f.SetCellInt(sheet, axis, int(value.(uint16)))
	case uint32:
		f.SetCellInt(sheet, axis, int(value.(uint32)))
	case uint64:
		f.SetCellInt(sheet, axis, int(value.(uint64)))
	default:
		f.SetCellStr(sheet, axis, fmt.Sprintf("%v", value))
	}
}

// GetCellValue provides function to get formatted value from cell by given
// worksheet name and axis in XLSX file. If it is possible to apply a format to
// the cell value, it will do so, if not then an error will be returned, along
// with the raw value of the cell.
func (f *File) GetCellValue(sheet, axis string) string {
	xlsx := f.workSheetReader(sheet)
	axis = f.mergeCellsParser(xlsx, axis)
	row, err := strconv.Atoi(strings.Map(intOnlyMapF, axis))
	if err != nil {
		return ""
	}
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
	for k := range xlsx.SheetData.Row {
		if xlsx.SheetData.Row[k].R == row {
			for i := range xlsx.SheetData.Row[k].C {
				if axis == xlsx.SheetData.Row[k].C[i].R {
					val, _ := xlsx.SheetData.Row[k].C[i].getValueFrom(f, f.sharedStringsReader())
					return val
				}
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
	styleSheet := f.stylesReader()
	ok := builtInNumFmtFunc[styleSheet.CellXfs.Xf[s].NumFmtID]
	if ok != nil {
		return ok(styleSheet.CellXfs.Xf[s].NumFmtID, v)
	}
	return v
}

// GetCellStyle provides function to get cell style index by given worksheet
// name and cell coordinates.
func (f *File) GetCellStyle(sheet, axis string) int {
	xlsx := f.workSheetReader(sheet)
	axis = f.mergeCellsParser(xlsx, axis)
	col := string(strings.Map(letterOnlyMapF, axis))
	row, err := strconv.Atoi(strings.Map(intOnlyMapF, axis))
	if err != nil {
		return 0
	}
	xAxis := row - 1
	yAxis := TitleToNumber(col)

	rows := xAxis + 1
	cell := yAxis + 1

	completeRow(xlsx, rows, cell)
	completeCol(xlsx, rows, cell)

	return f.prepareCellStyle(xlsx, cell, xlsx.SheetData.Row[xAxis].C[yAxis].S)
}

// GetCellFormula provides function to get formula from cell by given worksheet
// name and axis in XLSX file.
func (f *File) GetCellFormula(sheet, axis string) string {
	xlsx := f.workSheetReader(sheet)
	axis = f.mergeCellsParser(xlsx, axis)
	row, err := strconv.Atoi(strings.Map(intOnlyMapF, axis))
	if err != nil {
		return ""
	}
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
	for k := range xlsx.SheetData.Row {
		if xlsx.SheetData.Row[k].R == row {
			for i := range xlsx.SheetData.Row[k].C {
				if axis == xlsx.SheetData.Row[k].C[i].R {
					if xlsx.SheetData.Row[k].C[i].F != nil {
						return xlsx.SheetData.Row[k].C[i].F.Content
					}
				}
			}
		}
	}
	return ""
}

// SetCellFormula provides function to set cell formula by given string and
// worksheet name.
func (f *File) SetCellFormula(sheet, axis, formula string) {
	xlsx := f.workSheetReader(sheet)
	axis = f.mergeCellsParser(xlsx, axis)
	col := string(strings.Map(letterOnlyMapF, axis))
	row, err := strconv.Atoi(strings.Map(intOnlyMapF, axis))
	if err != nil {
		return
	}
	xAxis := row - 1
	yAxis := TitleToNumber(col)

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

// SetCellHyperLink provides function to set cell hyperlink by given worksheet
// name and link URL address. LinkType defines two types of hyperlink "External"
// for web site or "Location" for moving to one of cell in this workbook. The
// below is example for external link.
//
//    xlsx.SetCellHyperLink("Sheet1", "A3", "https://github.com/360EntSecGroup-Skylar/excelize", "External")
//    // Set underline and font color style for the cell.
//    style, _ := xlsx.NewStyle(`{"font":{"color":"#1265BE","underline":"single"}}`)
//    xlsx.SetCellStyle("Sheet1", "A3", "A3", style)
//
// A this is another example for "Location":
//
//    xlsx.SetCellHyperLink("Sheet1", "A3", "Sheet1!A40", "Location")
//
func (f *File) SetCellHyperLink(sheet, axis, link, linkType string) {
	xlsx := f.workSheetReader(sheet)
	axis = f.mergeCellsParser(xlsx, axis)
	linkTypes := map[string]xlsxHyperlink{
		"External": {},
		"Location": {Location: link},
	}
	hyperlink, ok := linkTypes[linkType]
	if !ok || axis == "" {
		return
	}
	hyperlink.Ref = axis
	if linkType == "External" {
		rID := f.addSheetRelationships(sheet, SourceRelationshipHyperLink, link, linkType)
		hyperlink.RID = "rId" + strconv.Itoa(rID)
	}
	if xlsx.Hyperlinks == nil {
		xlsx.Hyperlinks = &xlsxHyperlinks{}
	}
	xlsx.Hyperlinks.Hyperlink = append(xlsx.Hyperlinks.Hyperlink, hyperlink)
}

// GetCellHyperLink provides function to get cell hyperlink by given worksheet
// name and axis. Boolean type value link will be ture if the cell has a
// hyperlink and the target is the address of the hyperlink. Otherwise, the
// value of link will be false and the value of the target will be a blank
// string. For example get hyperlink of Sheet1!H6:
//
//    link, target := xlsx.GetCellHyperLink("Sheet1", "H6")
//
func (f *File) GetCellHyperLink(sheet, axis string) (bool, string) {
	var link bool
	var target string
	xlsx := f.workSheetReader(sheet)
	axis = f.mergeCellsParser(xlsx, axis)
	if xlsx.Hyperlinks == nil || axis == "" {
		return link, target
	}
	for h := range xlsx.Hyperlinks.Hyperlink {
		if xlsx.Hyperlinks.Hyperlink[h].Ref == axis {
			link = true
			target = xlsx.Hyperlinks.Hyperlink[h].Location
			if xlsx.Hyperlinks.Hyperlink[h].RID != "" {
				target = f.getSheetRelationshipsTargetByID(sheet, xlsx.Hyperlinks.Hyperlink[h].RID)
			}
		}
	}
	return link, target
}

// MergeCell provides function to merge cells by given coordinate area and sheet
// name. For example create a merged cell of D3:E9 on Sheet1:
//
//    xlsx.MergeCell("Sheet1", "D3", "E9")
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
	hxAxis := TitleToNumber(hcol)

	vcol := string(strings.Map(letterOnlyMapF, vcell))
	vrow, _ := strconv.Atoi(strings.Map(intOnlyMapF, vcell))
	vyAxis := vrow - 1
	vxAxis := TitleToNumber(vcol)

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
		mergeCell.Ref = ToAlphaString(hxAxis) + strconv.Itoa(hyAxis+1) + ":" + ToAlphaString(vxAxis) + strconv.Itoa(vyAxis+1)
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
		mergeCell.Ref = ToAlphaString(hxAxis) + strconv.Itoa(hyAxis+1) + ":" + ToAlphaString(vxAxis) + strconv.Itoa(vyAxis+1)
		mergeCells := xlsxMergeCells{}
		mergeCells.Cells = append(mergeCells.Cells, &mergeCell)
		xlsx.MergeCells = &mergeCells
	}
}

// SetCellInt provides function to set int type value of a cell by given
// worksheet name, cell coordinates and cell value.
func (f *File) SetCellInt(sheet, axis string, value int) {
	xlsx := f.workSheetReader(sheet)
	axis = f.mergeCellsParser(xlsx, axis)
	col := string(strings.Map(letterOnlyMapF, axis))
	row, err := strconv.Atoi(strings.Map(intOnlyMapF, axis))
	if err != nil {
		return
	}
	xAxis := row - 1
	yAxis := TitleToNumber(col)

	rows := xAxis + 1
	cell := yAxis + 1

	completeRow(xlsx, rows, cell)
	completeCol(xlsx, rows, cell)

	xlsx.SheetData.Row[xAxis].C[yAxis].S = f.prepareCellStyle(xlsx, cell, xlsx.SheetData.Row[xAxis].C[yAxis].S)
	xlsx.SheetData.Row[xAxis].C[yAxis].T = ""
	xlsx.SheetData.Row[xAxis].C[yAxis].V = strconv.Itoa(value)
}

// prepareCellStyle provides function to prepare style index of cell in
// worksheet by given column index and style index.
func (f *File) prepareCellStyle(xlsx *xlsxWorksheet, col, style int) int {
	if xlsx.Cols != nil && style == 0 {
		for _, v := range xlsx.Cols.Col {
			if v.Min <= col && col <= v.Max {
				style = v.Style
			}
		}
	}
	return style
}

// SetCellStr provides function to set string type value of a cell. Total number
// of characters that a cell can contain 32767 characters.
func (f *File) SetCellStr(sheet, axis, value string) {
	xlsx := f.workSheetReader(sheet)
	axis = f.mergeCellsParser(xlsx, axis)
	if len(value) > 32767 {
		value = value[0:32767]
	}
	col := string(strings.Map(letterOnlyMapF, axis))
	row, err := strconv.Atoi(strings.Map(intOnlyMapF, axis))
	if err != nil {
		return
	}
	xAxis := row - 1
	yAxis := TitleToNumber(col)

	rows := xAxis + 1
	cell := yAxis + 1

	completeRow(xlsx, rows, cell)
	completeCol(xlsx, rows, cell)

	// Leading space(s) character detection.
	if len(value) > 0 {
		if value[0] == 32 {
			xlsx.SheetData.Row[xAxis].C[yAxis].XMLSpace = xml.Attr{
				Name:  xml.Name{Space: NameSpaceXML, Local: "space"},
				Value: "preserve",
			}
		}
	}
	xlsx.SheetData.Row[xAxis].C[yAxis].S = f.prepareCellStyle(xlsx, cell, xlsx.SheetData.Row[xAxis].C[yAxis].S)
	xlsx.SheetData.Row[xAxis].C[yAxis].T = "str"
	xlsx.SheetData.Row[xAxis].C[yAxis].V = value
}

// SetCellDefault provides function to set string type value of a cell as
// default format without escaping the cell.
func (f *File) SetCellDefault(sheet, axis, value string) {
	xlsx := f.workSheetReader(sheet)
	axis = f.mergeCellsParser(xlsx, axis)
	col := string(strings.Map(letterOnlyMapF, axis))
	row, err := strconv.Atoi(strings.Map(intOnlyMapF, axis))
	if err != nil {
		return
	}
	xAxis := row - 1
	yAxis := TitleToNumber(col)

	rows := xAxis + 1
	cell := yAxis + 1

	completeRow(xlsx, rows, cell)
	completeCol(xlsx, rows, cell)

	xlsx.SheetData.Row[xAxis].C[yAxis].S = f.prepareCellStyle(xlsx, cell, xlsx.SheetData.Row[xAxis].C[yAxis].S)
	xlsx.SheetData.Row[xAxis].C[yAxis].T = ""
	xlsx.SheetData.Row[xAxis].C[yAxis].V = value
}

// checkCellInArea provides function to determine if a given coordinate is
// within an area.
func checkCellInArea(cell, area string) bool {
	result := false
	cell = strings.ToUpper(cell)
	col := string(strings.Map(letterOnlyMapF, cell))
	row, _ := strconv.Atoi(strings.Map(intOnlyMapF, cell))
	xAxis := row - 1
	yAxis := TitleToNumber(col)

	ref := strings.Split(area, ":")
	hCol := string(strings.Map(letterOnlyMapF, ref[0]))
	hRow, _ := strconv.Atoi(strings.Map(intOnlyMapF, ref[0]))
	hyAxis := hRow - 1
	hxAxis := TitleToNumber(hCol)

	vCol := string(strings.Map(letterOnlyMapF, ref[1]))
	vRow, _ := strconv.Atoi(strings.Map(intOnlyMapF, ref[1]))
	vyAxis := vRow - 1
	vxAxis := TitleToNumber(vCol)

	if hxAxis <= yAxis && yAxis <= vxAxis && hyAxis <= xAxis && xAxis <= vyAxis {
		result = true
	}

	return result
}
