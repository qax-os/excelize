package excelize

import (
	"encoding/json"
	"encoding/xml"
	"strconv"
	"strings"
)

// parseFormatBordersSet provides function to parse the format settings of the
// borders.
func parseFormatBordersSet(bordersSet string) (*formatBorder, error) {
	var format formatBorder
	err := json.Unmarshal([]byte(bordersSet), &format)
	return &format, err
}

// SetBorder provides function to get value from cell by given sheet index and
// coordinate area in XLSX file. Note that the color field uses RGB color code
// and diagonalDown and diagonalUp type border should be use same color in the
// same coordinate area.
//
// For example create a borders of cell H9 on
// Sheet1:
//
//    err := xlsx.SetBorder("Sheet1", "H9", "H9", `{"border":[{"type":"left","color":"0000FF","style":3},{"type":"top","color":"00FF00","style":4},{"type":"bottom","color":"FFFF00","style":5},{"type":"right","color":"FF0000","style":6},{"type":"diagonalDown","color":"A020F0","style":7},{"type":"diagonalUp","color":"A020F0","style":8}]}`)
//    if err != nil {
//        fmt.Println(err)
//    }
//
// The following shows the border styles sorted by excelize index number:
//
//    +-------+---------------+--------+-----------------+
//    | Index | Name          | Weight | Style           |
//    +=======+===============+========+=================+
//    | 0     | None          | 0      |                 |
//    +-------+---------------+--------+-----------------+
//    | 1     | Continuous    | 1      | ``-----------`` |
//    +-------+---------------+--------+-----------------+
//    | 2     | Continuous    | 2      | ``-----------`` |
//    +-------+---------------+--------+-----------------+
//    | 3     | Dash          | 1      | ``- - - - - -`` |
//    +-------+---------------+--------+-----------------+
//    | 4     | Dot           | 1      | ``. . . . . .`` |
//    +-------+---------------+--------+-----------------+
//    | 5     | Continuous    | 3      | ``-----------`` |
//    +-------+---------------+--------+-----------------+
//    | 6     | Double        | 3      | ``===========`` |
//    +-------+---------------+--------+-----------------+
//    | 7     | Continuous    | 0      | ``-----------`` |
//    +-------+---------------+--------+-----------------+
//    | 8     | Dash          | 2      | ``- - - - - -`` |
//    +-------+---------------+--------+-----------------+
//    | 9     | Dash Dot      | 1      | ``- . - . - .`` |
//    +-------+---------------+--------+-----------------+
//    | 10    | Dash Dot      | 2      | ``- . - . - .`` |
//    +-------+---------------+--------+-----------------+
//    | 11    | Dash Dot Dot  | 1      | ``- . . - . .`` |
//    +-------+---------------+--------+-----------------+
//    | 12    | Dash Dot Dot  | 2      | ``- . . - . .`` |
//    +-------+---------------+--------+-----------------+
//    | 13    | SlantDash Dot | 2      | ``/ - . / - .`` |
//    +-------+---------------+--------+-----------------+
//
// The following shows the borders in the order shown in the Excel dialog:
//
//    +-------+-----------------+-------+-----------------+
//    | Index | Style           | Index | Style           |
//    +=======+=================+=======+=================+
//    | 0     | None            | 12    | ``- . . - . .`` |
//    +-------+-----------------+-------+-----------------+
//    | 7     | ``-----------`` | 13    | ``/ - . / - .`` |
//    +-------+-----------------+-------+-----------------+
//    | 4     | ``. . . . . .`` | 10    | ``- . - . - .`` |
//    +-------+-----------------+-------+-----------------+
//    | 11    | ``- . . - . .`` | 8     | ``- - - - - -`` |
//    +-------+-----------------+-------+-----------------+
//    | 9     | ``- . - . - .`` | 2     | ``-----------`` |
//    +-------+-----------------+-------+-----------------+
//    | 3     | ``- - - - - -`` | 5     | ``-----------`` |
//    +-------+-----------------+-------+-----------------+
//    | 1     | ``-----------`` | 6     | ``===========`` |
//    +-------+-----------------+-------+-----------------+
//
func (f *File) SetBorder(sheet, hcell, vcell, style string) error {
	var styleSheet xlsxStyleSheet
	xml.Unmarshal([]byte(f.readXML("xl/styles.xml")), &styleSheet)
	formatBorder, err := parseFormatBordersSet(style)
	if err != nil {
		return err
	}
	borderID := setBorders(&styleSheet, formatBorder)
	cellXfsID := setCellXfs(&styleSheet, borderID)
	output, err := xml.Marshal(styleSheet)
	if err != nil {
		return err
	}
	f.saveFileList("xl/styles.xml", replaceWorkSheetsRelationshipsNameSpace(string(output)))
	f.setCellStyle(sheet, hcell, vcell, cellXfsID)
	return err
}

// setBorders provides function to add border elements in the styles.xml by
// given borders format settings.
func setBorders(style *xlsxStyleSheet, formatBorder *formatBorder) int {
	var styles = []string{
		"none",
		"thin",
		"medium",
		"dashed",
		"dotted",
		"thick",
		"double",
		"hair",
		"mediumDashed",
		"dashDot",
		"mediumDashDot",
		"dashDotDot",
		"mediumDashDotDot",
		"slantDashDot",
	}

	var border xlsxBorder
	for _, v := range formatBorder.Border {
		if v.Style > 13 || v.Style < 0 {
			continue
		}
		var color xlsxColor
		color.RGB = v.Color
		switch v.Type {
		case "left":
			border.Left.Style = styles[v.Style]
			border.Left.Color = &color
		case "right":
			border.Right.Style = styles[v.Style]
			border.Right.Color = &color
		case "top":
			border.Top.Style = styles[v.Style]
			border.Top.Color = &color
		case "bottom":
			border.Bottom.Style = styles[v.Style]
			border.Bottom.Color = &color
		case "diagonalUp":
			border.Diagonal.Style = styles[v.Style]
			border.Diagonal.Color = &color
			border.DiagonalUp = true
		case "diagonalDown":
			border.Diagonal.Style = styles[v.Style]
			border.Diagonal.Color = &color
			border.DiagonalDown = true
		}
	}
	style.Borders.Count++
	style.Borders.Border = append(style.Borders.Border, &border)
	return style.Borders.Count - 1
}

// setCellXfs provides function to set describes all of the formatting for a
// cell.
func setCellXfs(style *xlsxStyleSheet, borderID int) int {
	var xf xlsxXf
	xf.BorderID = borderID
	style.CellXfs.Count++
	style.CellXfs.Xf = append(style.CellXfs.Xf, xf)
	return style.CellXfs.Count - 1
}

// setCellStyle provides function to add style attribute for cells by given
// sheet index, coordinate area and style ID.
func (f *File) setCellStyle(sheet, hcell, vcell string, styleID int) {
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

	// Correct the coordinate area, such correct C1:B3 to B1:C3.
	hcell = toAlphaString(hxAxis+1) + strconv.Itoa(hyAxis+1)
	vcell = toAlphaString(vxAxis+1) + strconv.Itoa(vyAxis+1)

	var xlsx xlsxWorksheet
	name := "xl/worksheets/" + strings.ToLower(sheet) + ".xml"
	xml.Unmarshal([]byte(f.readXML(name)), &xlsx)
	if f.checked == nil {
		f.checked = make(map[string]bool)
	}
	ok := f.checked[name]
	if !ok {
		xlsx = checkRow(xlsx)
		f.checked[name] = true
	}

	xlsx = completeRow(xlsx, vxAxis+1, vyAxis+1)
	xlsx = completeCol(xlsx, vxAxis+1, vyAxis+1)

	for r, row := range xlsx.SheetData.Row {
		for k, c := range row.C {
			if checkCellInArea(c.R, hcell+":"+vcell) {
				xlsx.SheetData.Row[r].C[k].S = styleID
			}
		}
	}
	output, _ := xml.Marshal(xlsx)
	f.saveFileList(name, replaceWorkSheetsRelationshipsNameSpace(string(output)))
}
