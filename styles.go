package excelize

import (
	"encoding/json"
	"encoding/xml"
	"strconv"
	"strings"
)

// parseFormatStyleSet provides function to parse the format settings of the
// borders.
func parseFormatStyleSet(style string) (*formatCellStyle, error) {
	var format formatCellStyle
	err := json.Unmarshal([]byte(style), &format)
	return &format, err
}

// SetCellStyle provides function to get value from cell by given sheet index
// and coordinate area in XLSX file. Note that the color field uses RGB color
// code and diagonalDown and diagonalUp type border should be use same color in
// the same coordinate area.
//
// For example create a borders of cell H9 on Sheet1:
//
//    err := xlsx.SetBorder("Sheet1", "H9", "H9", `{"border":[{"type":"left","color":"0000FF","style":3},{"type":"top","color":"00FF00","style":4},{"type":"bottom","color":"FFFF00","style":5},{"type":"right","color":"FF0000","style":6},{"type":"diagonalDown","color":"A020F0","style":7},{"type":"diagonalUp","color":"A020F0","style":8}]}`)
//    if err != nil {
//        fmt.Println(err)
//    }
//
// Set gradient fill with vertical variants shading styles for cell H9 on
// Sheet1:
//
//    err := xlsx.SetBorder("Sheet1", "H9", "H9", `{"fill":[{"type":"gradient","color":["#FFFFFF","#E0EBF5"],"shading":1}]}`)
//    if err != nil {
//        fmt.Println(err)
//    }
//
// Set solid style pattern fill for cell H9 on Sheet1:
//
//    err := xlsx.SetBorder("Sheet1", "H9", "H9", `{"fill":[{"type":"pattern","color":["#E0EBF5"],"pattern":1}]}`)
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
// The following shows the shading styles sorted by excelize index number:
//
//    +-------+-----------------+-------+-----------------+
//    | Index | Style           | Index | Style           |
//    +=======+=================+=======+=================+
//    | 0     | Horizontal      | 3     | Diagonal down   |
//    +-------+-----------------+-------+-----------------+
//    | 1     | Vertical        | 4     | From corner     |
//    +-------+-----------------+-------+-----------------+
//    | 2     | Diagonal Up     | 5     | From center     |
//    +-------+-----------------+-------+-----------------+
//
// The following shows the patterns styles sorted by excelize index number:
//
//    +-------+-----------------+-------+-----------------+
//    | Index | Style           | Index | Style           |
//    +=======+=================+=======+=================+
//    | 0     | None            | 10    | darkTrellis     |
//    +-------+-----------------+-------+-----------------+
//    | 1     | solid           | 11    | lightHorizontal |
//    +-------+-----------------+-------+-----------------+
//    | 2     | mediumGray      | 12    | lightVertical   |
//    +-------+-----------------+-------+-----------------+
//    | 3     | darkGray        | 13    | lightDown       |
//    +-------+-----------------+-------+-----------------+
//    | 4     | lightGray       | 14    | lightUp         |
//    +-------+-----------------+-------+-----------------+
//    | 5     | darkHorizontal  | 15    | lightGrid       |
//    +-------+-----------------+-------+-----------------+
//    | 6     | darkVertical    | 16    | lightTrellis    |
//    +-------+-----------------+-------+-----------------+
//    | 7     | darkDown        | 17    | gray125         |
//    +-------+-----------------+-------+-----------------+
//    | 8     | darkUp          | 18    | gray0625        |
//    +-------+-----------------+-------+-----------------+
//    | 9     | darkGrid        |       |                 |
//    +-------+-----------------+-------+-----------------+
//
func (f *File) SetCellStyle(sheet, hcell, vcell, style string) error {
	var styleSheet xlsxStyleSheet
	xml.Unmarshal([]byte(f.readXML("xl/styles.xml")), &styleSheet)
	formatCellStyle, err := parseFormatStyleSet(style)
	if err != nil {
		return err
	}
	borderID := setBorders(&styleSheet, formatCellStyle)
	fillID := setFills(&styleSheet, formatCellStyle)
	cellXfsID := setCellXfs(&styleSheet, fillID, borderID)
	output, err := xml.Marshal(styleSheet)
	if err != nil {
		return err
	}
	f.saveFileList("xl/styles.xml", replaceWorkSheetsRelationshipsNameSpace(string(output)))
	f.setCellStyle(sheet, hcell, vcell, cellXfsID)
	return err
}

// setFills provides function to add fill elements in the styles.xml by given
// cell format settings.
func setFills(style *xlsxStyleSheet, formatCellStyle *formatCellStyle) int {
	var patterns = []string{
		"none",
		"solid",
		"mediumGray",
		"darkGray",
		"lightGray",
		"darkHorizontal",
		"darkVertical",
		"darkDown",
		"darkUp",
		"darkGrid",
		"darkTrellis",
		"lightHorizontal",
		"lightVertical",
		"lightDown",
		"lightUp",
		"lightGrid",
		"lightTrellis",
		"gray125",
		"gray0625",
	}

	var variants = []float64{
		90,
		0,
		45,
		135,
	}

	var fill xlsxFill
	for _, v := range formatCellStyle.Fill {
		switch v.Type {
		case "gradient":
			if len(v.Color) != 2 {
				continue
			}
			var gradient xlsxGradientFill
			switch v.Shading {
			case 0, 1, 2, 3:
				gradient.Degree = variants[v.Shading]
			case 4:
				gradient.Type = "path"
			case 5:
				gradient.Type = "path"
				gradient.Bottom = 0.5
				gradient.Left = 0.5
				gradient.Right = 0.5
				gradient.Top = 0.5
			default:
				continue
			}
			var stops []*xlsxGradientFillStop
			for index, color := range v.Color {
				var stop xlsxGradientFillStop
				stop.Position = float64(index)
				stop.Color.RGB = getPaletteColor(color)
				stops = append(stops, &stop)
			}
			gradient.Stop = stops
			fill.GradientFill = &gradient
		case "pattern":
			if v.Pattern > 18 || v.Pattern < 0 {
				continue
			}
			if len(v.Color) < 1 {
				continue
			}
			var pattern xlsxPatternFill
			pattern.PatternType = patterns[v.Pattern]
			pattern.FgColor.RGB = getPaletteColor(v.Color[0])
			fill.PatternFill = &pattern
		}
	}
	style.Fills.Count++
	style.Fills.Fill = append(style.Fills.Fill, &fill)
	return style.Fills.Count - 1
}

// setBorders provides function to add border elements in the styles.xml by
// given borders format settings.
func setBorders(style *xlsxStyleSheet, formatCellStyle *formatCellStyle) int {
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
	for _, v := range formatCellStyle.Border {
		if v.Style > 13 || v.Style < 0 {
			continue
		}
		var color xlsxColor
		color.RGB = getPaletteColor(v.Color)
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
func setCellXfs(style *xlsxStyleSheet, fillID, borderID int) int {
	var xf xlsxXf
	xf.FillID = fillID
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

	xlsx := f.workSheetReader(sheet)

	completeRow(xlsx, vxAxis+1, vyAxis+1)
	completeCol(xlsx, vxAxis+1, vyAxis+1)

	for r, row := range xlsx.SheetData.Row {
		for k, c := range row.C {
			if checkCellInArea(c.R, hcell+":"+vcell) {
				xlsx.SheetData.Row[r].C[k].S = styleID
			}
		}
	}
}

// getPaletteColor provides function to convert the RBG color by given string.
func getPaletteColor(color string) string {
	return "FF" + strings.Replace(strings.ToUpper(color), "#", "", -1)
}
