package excelize

import (
	"encoding/json"
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
)

// AddComment provides the method to add comment in a sheet by given worksheet
// index, cell and format set (such as author and text). Note that the max
// author length is 255 and the max text length is 32512. For example, add a
// comment in Sheet1!$A$30:
//
//    err := f.AddButton("Sheet1", "A30", `{"author":"Excelize: ","text":"This is a comment."}`)
//
func (f *File) AddButton(sheet, cell, format string) error {
	formatSet, err := parseFormatButtonSet(format)
	fmt.Printf("%+v\n", formatSet)
	if err != nil {
		return err
	}
	// Read sheet data.
	ws, err := f.workSheetReader(sheet)
	if err != nil {
		return err
	}
	buttonID := f.countButtons() + 1
	fmt.Println(buttonID)
	drawingVML := "xl/drawings/vmlDrawing" + strconv.Itoa(buttonID) + ".vml"
	fmt.Println(drawingVML)

	sheetRelationshipsDrawingVML := "../drawings/vmlDrawing" + strconv.Itoa(buttonID) + ".vml"
	if ws.LegacyDrawing != nil {
		// The worksheet already has a comments relationships, use the relationships drawing ../drawings/vmlDrawing%d.vml.
		sheetRelationshipsDrawingVML = f.getSheetRelationshipsTargetByID(sheet, ws.LegacyDrawing.RID)
		buttonID, _ = strconv.Atoi(strings.TrimSuffix(strings.TrimPrefix(sheetRelationshipsDrawingVML, "../drawings/vmlDrawing"), ".vml"))
		drawingVML = strings.Replace(sheetRelationshipsDrawingVML, "..", "xl", -1)
	} else {
		// Add first comment for given sheet.
		sheetRels := "xl/worksheets/_rels/" + strings.TrimPrefix(f.sheetMap[trimSheetName(sheet)], "xl/worksheets/") + ".rels"
		rID := f.addRels(sheetRels, SourceRelationshipDrawingVML, sheetRelationshipsDrawingVML, "")
		// f.addRels(sheetRels, SourceRelationshipComments, sheetRelationshipsComments, "")
		f.addSheetNameSpace(sheet, SourceRelationship)
		f.addSheetLegacyDrawing(sheet, rID)
	}
	// ctrlProps
	err = f.addDrawingVMLButton(buttonID, drawingVML, cell, formatSet)
	if err != nil {
		return err
	}

	// commentsXML := "xl/comments" + strconv.Itoa(buttonID) + ".xml"
	// f.addComment(commentsXML, cell, formatSet)

	f.addContentTypePart(buttonID, "comments")

	return err
}

// parseFormatCommentsSet provides a function to parse the format settings of
// the comment with default value.
func parseFormatButtonSet(formatSet string) (*formatButton, error) {
	format := formatButton{
		Caption: "Button 1",
		Width:   160,
		Height:  160,
	}
	err := json.Unmarshal([]byte(formatSet), &format)
	return &format, err
}

// countComments provides a function to get comments files count storage in
// the folder xl.
func (f *File) countButtons() int {
	// TODO: implement logic
	return 0
}

// addDrawingVML provides a function to create comment as
// xl/drawings/vmlDrawing%d.vml by given commit ID and cell.
func (f *File) addDrawingVMLButton(buttonID int, drawingVML, cell string, formatSet *formatButton) error {
	col, row, err := CellNameToCoordinates(cell)
	if err != nil {
		return err
	}
	colIdx := col - 1
	rowIdx := row - 1

	width := int(float64(formatSet.Width) * formatSet.ScaleX)
	widthStr := strconv.Itoa(width)
	height := int(float64(formatSet.Height) * formatSet.ScaleY)
	heightStr := strconv.Itoa(height)

	// work on position parse
	// colStart, rowStart, colEnd, rowEnd, x2, y2 := f.positionObjectPixels(sheet, colIdx, rowIdx, formatSet.OffsetX, formatSet.OffsetY, width, height)

	vml := f.VMLDrawing[drawingVML]
	if vml == nil {
		vml = &vmlDrawing{
			XMLNSv:  "urn:schemas-microsoft-com:vml",
			XMLNSo:  "urn:schemas-microsoft-com:office:office",
			XMLNSx:  "urn:schemas-microsoft-com:office:excel",
			XMLNSmv: "http://macVmlSchemaUri", // remove this?
			Shapelayout: &xlsxShapelayout{
				Ext: "edit",
				IDmap: &xlsxIDmap{
					Ext:  "edit",
					Data: buttonID,
				},
			},
			Shapetype: &xlsxShapetype{
				ID:        "_x0000_t201",
				Coordsize: "21600,21600",
				Spt:       201,
				Path:      "m,l,21600r21600,l21600,xe",
				Stroke: &xlsxStroke{
					Joinstyle: "miter",
				},
				VPath: &vPath{
					Gradientshapeok: "t", // not used in button
					Connecttype:     "rect",
					// missing: Shadowok: "f", Extrusionok: "f", Strokeok: "f", Fillok: "f"
				},
				// missing: &Lock {Ext, Shapetype}
			},
		}
	}
	sp := encodeShapeButton{
		Fill: &vFillButton{
			Color2:           "buttonFace [67]",
			Detectmouseclick: "t",
		},
		Lock: &oLockButton{
			Ext:      "edit",
			Rotation: "t",
		},
		TextBox: &vTextboxButton{
			Style:       "mso-direction-alt:auto",
			Singleclick: "f",
			Div: &xlsxDivButton{
				Style: "text-align:center",
				Font: &fontButton{
					Face:    "Calibri",
					Size:    "220",
					Color:   "#000000",
					Caption: formatSet.Caption,
				},
			},
		},
		ClientData: &xClientDataButton{
			ObjectType: "Button",
			Anchor: fmt.Sprintf(
				"%d, 0, %d, 0, %d, %d, %d, 10",
				1+colIdx, 1+rowIdx, 2+colIdx, colIdx, 2+rowIdx),
			PrintObject: "False",
			AutoFill:    "False",
			FmlaMacro:   "[0]!" + formatSet.Macro,
			TextHAlign:  "Center",
			TextVAlign:  "Center",
		},
	}
	s, _ := xml.Marshal(sp)

	shape := xlsxShape{
		ID:          "_x0000_s1025",
		Type:        "#_x0000_t201",
		Style:       "position:absolute;margin-left:0pt;margin-top:0pt;width:" + widthStr + "pt;height:" + heightStr + "pt;z-index:1;mso-wrap-style:tight",
		Button:      "t",
		Fillcolor:   "buttonFace [67]",
		Strokecolor: "windowText [64]",
		Insetmode:   "auto",
		Val:         string(s[19 : len(s)-20]),
	}
	d := f.decodeVMLDrawingReader(drawingVML)
	if d != nil {
		for _, v := range d.Shape {
			s := xlsxShape{
				ID:          "_x0000_s1025",
				Type:        "#_x0000_t201",
				Style:       "position:absolute;margin-left:0pt;margin-top:0pt;width:60pt;height:22.5pt;z-index:1;mso-wrap-style:tight",
				Fillcolor:   "buttonFace [67]",
				Strokecolor: "windowText [64]",
				Insetmode:   "auto",
				Val:         v.Val,
			}
			vml.Shape = append(vml.Shape, s)
		}
	}
	vml.Shape = append(vml.Shape, shape)
	f.VMLDrawing[drawingVML] = vml
	return err
}
