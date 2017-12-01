package excelize

import (
	"bytes"
	"encoding/json"
	"encoding/xml"
	"errors"
	"image"
	"io/ioutil"
	"os"
	"path"
	"path/filepath"
	"strconv"
	"strings"
)

// parseFormatPictureSet provides function to parse the format settings of the
// picture with default value.
func parseFormatPictureSet(formatSet string) *formatPicture {
	format := formatPicture{
		FPrintsWithSheet: true,
		FLocksWithSheet:  false,
		NoChangeAspect:   false,
		OffsetX:          0,
		OffsetY:          0,
		XScale:           1.0,
		YScale:           1.0,
	}
	json.Unmarshal([]byte(formatSet), &format)
	return &format
}

// AddPicture provides the method to add picture in a sheet by given picture
// format set (such as offset, scale, aspect ratio setting and print settings)
// and file path. For example:
//
//    package main
//
//    import (
//        "fmt"
//        _ "image/gif"
//        _ "image/jpeg"
//        _ "image/png"
//
//        "github.com/360EntSecGroup-Skylar/excelize"
//    )
//
//    func main() {
//        xlsx := excelize.NewFile()
//        // Insert a picture.
//        err := xlsx.AddPicture("Sheet1", "A2", "./image1.jpg", "")
//        if err != nil {
//            fmt.Println(err)
//        }
//        // Insert a picture to sheet with scaling.
//        err = xlsx.AddPicture("Sheet1", "D2", "./image1.png", `{"x_scale": 0.5, "y_scale": 0.5}`)
//        if err != nil {
//            fmt.Println(err)
//        }
//        // Insert a picture offset in the cell with printing support.
//        err = xlsx.AddPicture("Sheet1", "H2", "./image3.gif", `{"x_offset": 15, "y_offset": 10, "print_obj": true, "lock_aspect_ratio": false, "locked": false}`)
//        if err != nil {
//            fmt.Println(err)
//        }
//        err = xlsx.SaveAs("./Workbook.xlsx")
//        if err != nil {
//            fmt.Println(err)
//        }
//    }
//
func (f *File) AddPicture(sheet, cell, picture, format string) error {
	var err error
	// Check picture exists first.
	if _, err = os.Stat(picture); os.IsNotExist(err) {
		return err
	}
	ext, ok := supportImageTypes[path.Ext(picture)]
	if !ok {
		return errors.New("Unsupported image extension")
	}
	readFile, _ := os.Open(picture)
	image, _, err := image.DecodeConfig(readFile)
	_, file := filepath.Split(picture)
	formatSet := parseFormatPictureSet(format)
	// Read sheet data.
	xlsx := f.workSheetReader(sheet)
	// Add first picture for given sheet, create xl/drawings/ and xl/drawings/_rels/ folder.
	drawingID := f.countDrawings() + 1
	pictureID := f.countMedia() + 1
	drawingXML := "xl/drawings/drawing" + strconv.Itoa(drawingID) + ".xml"
	drawingID, drawingXML = f.prepareDrawing(xlsx, drawingID, sheet, drawingXML)
	drawingRID := f.addDrawingRelationships(drawingID, SourceRelationshipImage, "../media/image"+strconv.Itoa(pictureID)+ext)
	f.addDrawingPicture(sheet, drawingXML, cell, file, image.Width, image.Height, drawingRID, formatSet)
	f.addMedia(picture, ext)
	f.addContentTypePart(drawingID, "drawings")
	return err
}

// addSheetRelationships provides function to add
// xl/worksheets/_rels/sheet%d.xml.rels by given worksheet name, relationship
// type and target.
func (f *File) addSheetRelationships(sheet, relType, target, targetMode string) int {
	name, ok := f.sheetMap[trimSheetName(sheet)]
	if !ok {
		name = strings.ToLower(sheet) + ".xml"
	}
	var rels = "xl/worksheets/_rels/" + strings.TrimPrefix(name, "xl/worksheets/") + ".rels"
	var sheetRels xlsxWorkbookRels
	var rID = 1
	var ID bytes.Buffer
	ID.WriteString("rId")
	ID.WriteString(strconv.Itoa(rID))
	_, ok = f.XLSX[rels]
	if ok {
		ID.Reset()
		xml.Unmarshal([]byte(f.readXML(rels)), &sheetRels)
		rID = len(sheetRels.Relationships) + 1
		ID.WriteString("rId")
		ID.WriteString(strconv.Itoa(rID))
	}
	sheetRels.Relationships = append(sheetRels.Relationships, xlsxWorkbookRelation{
		ID:         ID.String(),
		Type:       relType,
		Target:     target,
		TargetMode: targetMode,
	})
	output, _ := xml.Marshal(sheetRels)
	f.saveFileList(rels, string(output))
	return rID
}

// deleteSheetRelationships provides function to delete relationships in
// xl/worksheets/_rels/sheet%d.xml.rels by given worksheet name and relationship
// index.
func (f *File) deleteSheetRelationships(sheet, rID string) {
	name, ok := f.sheetMap[trimSheetName(sheet)]
	if !ok {
		name = strings.ToLower(sheet) + ".xml"
	}
	var rels = "xl/worksheets/_rels/" + strings.TrimPrefix(name, "xl/worksheets/") + ".rels"
	var sheetRels xlsxWorkbookRels
	xml.Unmarshal([]byte(f.readXML(rels)), &sheetRels)
	for k, v := range sheetRels.Relationships {
		if v.ID == rID {
			sheetRels.Relationships = append(sheetRels.Relationships[:k], sheetRels.Relationships[k+1:]...)
		}
	}
	output, _ := xml.Marshal(sheetRels)
	f.saveFileList(rels, string(output))
}

// addSheetLegacyDrawing provides function to add legacy drawing element to
// xl/worksheets/sheet%d.xml by given worksheet name and relationship index.
func (f *File) addSheetLegacyDrawing(sheet string, rID int) {
	xlsx := f.workSheetReader(sheet)
	xlsx.LegacyDrawing = &xlsxLegacyDrawing{
		RID: "rId" + strconv.Itoa(rID),
	}
}

// addSheetDrawing provides function to add drawing element to
// xl/worksheets/sheet%d.xml by given worksheet name and relationship index.
func (f *File) addSheetDrawing(sheet string, rID int) {
	xlsx := f.workSheetReader(sheet)
	xlsx.Drawing = &xlsxDrawing{
		RID: "rId" + strconv.Itoa(rID),
	}
}

// addSheetPicture provides function to add picture element to
// xl/worksheets/sheet%d.xml by given worksheet name and relationship index.
func (f *File) addSheetPicture(sheet string, rID int) {
	xlsx := f.workSheetReader(sheet)
	xlsx.Picture = &xlsxPicture{
		RID: "rId" + strconv.Itoa(rID),
	}
}

// countDrawings provides function to get drawing files count storage in the
// folder xl/drawings.
func (f *File) countDrawings() int {
	count := 0
	for k := range f.XLSX {
		if strings.Contains(k, "xl/drawings/drawing") {
			count++
		}
	}
	return count
}

// addDrawingPicture provides function to add picture by given sheet,
// drawingXML, cell, file name, width, height relationship index and format
// sets.
func (f *File) addDrawingPicture(sheet, drawingXML, cell, file string, width, height, rID int, formatSet *formatPicture) {
	cell = strings.ToUpper(cell)
	fromCol := string(strings.Map(letterOnlyMapF, cell))
	fromRow, _ := strconv.Atoi(strings.Map(intOnlyMapF, cell))
	row := fromRow - 1
	col := TitleToNumber(fromCol)
	width = int(float64(width) * formatSet.XScale)
	height = int(float64(height) * formatSet.YScale)
	colStart, rowStart, _, _, colEnd, rowEnd, x2, y2 := f.positionObjectPixels(sheet, col, row, formatSet.OffsetX, formatSet.OffsetY, width, height)
	content := xlsxWsDr{}
	content.A = NameSpaceDrawingML
	content.Xdr = NameSpaceDrawingMLSpreadSheet
	cNvPrID := f.drawingParser(drawingXML, &content)
	twoCellAnchor := xdrCellAnchor{}
	twoCellAnchor.EditAs = "oneCell"
	from := xlsxFrom{}
	from.Col = colStart
	from.ColOff = formatSet.OffsetX * EMU
	from.Row = rowStart
	from.RowOff = formatSet.OffsetY * EMU
	to := xlsxTo{}
	to.Col = colEnd
	to.ColOff = x2 * EMU
	to.Row = rowEnd
	to.RowOff = y2 * EMU
	twoCellAnchor.From = &from
	twoCellAnchor.To = &to
	pic := xlsxPic{}
	pic.NvPicPr.CNvPicPr.PicLocks.NoChangeAspect = formatSet.NoChangeAspect
	pic.NvPicPr.CNvPr.ID = f.countCharts() + f.countMedia() + 1
	pic.NvPicPr.CNvPr.Descr = file
	pic.NvPicPr.CNvPr.Name = "Picture " + strconv.Itoa(cNvPrID)
	pic.BlipFill.Blip.R = SourceRelationship
	pic.BlipFill.Blip.Embed = "rId" + strconv.Itoa(rID)
	pic.SpPr.PrstGeom.Prst = "rect"

	twoCellAnchor.Pic = &pic
	twoCellAnchor.ClientData = &xdrClientData{
		FLocksWithSheet:  formatSet.FLocksWithSheet,
		FPrintsWithSheet: formatSet.FPrintsWithSheet,
	}
	content.TwoCellAnchor = append(content.TwoCellAnchor, &twoCellAnchor)
	output, _ := xml.Marshal(content)
	f.saveFileList(drawingXML, string(output))
}

// addDrawingRelationships provides function to add image part relationships in
// the file xl/drawings/_rels/drawing%d.xml.rels by given drawing index,
// relationship type and target.
func (f *File) addDrawingRelationships(index int, relType, target string) int {
	var rels = "xl/drawings/_rels/drawing" + strconv.Itoa(index) + ".xml.rels"
	var drawingRels xlsxWorkbookRels
	var rID = 1
	var ID bytes.Buffer
	ID.WriteString("rId")
	ID.WriteString(strconv.Itoa(rID))
	_, ok := f.XLSX[rels]
	if ok {
		ID.Reset()
		xml.Unmarshal([]byte(f.readXML(rels)), &drawingRels)
		rID = len(drawingRels.Relationships) + 1
		ID.WriteString("rId")
		ID.WriteString(strconv.Itoa(rID))
	}
	drawingRels.Relationships = append(drawingRels.Relationships, xlsxWorkbookRelation{
		ID:     ID.String(),
		Type:   relType,
		Target: target,
	})
	output, _ := xml.Marshal(drawingRels)
	f.saveFileList(rels, string(output))
	return rID
}

// countMedia provides function to get media files count storage in the folder
// xl/media/image.
func (f *File) countMedia() int {
	count := 0
	for k := range f.XLSX {
		if strings.Contains(k, "xl/media/image") {
			count++
		}
	}
	return count
}

// addMedia provides function to add picture into folder xl/media/image by given
// file name and extension name.
func (f *File) addMedia(file, ext string) {
	count := f.countMedia()
	dat, _ := ioutil.ReadFile(file)
	media := "xl/media/image" + strconv.Itoa(count+1) + ext
	f.XLSX[media] = string(dat)
}

// setContentTypePartImageExtensions provides function to set the content type
// for relationship parts and the Main Document part.
func (f *File) setContentTypePartImageExtensions() {
	var imageTypes = map[string]bool{"jpeg": false, "png": false, "gif": false}
	content := f.contentTypesReader()
	for _, v := range content.Defaults {
		_, ok := imageTypes[v.Extension]
		if ok {
			imageTypes[v.Extension] = true
		}
	}
	for k, v := range imageTypes {
		if !v {
			content.Defaults = append(content.Defaults, xlsxDefault{
				Extension:   k,
				ContentType: "image/" + k,
			})
		}
	}
}

// setContentTypePartVMLExtensions provides function to set the content type
// for relationship parts and the Main Document part.
func (f *File) setContentTypePartVMLExtensions() {
	vml := false
	content := f.contentTypesReader()
	for _, v := range content.Defaults {
		if v.Extension == "vml" {
			vml = true
		}
	}
	if !vml {
		content.Defaults = append(content.Defaults, xlsxDefault{
			Extension:   "vml",
			ContentType: "application/vnd.openxmlformats-officedocument.vmlDrawing",
		})
	}
}

// addContentTypePart provides function to add content type part relationships
// in the file [Content_Types].xml by given index.
func (f *File) addContentTypePart(index int, contentType string) {
	setContentType := map[string]func(){
		"comments": f.setContentTypePartVMLExtensions,
		"drawings": f.setContentTypePartImageExtensions,
	}
	partNames := map[string]string{
		"chart":    "/xl/charts/chart" + strconv.Itoa(index) + ".xml",
		"comments": "/xl/comments" + strconv.Itoa(index) + ".xml",
		"drawings": "/xl/drawings/drawing" + strconv.Itoa(index) + ".xml",
		"table":    "/xl/tables/table" + strconv.Itoa(index) + ".xml",
	}
	contentTypes := map[string]string{
		"chart":    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
		"comments": "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
		"drawings": "application/vnd.openxmlformats-officedocument.drawing+xml",
		"table":    "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
	}
	s, ok := setContentType[contentType]
	if ok {
		s()
	}
	content := f.contentTypesReader()
	for _, v := range content.Overrides {
		if v.PartName == partNames[contentType] {
			return
		}
	}
	content.Overrides = append(content.Overrides, xlsxOverride{
		PartName:    partNames[contentType],
		ContentType: contentTypes[contentType],
	})
}

// getSheetRelationshipsTargetByID provides function to get Target attribute
// value in xl/worksheets/_rels/sheet%d.xml.rels by given worksheet name and
// relationship index.
func (f *File) getSheetRelationshipsTargetByID(sheet, rID string) string {
	name, ok := f.sheetMap[trimSheetName(sheet)]
	if !ok {
		name = strings.ToLower(sheet) + ".xml"
	}
	var rels = "xl/worksheets/_rels/" + strings.TrimPrefix(name, "xl/worksheets/") + ".rels"
	var sheetRels xlsxWorkbookRels
	xml.Unmarshal([]byte(f.readXML(rels)), &sheetRels)
	for _, v := range sheetRels.Relationships {
		if v.ID == rID {
			return v.Target
		}
	}
	return ""
}

// GetPicture provides function to get picture base name and raw content embed
// in XLSX by given worksheet and cell name. This function returns the file name
// in XLSX and file contents as []byte data types. For example:
//
//    xlsx, err := excelize.OpenFile("./Workbook.xlsx")
//    if err != nil {
//        fmt.Println(err)
//        return
//    }
//    file, raw := xlsx.GetPicture("Sheet1", "A2")
//    if file == "" {
//        return
//    }
//    err := ioutil.WriteFile(file, raw, 0644)
//    if err != nil {
//        fmt.Println(err)
//    }
//
func (f *File) GetPicture(sheet, cell string) (string, []byte) {
	xlsx := f.workSheetReader(sheet)
	if xlsx.Drawing == nil {
		return "", []byte{}
	}
	target := f.getSheetRelationshipsTargetByID(sheet, xlsx.Drawing.RID)
	drawingXML := strings.Replace(target, "..", "xl", -1)

	_, ok := f.XLSX[drawingXML]
	if !ok {
		return "", []byte{}
	}
	decodeWsDr := decodeWsDr{}
	xml.Unmarshal([]byte(f.readXML(drawingXML)), &decodeWsDr)

	cell = strings.ToUpper(cell)
	fromCol := string(strings.Map(letterOnlyMapF, cell))
	fromRow, _ := strconv.Atoi(strings.Map(intOnlyMapF, cell))
	row := fromRow - 1
	col := TitleToNumber(fromCol)

	drawingRelationships := strings.Replace(strings.Replace(target, "../drawings", "xl/drawings/_rels", -1), ".xml", ".xml.rels", -1)

	for _, anchor := range decodeWsDr.TwoCellAnchor {
		decodeTwoCellAnchor := decodeTwoCellAnchor{}
		xml.Unmarshal([]byte("<decodeTwoCellAnchor>"+anchor.Content+"</decodeTwoCellAnchor>"), &decodeTwoCellAnchor)
		if decodeTwoCellAnchor.From != nil && decodeTwoCellAnchor.Pic != nil {
			if decodeTwoCellAnchor.From.Col == col && decodeTwoCellAnchor.From.Row == row {
				xlsxWorkbookRelation := f.getDrawingRelationships(drawingRelationships, decodeTwoCellAnchor.Pic.BlipFill.Blip.Embed)
				_, ok := supportImageTypes[filepath.Ext(xlsxWorkbookRelation.Target)]
				if ok {
					return filepath.Base(xlsxWorkbookRelation.Target), []byte(f.XLSX[strings.Replace(xlsxWorkbookRelation.Target, "..", "xl", -1)])
				}
			}
		}
	}
	return "", []byte{}
}

// getDrawingRelationships provides function to get drawing relationships from
// xl/drawings/_rels/drawing%s.xml.rels by given file name and relationship ID.
func (f *File) getDrawingRelationships(rels, rID string) *xlsxWorkbookRelation {
	_, ok := f.XLSX[rels]
	if !ok {
		return nil
	}
	var drawingRels xlsxWorkbookRels
	xml.Unmarshal([]byte(f.readXML(rels)), &drawingRels)
	for _, v := range drawingRels.Relationships {
		if v.ID == rID {
			return &v
		}
	}
	return nil
}
