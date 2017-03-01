package excelize

import (
	"bytes"
	"encoding/json"
	"encoding/xml"
	"errors"
	"fmt"
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
func parseFormatPictureSet(formatSet string) *xlsxFormatPicture {
	format := xlsxFormatPicture{
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
//        "os"
//        _ "image/gif"
//        _ "image/jpeg"
//        _ "image/png"
//
//        "github.com/Luxurioust/excelize"
//    )
//
//    func main() {
//        xlsx := excelize.CreateFile()
//        // Insert a picture.
//        err := xlsx.AddPicture("Sheet1", "A2", "/tmp/image1.jpg", "")
//        if err != nil {
//            fmt.Println(err)
//        }
//        // Insert a picture to sheet with scaling.
//        err = xlsx.AddPicture("Sheet1", "D2", "/tmp/image1.png", `{"x_scale": 0.5, "y_scale": 0.5}`)
//        if err != nil {
//            fmt.Println(err)
//        }
//        // Insert a picture offset in the cell with printing support.
//        err = xlsx.AddPicture("Sheet1", "H2", "/tmp/image3.gif", `{"x_offset": 15, "y_offset": 10, "print_obj": true, "lock_aspect_ratio": false, "locked": false}`)
//        if err != nil {
//            fmt.Println(err)
//        }
//        err = xlsx.WriteTo("/tmp/Workbook.xlsx")
//        if err != nil {
//            fmt.Println(err)
//            os.Exit(1)
//        }
//    }
//
func (f *File) AddPicture(sheet, cell, picture, format string) error {
	var supportTypes = map[string]string{".gif": ".gif", ".jpg": ".jpeg", ".jpeg": ".jpeg", ".png": ".png"}
	var err error
	// Check picture exists first.
	if _, err = os.Stat(picture); os.IsNotExist(err) {
		return err
	}
	ext, ok := supportTypes[path.Ext(picture)]
	if !ok {
		return errors.New("Unsupported image extension")
	}
	readFile, _ := os.Open(picture)
	image, _, err := image.DecodeConfig(readFile)
	_, file := filepath.Split(picture)
	formatSet := parseFormatPictureSet(format)
	// Read sheet data.
	var xlsx xlsxWorksheet
	name := "xl/worksheets/" + strings.ToLower(sheet) + ".xml"
	xml.Unmarshal([]byte(f.readXML(name)), &xlsx)
	// Add first picture for given sheet, create xl/drawings/ and xl/drawings/_rels/ folder.
	drawingID := f.countDrawings() + 1
	pictureID := f.countMedia() + 1
	drawingXML := "xl/drawings/drawing" + strconv.Itoa(drawingID) + ".xml"
	sheetRelationshipsDrawingXML := "../drawings/drawing" + strconv.Itoa(drawingID) + ".xml"

	var drawingRID int
	if xlsx.Drawing != nil {
		// The worksheet already has a picture or chart relationships, use the relationships drawing ../drawings/drawing%d.xml.
		sheetRelationshipsDrawingXML = f.getSheetRelationshipsTargetByID(sheet, xlsx.Drawing.RID)
		drawingID, _ = strconv.Atoi(strings.TrimSuffix(strings.TrimPrefix(sheetRelationshipsDrawingXML, "../drawings/drawing"), ".xml"))
		drawingXML = strings.Replace(sheetRelationshipsDrawingXML, "..", "xl", -1)
	} else {
		// Add first picture for given sheet.
		rID := f.addSheetRelationships(sheet, SourceRelationshipDrawingML, sheetRelationshipsDrawingXML, "")
		f.addSheetDrawing(sheet, rID)
	}
	drawingRID = f.addDrawingRelationships(drawingID, SourceRelationshipImage, "../media/image"+strconv.Itoa(pictureID)+ext)
	f.addDrawing(sheet, drawingXML, cell, file, image.Width, image.Height, drawingRID, formatSet)
	f.addMedia(picture, ext)
	f.addDrawingContentTypePart(drawingID)
	return err
}

// addSheetRelationships provides function to add
// xl/worksheets/_rels/sheet%d.xml.rels by given sheet name, relationship type
// and target.
func (f *File) addSheetRelationships(sheet, relType, target, targetMode string) int {
	var rels = "xl/worksheets/_rels/" + strings.ToLower(sheet) + ".xml.rels"
	var sheetRels xlsxWorkbookRels
	var rID = 1
	var ID bytes.Buffer
	ID.WriteString("rId")
	ID.WriteString(strconv.Itoa(rID))
	_, ok := f.XLSX[rels]
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
	output, err := xml.Marshal(sheetRels)
	if err != nil {
		fmt.Println(err)
	}
	f.saveFileList(rels, string(output))
	return rID
}

// addSheetDrawing provides function to add drawing element to
// xl/worksheets/sheet%d.xml by given sheet name and relationship index.
func (f *File) addSheetDrawing(sheet string, rID int) {
	var xlsx xlsxWorksheet
	name := "xl/worksheets/" + strings.ToLower(sheet) + ".xml"
	xml.Unmarshal([]byte(f.readXML(name)), &xlsx)
	xlsx.Drawing = &xlsxDrawing{
		RID: "rId" + strconv.Itoa(rID),
	}
	output, err := xml.Marshal(xlsx)
	if err != nil {
		fmt.Println(err)
	}
	f.saveFileList(name, replaceWorkSheetsRelationshipsNameSpace(string(output)))
}

// addSheetPicture provides function to add picture element to
// xl/worksheets/sheet%d.xml by given sheet name and relationship index.
func (f *File) addSheetPicture(sheet string, rID int) {
	var xlsx xlsxWorksheet
	name := "xl/worksheets/" + strings.ToLower(sheet) + ".xml"
	xml.Unmarshal([]byte(f.readXML(name)), &xlsx)
	xlsx.Picture = &xlsxPicture{
		RID: "rId" + strconv.Itoa(rID),
	}
	output, err := xml.Marshal(xlsx)
	if err != nil {
		fmt.Println(err)
	}
	f.saveFileList(name, replaceWorkSheetsRelationshipsNameSpace(string(output)))
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

// addDrawing provides function to add picture by given drawingXML, xAxis,
// yAxis, file name and relationship index. In order to solve the problem that
// the label structure is changed after serialization and deserialization, two
// different structures: decodeWsDr and encodeWsDr are defined.
func (f *File) addDrawing(sheet, drawingXML, cell, file string, width, height, rID int, formatSet *xlsxFormatPicture) {
	cell = strings.ToUpper(cell)
	fromCol := string(strings.Map(letterOnlyMapF, cell))
	fromRow, _ := strconv.Atoi(strings.Map(intOnlyMapF, cell))
	row := fromRow - 1
	col := titleToNumber(fromCol)
	width = int(float64(width) * formatSet.XScale)
	height = int(float64(height) * formatSet.YScale)
	colStart, rowStart, _, _, colEnd, rowEnd, x2, y2 := f.positionObjectPixels(sheet, col, row, formatSet.OffsetX, formatSet.OffsetY, width, height)
	content := encodeWsDr{}
	content.WsDr.A = NameSpaceDrawingML
	content.WsDr.Xdr = NameSpaceSpreadSheetDrawing
	cNvPrID := 1
	_, ok := f.XLSX[drawingXML]
	if ok { // Append Model
		decodeWsDr := decodeWsDr{}
		xml.Unmarshal([]byte(f.readXML(drawingXML)), &decodeWsDr)
		cNvPrID = len(decodeWsDr.TwoCellAnchor) + 1
		for _, v := range decodeWsDr.OneCellAnchor {
			content.WsDr.OneCellAnchor = append(content.WsDr.OneCellAnchor, &xlsxCellAnchor{
				EditAs:       v.EditAs,
				GraphicFrame: v.Content,
			})
		}
		for _, v := range decodeWsDr.TwoCellAnchor {
			content.WsDr.TwoCellAnchor = append(content.WsDr.TwoCellAnchor, &xlsxCellAnchor{
				EditAs:       v.EditAs,
				GraphicFrame: v.Content,
			})
		}
	}
	twoCellAnchor := xlsxCellAnchor{}
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
	pic.NvPicPr.CNvPr.ID = cNvPrID
	pic.NvPicPr.CNvPr.Descr = file
	pic.NvPicPr.CNvPr.Name = "Picture " + strconv.Itoa(cNvPrID)
	pic.BlipFill.Blip.R = SourceRelationship
	pic.BlipFill.Blip.Embed = "rId" + strconv.Itoa(rID)
	pic.SpPr.PrstGeom.Prst = "rect"

	twoCellAnchor.Pic = &pic
	twoCellAnchor.ClientData = &xlsxClientData{
		FLocksWithSheet:  formatSet.FLocksWithSheet,
		FPrintsWithSheet: formatSet.FPrintsWithSheet,
	}
	content.WsDr.TwoCellAnchor = append(content.WsDr.TwoCellAnchor, &twoCellAnchor)
	output, err := xml.Marshal(content)
	if err != nil {
		fmt.Println(err)
	}
	// Create replacer with pairs as arguments and replace all pairs.
	r := strings.NewReplacer("<encodeWsDr>", "", "</encodeWsDr>", "")
	result := r.Replace(string(output))
	f.saveFileList(drawingXML, result)
}

// addDrawingRelationships provides function to add image part relationships in
// the file xl/drawings/_rels/drawing%d.xml.rels by given drawing index,
// relationship type and target.
func (f *File) addDrawingRelationships(index int, relType string, target string) int {
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
	output, err := xml.Marshal(drawingRels)
	if err != nil {
		fmt.Println(err)
	}
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
func (f *File) addMedia(file string, ext string) {
	count := f.countMedia()
	dat, _ := ioutil.ReadFile(file)
	media := "xl/media/image" + strconv.Itoa(count+1) + ext
	f.XLSX[media] = string(dat)
}

// setContentTypePartImageExtensions provides function to set the content type
// for relationship parts and the Main Document part.
func (f *File) setContentTypePartImageExtensions() {
	var imageTypes = map[string]bool{"jpeg": false, "png": false, "gif": false}
	var content xlsxTypes
	xml.Unmarshal([]byte(f.readXML("[Content_Types].xml")), &content)
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
	output, _ := xml.Marshal(content)
	f.saveFileList("[Content_Types].xml", string(output))
}

// addDrawingContentTypePart provides function to add image part relationships
// in http://purl.oclc.org/ooxml/officeDocument/relationships/image and
// appropriate content type.
func (f *File) addDrawingContentTypePart(index int) {
	f.setContentTypePartImageExtensions()
	var content xlsxTypes
	xml.Unmarshal([]byte(f.readXML("[Content_Types].xml")), &content)
	for _, v := range content.Overrides {
		if v.PartName == "/xl/drawings/drawing"+strconv.Itoa(index)+".xml" {
			output, _ := xml.Marshal(content)
			f.saveFileList(`[Content_Types].xml`, string(output))
			return
		}
	}
	content.Overrides = append(content.Overrides, xlsxOverride{
		PartName:    "/xl/drawings/drawing" + strconv.Itoa(index) + ".xml",
		ContentType: "application/vnd.openxmlformats-officedocument.drawing+xml",
	})
	output, _ := xml.Marshal(content)
	f.saveFileList("[Content_Types].xml", string(output))
}

// getSheetRelationshipsTargetByID provides function to get Target attribute
// value in xl/worksheets/_rels/sheet%d.xml.rels by given sheet name and
// relationship index.
func (f *File) getSheetRelationshipsTargetByID(sheet string, rID string) string {
	var rels = "xl/worksheets/_rels/" + strings.ToLower(sheet) + ".xml.rels"
	var sheetRels xlsxWorkbookRels
	xml.Unmarshal([]byte(f.readXML(rels)), &sheetRels)
	for _, v := range sheetRels.Relationships {
		if v.ID == rID {
			return v.Target
		}
	}
	return ""
}
