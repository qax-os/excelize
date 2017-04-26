package excelize

import (
	"bytes"
	"encoding/xml"
	"errors"
	"os"
	"path"
	"strconv"
	"strings"
)

// NewSheet provides function to create a new sheet by given index, when
// creating a new XLSX file, the default sheet will be create, when you create a
// new file, you need to ensure that the index is continuous.
func (f *File) NewSheet(index int, name string) {
	// Update docProps/app.xml
	f.setAppXML()
	// Update [Content_Types].xml
	f.setContentTypes(index)
	// Create new sheet /xl/worksheets/sheet%d.xml
	f.setSheet(index)
	// Update xl/_rels/workbook.xml.rels
	rID := f.addXlsxWorkbookRels(index)
	// Update xl/workbook.xml
	f.setWorkbook(name, rID)
	f.SheetCount++
}

// contentTypesReader provides function to get the pointer to the
// [Content_Types].xml structure after deserialization.
func (f *File) contentTypesReader() *xlsxTypes {
	if f.ContentTypes == nil {
		var content xlsxTypes
		xml.Unmarshal([]byte(f.readXML("[Content_Types].xml")), &content)
		f.ContentTypes = &content
	}
	return f.ContentTypes
}

// contentTypesWriter provides function to save [Content_Types].xml after
// serialize structure.
func (f *File) contentTypesWriter() {
	if f.ContentTypes != nil {
		output, _ := xml.Marshal(f.ContentTypes)
		f.saveFileList("[Content_Types].xml", string(output))
	}
}

// workbookReader provides function to get the pointer to the xl/workbook.xml
// structure after deserialization.
func (f *File) workbookReader() *xlsxWorkbook {
	if f.WorkBook == nil {
		var content xlsxWorkbook
		xml.Unmarshal([]byte(f.readXML("xl/workbook.xml")), &content)
		f.WorkBook = &content
	}
	return f.WorkBook
}

// workbookWriter provides function to save xl/workbook.xml after serialize
// structure.
func (f *File) workbookWriter() {
	if f.WorkBook != nil {
		output, _ := xml.Marshal(f.WorkBook)
		f.saveFileList("xl/workbook.xml", replaceRelationshipsNameSpace(string(output)))
	}
}

// worksheetWriter provides function to save xl/worksheets/sheet%d.xml after
// serialize structure.
func (f *File) worksheetWriter() {
	for path, sheet := range f.Sheet {
		if sheet != nil {
			output, _ := xml.Marshal(sheet)
			f.saveFileList(path, replaceWorkSheetsRelationshipsNameSpace(string(output)))
		}
	}
}

// Read and update property of contents type of XLSX.
func (f *File) setContentTypes(index int) {
	content := f.contentTypesReader()
	content.Overrides = append(content.Overrides, xlsxOverride{
		PartName:    "/xl/worksheets/sheet" + strconv.Itoa(index) + ".xml",
		ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
	})
}

// Update sheet property by given index.
func (f *File) setSheet(index int) {
	var xlsx xlsxWorksheet
	xlsx.Dimension.Ref = "A1"
	xlsx.SheetViews.SheetView = append(xlsx.SheetViews.SheetView, xlsxSheetView{
		WorkbookViewID: 0,
	})
	path := "xl/worksheets/sheet" + strconv.Itoa(index) + ".xml"
	f.Sheet[path] = &xlsx
}

// setWorkbook update workbook property of XLSX. Maximum 31 characters are
// allowed in sheet title.
func (f *File) setWorkbook(name string, rid int) {
	r := strings.NewReplacer(":", "", "\\", "", "/", "", "?", "", "*", "", "[", "", "]", "")
	name = r.Replace(name)
	if len(name) > 31 {
		name = name[0:31]
	}
	content := f.workbookReader()
	content.Sheets.Sheet = append(content.Sheets.Sheet, xlsxSheet{
		Name:    name,
		SheetID: strconv.Itoa(rid),
		ID:      "rId" + strconv.Itoa(rid),
	})
}

// workbookRelsReader provides function to read and unmarshal workbook
// relationships of XLSX file.
func (f *File) workbookRelsReader() *xlsxWorkbookRels {
	if f.WorkBookRels == nil {
		var content xlsxWorkbookRels
		xml.Unmarshal([]byte(f.readXML("xl/_rels/workbook.xml.rels")), &content)
		f.WorkBookRels = &content
	}
	return f.WorkBookRels
}

// workbookRelsWriter provides function to save xl/_rels/workbook.xml.rels after
// serialize structure.
func (f *File) workbookRelsWriter() {
	if f.WorkBookRels != nil {
		output, _ := xml.Marshal(f.WorkBookRels)
		f.saveFileList("xl/_rels/workbook.xml.rels", string(output))
	}
}

// addXlsxWorkbookRels update workbook relationships property of XLSX.
func (f *File) addXlsxWorkbookRels(sheet int) int {
	content := f.workbookRelsReader()
	rID := 0
	for _, v := range content.Relationships {
		t, _ := strconv.Atoi(strings.TrimPrefix(v.ID, "rId"))
		if t > rID {
			rID = t
		}
	}
	rID++
	ID := bytes.Buffer{}
	ID.WriteString("rId")
	ID.WriteString(strconv.Itoa(rID))
	target := bytes.Buffer{}
	target.WriteString("worksheets/sheet")
	target.WriteString(strconv.Itoa(sheet))
	target.WriteString(".xml")
	content.Relationships = append(content.Relationships, xlsxWorkbookRelation{
		ID:     ID.String(),
		Target: target.String(),
		Type:   SourceRelationshipWorkSheet,
	})
	return rID
}

// setAppXML update docProps/app.xml file of XML.
func (f *File) setAppXML() {
	f.saveFileList("docProps/app.xml", templateDocpropsApp)
}

// Some tools that read XLSX files have very strict requirements about the
// structure of the input XML. In particular both Numbers on the Mac and SAS
// dislike inline XML namespace declarations, or namespace prefixes that don't
// match the ones that Excel itself uses. This is a problem because the Go XML
// library doesn't multiple namespace declarations in a single element of a
// document. This function is a horrible hack to fix that after the XML
// marshalling is completed.
func replaceRelationshipsNameSpace(workbookMarshal string) string {
	oldXmlns := `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`
	newXmlns := `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">`
	return strings.Replace(workbookMarshal, oldXmlns, newXmlns, -1)
}

// SetActiveSheet provides function to set default active sheet of XLSX by given
// index.
func (f *File) SetActiveSheet(index int) {
	if index < 1 {
		index = 1
	}
	index--
	content := f.workbookReader()
	if len(content.BookViews.WorkBookView) > 0 {
		content.BookViews.WorkBookView[0].ActiveTab = index
	} else {
		content.BookViews.WorkBookView = append(content.BookViews.WorkBookView, xlsxWorkBookView{
			ActiveTab: index,
		})
	}
	sheets := len(content.Sheets.Sheet)
	index++
	for i := 0; i < sheets; i++ {
		sheetIndex := i + 1
		xlsx := f.workSheetReader("sheet" + strconv.Itoa(sheetIndex))
		if index == sheetIndex {
			if len(xlsx.SheetViews.SheetView) > 0 {
				xlsx.SheetViews.SheetView[0].TabSelected = true
			} else {
				xlsx.SheetViews.SheetView = append(xlsx.SheetViews.SheetView, xlsxSheetView{
					TabSelected: true,
				})
			}
		} else {
			if len(xlsx.SheetViews.SheetView) > 0 {
				xlsx.SheetViews.SheetView[0].TabSelected = false
			}
		}
	}
	return
}

// GetActiveSheetIndex provides function to get active sheet of XLSX. If not
// found the active sheet will be return integer 0.
func (f *File) GetActiveSheetIndex() int {
	buffer := bytes.Buffer{}
	content := f.workbookReader()
	for _, v := range content.Sheets.Sheet {
		xlsx := xlsxWorksheet{}
		buffer.WriteString("xl/worksheets/sheet")
		buffer.WriteString(strings.TrimPrefix(v.ID, "rId"))
		buffer.WriteString(".xml")
		xml.Unmarshal([]byte(f.readXML(buffer.String())), &xlsx)
		for _, sheetView := range xlsx.SheetViews.SheetView {
			if sheetView.TabSelected {
				ID, _ := strconv.Atoi(strings.TrimPrefix(v.ID, "rId"))
				return ID
			}
		}
		buffer.Reset()
	}
	return 0
}

// SetSheetName provides function to set the sheet name be given old and new
// sheet name. Maximum 31 characters are allowed in sheet title and this
// function only changes the name of the sheet and will not update the sheet
// name in the formula or reference associated with the cell. So there may be
// problem formula error or reference missing.
func (f *File) SetSheetName(oldName, newName string) {
	oldName = trimSheetName(oldName)
	newName = trimSheetName(newName)
	content := f.workbookReader()
	for k, v := range content.Sheets.Sheet {
		if v.Name == oldName {
			content.Sheets.Sheet[k].Name = newName
		}
	}
}

// GetSheetName provides function to get sheet name of XLSX by given worksheet
// index. If given sheet index is invalid, will return an empty string.
func (f *File) GetSheetName(index int) string {
	content := f.workbookReader()
	rels := f.workbookRelsReader()
	for _, rel := range rels.Relationships {
		rID, _ := strconv.Atoi(strings.TrimSuffix(strings.TrimPrefix(rel.Target, "worksheets/sheet"), ".xml"))
		if rID == index {
			for _, v := range content.Sheets.Sheet {
				if v.ID == rel.ID {
					return v.Name
				}
			}
		}
	}
	return ""
}

// GetSheetIndex provides function to get worksheet index of XLSX by given sheet
// name. If given sheet name is invalid, will return an integer type value 0.
func (f *File) GetSheetIndex(name string) int {
	content := f.workbookReader()
	rels := f.workbookRelsReader()
	for _, v := range content.Sheets.Sheet {
		if v.Name == name {
			for _, rel := range rels.Relationships {
				if v.ID == rel.ID {
					rID, _ := strconv.Atoi(strings.TrimSuffix(strings.TrimPrefix(rel.Target, "worksheets/sheet"), ".xml"))
					return rID
				}
			}
		}
	}
	return 0
}

// GetSheetMap provides function to get sheet map of XLSX. For example:
//
//    xlsx, err := excelize.OpenFile("/tmp/Workbook.xlsx")
//    if err != nil {
//        fmt.Println(err)
//        os.Exit(1)
//    }
//    for k, v := range xlsx.GetSheetMap()
//        fmt.Println(k, v)
//    }
//
func (f *File) GetSheetMap() map[int]string {
	content := f.workbookReader()
	rels := f.workbookRelsReader()
	sheetMap := map[int]string{}
	for _, v := range content.Sheets.Sheet {
		for _, rel := range rels.Relationships {
			if rel.ID == v.ID {
				rID, _ := strconv.Atoi(strings.TrimSuffix(strings.TrimPrefix(rel.Target, "worksheets/sheet"), ".xml"))
				sheetMap[rID] = v.Name
			}
		}
	}
	return sheetMap
}

// SetSheetBackground provides function to set background picture by given sheet
// index.
func (f *File) SetSheetBackground(sheet, picture string) error {
	var err error
	// Check picture exists first.
	if _, err = os.Stat(picture); os.IsNotExist(err) {
		return err
	}
	ext, ok := supportImageTypes[path.Ext(picture)]
	if !ok {
		return errors.New("Unsupported image extension")
	}
	pictureID := f.countMedia() + 1
	rID := f.addSheetRelationships(sheet, SourceRelationshipImage, "../media/image"+strconv.Itoa(pictureID)+ext, "")
	f.addSheetPicture(sheet, rID)
	f.addMedia(picture, ext)
	f.setContentTypePartImageExtensions()
	return err
}

// DeleteSheet provides function to delete worksheet in a workbook by given
// sheet name. Use this method with caution, which will affect changes in
// references such as formulas, charts, and so on. If there is any referenced
// value of the deleted worksheet, it will cause a file error when you open it.
// This function will be invalid when only the one worksheet is left.
func (f *File) DeleteSheet(name string) {
	content := f.workbookReader()
	for k, v := range content.Sheets.Sheet {
		if v.Name != name || len(content.Sheets.Sheet) < 2 {
			continue
		}
		content.Sheets.Sheet = append(content.Sheets.Sheet[:k], content.Sheets.Sheet[k+1:]...)
		sheet := "xl/worksheets/sheet" + strings.TrimPrefix(v.ID, "rId") + ".xml"
		rels := "xl/worksheets/_rels/sheet" + strings.TrimPrefix(v.ID, "rId") + ".xml.rels"
		target := f.deleteSheetFromWorkbookRels(v.ID)
		f.deleteSheetFromContentTypes(target)
		_, ok := f.XLSX[sheet]
		if ok {
			delete(f.XLSX, sheet)
		}
		_, ok = f.XLSX[rels]
		if ok {
			delete(f.XLSX, rels)
		}
		_, ok = f.Sheet[sheet]
		if ok {
			delete(f.Sheet, sheet)
		}
		f.SheetCount--
	}
}

// deleteSheetFromWorkbookRels provides function to remove worksheet
// relationships by given relationships ID in the file
// xl/_rels/workbook.xml.rels.
func (f *File) deleteSheetFromWorkbookRels(rID string) string {
	content := f.workbookRelsReader()
	for k, v := range content.Relationships {
		if v.ID != rID {
			continue
		}
		content.Relationships = append(content.Relationships[:k], content.Relationships[k+1:]...)
		return v.Target
	}
	return ""
}

// deleteSheetFromContentTypes provides function to remove worksheet
// relationships by given target name in the file [Content_Types].xml.
func (f *File) deleteSheetFromContentTypes(target string) {
	content := f.contentTypesReader()
	for k, v := range content.Overrides {
		if v.PartName != "/xl/"+target {
			continue
		}
		content.Overrides = append(content.Overrides[:k], content.Overrides[k+1:]...)
	}
}

// CopySheet provides function to duplicate a worksheet by gave source and
// target worksheet index. Note that currently doesn't support duplicate
// workbooks that contain tables, charts or pictures. For Example:
//
//    // Sheet1 already exists...
//    xlsx.NewSheet(2, "sheet2")
//    err := xlsx.CopySheet(1, 2)
//    if err != nil {
//        fmt.Println(err)
//        os.Exit(1)
//    }
//
func (f *File) CopySheet(from, to int) error {
	if from < 1 || to < 1 || from == to || f.GetSheetName(from) == "" || f.GetSheetName(to) == "" {
		return errors.New("Invalid worksheet index")
	}
	f.copySheet(from, to)
	return nil
}

// copySheet provides function to duplicate a worksheet by gave source and
// target worksheet index.
func (f *File) copySheet(from, to int) {
	sheet := f.workSheetReader("sheet" + strconv.Itoa(from))
	var worksheet xlsxWorksheet
	worksheet = *sheet
	path := "xl/worksheets/sheet" + strconv.Itoa(to) + ".xml"
	if len(worksheet.SheetViews.SheetView) > 0 {
		worksheet.SheetViews.SheetView[0].TabSelected = false
	}
	worksheet.Drawing = nil
	worksheet.TableParts = nil
	worksheet.PageSetUp = nil
	f.Sheet[path] = &worksheet
	toRels := "xl/worksheets/_rels/sheet" + strconv.Itoa(to) + ".xml.rels"
	fromRels := "xl/worksheets/_rels/sheet" + strconv.Itoa(from) + ".xml.rels"
	_, ok := f.XLSX[fromRels]
	if ok {
		f.XLSX[toRels] = f.XLSX[fromRels]
	}
}

// HideSheet provides function to hide worksheet by given name. A workbook must
// contain at least one visible worksheet. If the given worksheet has been
// activated, this setting will be invalidated.
func (f *File) HideSheet(name string) {
	name = trimSheetName(name)
	content := f.workbookReader()
	count := 0
	for _, v := range content.Sheets.Sheet {
		if v.State != "hidden" {
			count++
		}
	}
	for k, v := range content.Sheets.Sheet {
		sheetIndex := k + 1
		xlsx := f.workSheetReader("sheet" + strconv.Itoa(sheetIndex))
		tabSelected := false
		if len(xlsx.SheetViews.SheetView) > 0 {
			tabSelected = xlsx.SheetViews.SheetView[0].TabSelected
		}
		if v.Name == name && count > 1 && !tabSelected {
			content.Sheets.Sheet[k].State = "hidden"
		}
	}
}

// UnhideSheet provides function to unhide worksheet by given name.
func (f *File) UnhideSheet(name string) {
	name = trimSheetName(name)
	content := f.workbookReader()
	for k, v := range content.Sheets.Sheet {
		if v.Name == name {
			content.Sheets.Sheet[k].State = ""
		}
	}
}

// trimSheetName provides function to trim invaild characters by given worksheet
// name.
func trimSheetName(name string) string {
	r := strings.NewReplacer(":", "", "\\", "", "/", "", "?", "", "*", "", "[", "", "]", "")
	name = r.Replace(name)
	if len(name) > 31 {
		name = name[0:31]
	}
	return name
}
