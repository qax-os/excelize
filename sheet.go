package excelize

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
)

// NewSheet provice function to greate a new sheet by given index, when
// creating a new XLSX file, the default sheet will be create, when you
// create a new file, you need to ensure that the index is continuous.
func (f *File) NewSheet(index int, name string) {
	// Update docProps/app.xml
	f.setAppXML()
	// Update [Content_Types].xml
	f.setContentTypes(index)
	// Create new sheet /xl/worksheets/sheet%d.xml
	f.setSheet(index)
	// Update xl/_rels/workbook.xml.rels
	f.addXlsxWorkbookRels(index)
	// Update xl/workbook.xml
	f.setWorkbook(index, name)
}

// Read and update property of contents type of XLSX.
func (f *File) setContentTypes(index int) {
	var content xlsxTypes
	xml.Unmarshal([]byte(f.readXML(`[Content_Types].xml`)), &content)
	content.Overrides = append(content.Overrides, xlsxOverride{
		PartName:    `/xl/worksheets/sheet` + strconv.Itoa(index) + `.xml`,
		ContentType: `application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`,
	})
	output, err := xml.Marshal(content)
	if err != nil {
		fmt.Println(err)
	}
	f.saveFileList(`[Content_Types].xml`, string(output))
}

// Update sheet property by given index.
func (f *File) setSheet(index int) {
	var xlsx xlsxWorksheet
	xlsx.Dimension.Ref = `A1`
	xlsx.SheetViews.SheetView = append(xlsx.SheetViews.SheetView, xlsxSheetView{
		WorkbookViewID: 0,
	})
	output, err := xml.Marshal(xlsx)
	if err != nil {
		fmt.Println(err)
	}
	path := `xl/worksheets/sheet` + strconv.Itoa(index) + `.xml`
	f.saveFileList(path, replaceRelationshipsID(replaceWorkSheetsRelationshipsNameSpace(string(output))))
}

// Update workbook property of XLSX.
func (f *File) setWorkbook(index int, name string) {
	var content xlsxWorkbook
	xml.Unmarshal([]byte(f.readXML(`xl/workbook.xml`)), &content)

	rels := f.readXlsxWorkbookRels()
	rID := len(rels.Relationships)
	content.Sheets.Sheet = append(content.Sheets.Sheet, xlsxSheet{
		Name:    name,
		SheetID: strconv.Itoa(index),
		ID:      "rId" + strconv.Itoa(rID),
	})
	output, err := xml.Marshal(content)
	if err != nil {
		fmt.Println(err)
	}
	f.saveFileList(`xl/workbook.xml`, replaceRelationshipsNameSpace(string(output)))
}

// Read and unmarshal workbook relationships of XLSX.
func (f *File) readXlsxWorkbookRels() xlsxWorkbookRels {
	var content xlsxWorkbookRels
	xml.Unmarshal([]byte(f.readXML(`xl/_rels/workbook.xml.rels`)), &content)
	return content
}

// Update workbook relationships property of XLSX.
func (f *File) addXlsxWorkbookRels(sheet int) {
	content := f.readXlsxWorkbookRels()
	rID := len(content.Relationships) + 1
	ID := bytes.Buffer{}
	ID.WriteString(`rId`)
	ID.WriteString(strconv.Itoa(rID))
	target := bytes.Buffer{}
	target.WriteString(`worksheets/sheet`)
	target.WriteString(strconv.Itoa(sheet))
	target.WriteString(`.xml`)
	content.Relationships = append(content.Relationships, xlsxWorkbookRelation{
		ID:     ID.String(),
		Target: target.String(),
		Type:   `http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`,
	})
	output, err := xml.Marshal(content)
	if err != nil {
		fmt.Println(err)
	}
	f.saveFileList(`xl/_rels/workbook.xml.rels`, string(output))
}

// Update docProps/app.xml file of XML.
func (f *File) setAppXML() {
	f.saveFileList(`docProps/app.xml`, templateDocpropsApp)
}

// Some tools that read XLSX files have very strict requirements about
// the structure of the input XML.  In particular both Numbers on the Mac
// and SAS dislike inline XML namespace declarations, or namespace
// prefixes that don't match the ones that Excel itself uses.  This is a
// problem because the Go XML library doesn't multiple namespace
// declarations in a single element of a document.  This function is a
// horrible hack to fix that after the XML marshalling is completed.
func replaceRelationshipsNameSpace(workbookMarshal string) string {
	oldXmlns := `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`
	newXmlns := `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
	return strings.Replace(workbookMarshal, oldXmlns, newXmlns, -1)
}

// replace relationships ID in worksheets/sheet%d.xml
func replaceRelationshipsID(workbookMarshal string) string {
	rids := strings.Replace(workbookMarshal, `<drawing rid="" />`, ``, -1)
	rids = strings.Replace(rids, `<hyperlinks></hyperlinks>`, ``, -1)
	return strings.Replace(rids, `<drawing rid="`, `<drawing r:id="`, -1)
}

// SetActiveSheet provide function to set default active sheet of XLSX by given index.
func (f *File) SetActiveSheet(index int) {
	var content xlsxWorkbook
	if index < 1 {
		index = 1
	}
	index--
	xml.Unmarshal([]byte(f.readXML(`xl/workbook.xml`)), &content)
	if len(content.BookViews.WorkBookView) > 0 {
		content.BookViews.WorkBookView[0].ActiveTab = index
	} else {
		content.BookViews.WorkBookView = append(content.BookViews.WorkBookView, xlsxWorkBookView{
			ActiveTab: index,
		})
	}
	sheets := len(content.Sheets.Sheet)
	output, err := xml.Marshal(content)
	if err != nil {
		fmt.Println(err)
	}
	f.saveFileList(`xl/workbook.xml`, workBookCompatibility(replaceRelationshipsNameSpace(string(output))))
	index++
	buffer := bytes.Buffer{}
	for i := 0; i < sheets; i++ {
		xlsx := xlsxWorksheet{}
		sheetIndex := i + 1
		buffer.WriteString(`xl/worksheets/sheet`)
		buffer.WriteString(strconv.Itoa(sheetIndex))
		buffer.WriteString(`.xml`)
		xml.Unmarshal([]byte(f.readXML(buffer.String())), &xlsx)
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
		sheet, err := xml.Marshal(xlsx)
		if err != nil {
			fmt.Println(err)
		}
		f.saveFileList(buffer.String(), replaceRelationshipsID(replaceWorkSheetsRelationshipsNameSpace(string(sheet))))
		buffer.Reset()
	}
	return
}

// Replace xl/workbook.xml XML tags to self-closing for compatible Office Excel 2007.
func workBookCompatibility(workbookMarshal string) string {
	workbookMarshal = strings.Replace(workbookMarshal, `xmlns:relationships="http://schemas.openxmlformats.org/officeDocument/2006/relationships" relationships:id="`, `r:id="`, -1)
	workbookMarshal = strings.Replace(workbookMarshal, `></sheet>`, ` />`, -1)
	workbookMarshal = strings.Replace(workbookMarshal, `></workbookView>`, ` />`, -1)
	workbookMarshal = strings.Replace(workbookMarshal, `></fileVersion>`, ` />`, -1)
	workbookMarshal = strings.Replace(workbookMarshal, `></workbookPr>`, ` />`, -1)
	workbookMarshal = strings.Replace(workbookMarshal, `></definedNames>`, ` />`, -1)
	workbookMarshal = strings.Replace(workbookMarshal, `></calcPr>`, ` />`, -1)
	workbookMarshal = strings.Replace(workbookMarshal, `></workbookProtection>`, ` />`, -1)
	workbookMarshal = strings.Replace(workbookMarshal, `></fileRecoveryPr>`, ` />`, -1)
	workbookMarshal = strings.Replace(workbookMarshal, `></hyperlink>`, ` />`, -1)
	return workbookMarshal
}
