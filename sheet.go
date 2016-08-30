package excelize

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
)

// Create a new sheet by given index, when creating a new XLSX file,
// the default sheet will be create, when you create a new file, you
//  need to ensure that the index is continuous.
func NewSheet(file []FileList, index int, name string) []FileList {
	// Update docProps/app.xml
	file = setAppXml(file)
	// Update [Content_Types].xml
	file = setContentTypes(file, index)
	// Create new sheet /xl/worksheets/sheet%d.xml
	file = setSheet(file, index)
	// Update xl/_rels/workbook.xml.rels
	file = addXlsxWorkbookRels(file, index)
	// Update xl/workbook.xml
	file = setWorkbook(file, index, name)
	return file
}

// Read and update property of contents type of XLSX
func setContentTypes(file []FileList, index int) []FileList {
	var content xlsxTypes
	xml.Unmarshal([]byte(readXml(file, `[Content_Types].xml`)), &content)
	content.Overrides = append(content.Overrides, xlsxOverride{
		PartName:    fmt.Sprintf("/xl/worksheets/sheet%d.xml", index),
		ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
	})
	output, err := xml.MarshalIndent(content, "", "")
	if err != nil {
		fmt.Println(err)
	}
	return saveFileList(file, `[Content_Types].xml`, string(output))
}

// Update sheet property by given index
func setSheet(file []FileList, index int) []FileList {
	var xlsx xlsxWorksheet
	xlsx.Dimension.Ref = "A1"
	xlsx.SheetViews.SheetView = append(xlsx.SheetViews.SheetView, xlsxSheetView{
		WorkbookViewId: 0,
	})
	output, err := xml.MarshalIndent(xlsx, "", "")
	if err != nil {
		fmt.Println(err)
	}
	path := fmt.Sprintf("xl/worksheets/sheet%d.xml", index)
	return saveFileList(file, path, replaceRelationshipsID(replaceWorkSheetsRelationshipsNameSpace(string(output))))
}

// Update workbook property of XLSX
func setWorkbook(file []FileList, index int, name string) []FileList {
	var content xlsxWorkbook
	xml.Unmarshal([]byte(readXml(file, `xl/workbook.xml`)), &content)

	rels := readXlsxWorkbookRels(file)
	rId := len(rels.Relationships)
	content.Sheets.Sheet = append(content.Sheets.Sheet, xlsxSheet{
		Name:    name,
		SheetId: strconv.Itoa(index),
		Id:      "rId" + strconv.Itoa(rId),
	})
	output, err := xml.MarshalIndent(content, "", "")
	if err != nil {
		fmt.Println(err)
	}
	return saveFileList(file, `xl/workbook.xml`, replaceRelationshipsNameSpace(string(output)))
}

// Read and unmarshal workbook relationships of XLSX
func readXlsxWorkbookRels(file []FileList) xlsxWorkbookRels {
	var content xlsxWorkbookRels
	xml.Unmarshal([]byte(readXml(file, `xl/_rels/workbook.xml.rels`)), &content)
	return content
}

// Update workbook relationships property of XLSX
func addXlsxWorkbookRels(file []FileList, sheet int) []FileList {
	content := readXlsxWorkbookRels(file)
	rId := len(content.Relationships) + 1
	content.Relationships = append(content.Relationships, xlsxWorkbookRelation{
		Id:     "rId" + strconv.Itoa(rId),
		Target: fmt.Sprintf("worksheets/sheet%d.xml", sheet),
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
	})
	output, err := xml.MarshalIndent(content, "", "")
	if err != nil {
		fmt.Println(err)
	}
	return saveFileList(file, `xl/_rels/workbook.xml.rels`, string(output))
}

// Update docProps/app.xml file of XML
func setAppXml(file []FileList) []FileList {
	return saveFileList(file, `docProps/app.xml`, TEMPLATE_DOCPROPS_APP)
}

// Some tools that read XLSX files have very strict requirements about
// the structure of the input XML.  In particular both Numbers on the Mac
// and SAS dislike inline XML namespace declarations, or namespace
// prefixes that don't match the ones that Excel itself uses.  This is a
// problem because the Go XML library doesn't multiple namespace
// declarations in a single element of a document.  This function is a
// horrible hack to fix that after the XML marshalling is completed.
func replaceRelationshipsNameSpace(workbookMarshal string) string {
	// newWorkbook := strings.Replace(workbookMarshal, `xmlns:relationships="http://schemas.openxmlformats.org/officeDocument/2006/relationships" relationships:id`, `r:id`, -1)
	// Dirty hack to fix issues #63 and #91; encoding/xml currently
	// "doesn't allow for additional namespaces to be defined in the
	// root element of the document," as described by @tealeg in the
	// comments for #63.
	oldXmlns := `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`
	newXmlns := `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
	return strings.Replace(workbookMarshal, oldXmlns, newXmlns, -1)
}

// replace relationships ID in worksheets/sheet%d.xml
func replaceRelationshipsID(workbookMarshal string) string {
	rids := strings.Replace(workbookMarshal, `<drawing rid="" />`, ``, -1)
	return strings.Replace(rids, `<drawing rid="`, `<drawing r:id="`, -1)
}

// Set default active sheet of XLSX by given index
func SetActiveSheet(file []FileList, index int) []FileList {
	var content xlsxWorkbook
	if index < 1 {
		index = 1
	}
	index -= 1
	xml.Unmarshal([]byte(readXml(file, `xl/workbook.xml`)), &content)
	if len(content.BookViews.WorkBookView) > 0 {
		content.BookViews.WorkBookView[0].ActiveTab = index
	} else {
		content.BookViews.WorkBookView = append(content.BookViews.WorkBookView, xlsxWorkBookView{
			ActiveTab: index,
		})
	}
	sheets := len(content.Sheets.Sheet)
	output, err := xml.MarshalIndent(content, "", "")
	if err != nil {
		fmt.Println(err)
	}
	file = saveFileList(file, `xl/workbook.xml`, workBookCompatibility(replaceRelationshipsNameSpace(string(output))))
	index += 1
	for i := 0; i < sheets; i++ {
		xlsx := xlsxWorksheet{}
		sheetIndex := i + 1
		path := fmt.Sprintf("xl/worksheets/sheet%d.xml", sheetIndex)
		xml.Unmarshal([]byte(readXml(file, path)), &xlsx)
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
		sheet, err := xml.MarshalIndent(xlsx, "", "")
		if err != nil {
			fmt.Println(err)
		}
		file = saveFileList(file, path, replaceRelationshipsID(replaceWorkSheetsRelationshipsNameSpace(string(sheet))))
	}
	return file
}

// Replace xl/workbook.xml XML tags to self-closing for compatible Office Excel 2007
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
	return workbookMarshal
}
