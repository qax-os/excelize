package excelize

import (
	"archive/zip"
	"bytes"
	"os"
)

// Create a new xlsx file
//
// For example:
//
// xlsx := CreateFile()
//
func CreateFile() []FileList {
	var file []FileList
	file = saveFileList(file, `_rels/.rels`, TEMPLATE_RELS)
	file = saveFileList(file, `docProps/app.xml`, TEMPLATE_DOCPROPS_APP)
	file = saveFileList(file, `docProps/core.xml`, TEMPLATE_DOCPROPS_CORE)
	file = saveFileList(file, `xl/_rels/workbook.xml.rels`, TEMPLATE_WORKBOOK_RELS)
	file = saveFileList(file, `xl/theme/theme1.xml`, TEMPLATE_THEME)
	file = saveFileList(file, `xl/worksheets/sheet1.xml`, TEMPLATE_SHEET)
	file = saveFileList(file, `xl/styles.xml`, TEMPLATE_STYLES)
	file = saveFileList(file, `xl/workbook.xml`, TEMPLATE_WORKBOOK)
	file = saveFileList(file, `[Content_Types].xml`, TEMPLATE_CONTENT_TYPES)
	return file
}

// Save after create or update to an xlsx file at the provided path.
func Save(files []FileList, name string) error {
	buf := new(bytes.Buffer)
	w := zip.NewWriter(buf)
	for _, file := range files {
		f, err := w.Create(file.Key)
		if err != nil {
			return err
		}
		_, err = f.Write([]byte(file.Value))
		if err != nil {
			return err
		}
	}
	err := w.Close()
	if err != nil {
		return err
	}
	f, err := os.OpenFile(name, os.O_WRONLY|os.O_TRUNC|os.O_CREATE, 0666)
	if err != nil {
		return err
	}
	buf.WriteTo(f)
	return err
}
