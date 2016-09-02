package excelize

import (
	"archive/zip"
	"bytes"
	"os"
)

// CreateFile provide function to create new file by default template
// For example:
// xlsx := CreateFile()
func CreateFile() []FileList {
	var file []FileList
	file = saveFileList(file, `_rels/.rels`, templateRels)
	file = saveFileList(file, `docProps/app.xml`, templateDocpropsApp)
	file = saveFileList(file, `docProps/core.xml`, templateDocpropsCore)
	file = saveFileList(file, `xl/_rels/workbook.xml.rels`, templateWorkbookRels)
	file = saveFileList(file, `xl/theme/theme1.xml`, templateTheme)
	file = saveFileList(file, `xl/worksheets/sheet1.xml`, templateSheet)
	file = saveFileList(file, `xl/styles.xml`, templateStyles)
	file = saveFileList(file, `xl/workbook.xml`, templateWorkbook)
	file = saveFileList(file, `[Content_Types].xml`, templateContentTypes)
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
