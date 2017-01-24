package excelize

import (
	"archive/zip"
	"bytes"
	"os"
)

// CreateFile provides function to create new file by default template. For
// example:
//
//    xlsx := CreateFile()
//
func CreateFile() *File {
	file := make(map[string]string)
	file["_rels/.rels"] = templateRels
	file["docProps/app.xml"] = templateDocpropsApp
	file["docProps/core.xml"] = templateDocpropsCore
	file["xl/_rels/workbook.xml.rels"] = templateWorkbookRels
	file["xl/theme/theme1.xml"] = templateTheme
	file["xl/worksheets/sheet1.xml"] = templateSheet
	file["xl/styles.xml"] = templateStyles
	file["xl/workbook.xml"] = templateWorkbook
	file["[Content_Types].xml"] = templateContentTypes
	return &File{
		XLSX: file,
	}
}

// Save provides function to override the xlsx file with origin path.
func (f *File) Save() error {
	buf := new(bytes.Buffer)
	w := zip.NewWriter(buf)
	for path, content := range f.XLSX {
		f, err := w.Create(path)
		if err != nil {
			return err
		}
		_, err = f.Write([]byte(content))
		if err != nil {
			return err
		}
	}
	err := w.Close()
	if err != nil {
		return err
	}
	file, err := os.OpenFile(f.Path, os.O_WRONLY|os.O_TRUNC|os.O_CREATE, 0666)
	if err != nil {
		return err
	}
	buf.WriteTo(file)
	return err
}

// WriteTo provides function to create or update to an xlsx file at the provided
// path.
func (f *File) WriteTo(name string) error {
	buf := new(bytes.Buffer)
	w := zip.NewWriter(buf)
	for path, content := range f.XLSX {
		f, err := w.Create(path)
		if err != nil {
			return err
		}
		_, err = f.Write([]byte(content))
		if err != nil {
			return err
		}
	}
	err := w.Close()
	if err != nil {
		return err
	}
	file, err := os.OpenFile(name, os.O_WRONLY|os.O_TRUNC|os.O_CREATE, 0666)
	if err != nil {
		return err
	}
	buf.WriteTo(file)
	return err
}
