package excelize

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
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
		XLSX:  file,
		Sheet: make(map[string]*xlsxWorksheet),
	}
}

// Save provides function to override the xlsx file with origin path.
func (f *File) Save() error {
	if f.Path == "" {
		return fmt.Errorf("No path defined for file, consider File.WriteTo or File.Write")
	}
	return f.WriteTo(f.Path)
}

// WriteTo provides function to create or update to an xlsx file at the provided
// path.
func (f *File) WriteTo(name string) error {
	file, err := os.OpenFile(name, os.O_WRONLY|os.O_TRUNC|os.O_CREATE, 0666)
	if err != nil {
		return err
	}
	defer file.Close()
	return f.Write(file)
}

// Write provides function to write to an io.Writer.
func (f *File) Write(w io.Writer) error {
	buf := new(bytes.Buffer)
	zw := zip.NewWriter(buf)
	for path, sheet := range f.Sheet {
		if sheet == nil {
			continue
		}
		output, err := xml.Marshal(sheet)
		if err != nil {
			return err
		}
		f.saveFileList(path, replaceWorkSheetsRelationshipsNameSpace(string(output)))
	}
	for path, content := range f.XLSX {
		fi, err := zw.Create(path)
		if err != nil {
			return err
		}
		_, err = fi.Write([]byte(content))
		if err != nil {
			return err
		}
	}
	err := zw.Close()
	if err != nil {
		return err
	}

	if _, err := buf.WriteTo(w); err != nil {
		return err
	}

	return nil
}
