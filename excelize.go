package excelize

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
)

// File define a populated xlsx.File struct.
type File struct {
	XLSX map[string]string
	Path string
}

// OpenFile take the name of an XLSX file and returns a populated
// xlsx.File struct for it.
func OpenFile(filename string) *File {
	var f *zip.ReadCloser
	file := make(map[string]string)
	f, _ = zip.OpenReader(filename)
	file, _ = ReadZip(f)
	return &File{
		XLSX: file,
		Path: filename,
	}
}

// SetCellInt provide function to set int type value of a cell
func (f *File) SetCellInt(sheet string, axis string, value int) {
	axis = strings.ToUpper(axis)
	var xlsx xlsxWorksheet
	col := getColIndex(axis)
	row := getRowIndex(axis)
	xAxis := row - 1
	yAxis := titleToNumber(col)

	name := `xl/worksheets/` + strings.ToLower(sheet) + `.xml`
	xml.Unmarshal([]byte(f.readXML(name)), &xlsx)

	rows := xAxis + 1
	cell := yAxis + 1

	xlsx = checkRow(xlsx)
	xlsx = completeRow(xlsx, rows, cell)
	xlsx = completeCol(xlsx, rows, cell)

	xlsx.SheetData.Row[xAxis].C[yAxis].T = ""
	xlsx.SheetData.Row[xAxis].C[yAxis].V = strconv.Itoa(value)

	output, err := xml.Marshal(xlsx)
	if err != nil {
		fmt.Println(err)
	}
	f.saveFileList(name, replaceRelationshipsID(replaceWorkSheetsRelationshipsNameSpace(string(output))))
}

// SetCellStr provide function to set string type value of a cell
func (f *File) SetCellStr(sheet string, axis string, value string) {
	axis = strings.ToUpper(axis)
	var xlsx xlsxWorksheet
	col := getColIndex(axis)
	row := getRowIndex(axis)
	xAxis := row - 1
	yAxis := titleToNumber(col)

	name := `xl/worksheets/` + strings.ToLower(sheet) + `.xml`
	xml.Unmarshal([]byte(f.readXML(name)), &xlsx)

	rows := xAxis + 1
	cell := yAxis + 1

	xlsx = checkRow(xlsx)
	xlsx = completeRow(xlsx, rows, cell)
	xlsx = completeCol(xlsx, rows, cell)

	xlsx.SheetData.Row[xAxis].C[yAxis].T = "str"
	xlsx.SheetData.Row[xAxis].C[yAxis].V = value

	output, err := xml.Marshal(xlsx)
	if err != nil {
		fmt.Println(err)
	}
	f.saveFileList(name, replaceRelationshipsID(replaceWorkSheetsRelationshipsNameSpace(string(output))))
}

// Completion column element tags of XML in a sheet
func completeCol(xlsx xlsxWorksheet, row int, cell int) xlsxWorksheet {
	if len(xlsx.SheetData.Row) < cell {
		for i := len(xlsx.SheetData.Row); i < cell; i++ {
			xlsx.SheetData.Row = append(xlsx.SheetData.Row, xlsxRow{
				R: i + 1,
			})
		}
	}
	buffer := bytes.Buffer{}
	for k, v := range xlsx.SheetData.Row {
		if len(v.C) < cell {
			start := len(v.C)
			for iii := start; iii < cell; iii++ {
				buffer.WriteString(toAlphaString(iii + 1))
				buffer.WriteString(strconv.Itoa(k + 1))
				xlsx.SheetData.Row[k].C = append(xlsx.SheetData.Row[k].C, xlsxC{
					R: buffer.String(),
				})
				buffer.Reset()
			}
		}
	}
	return xlsx
}

// Completion row element tags of XML in a sheet
func completeRow(xlsx xlsxWorksheet, row int, cell int) xlsxWorksheet {
	if len(xlsx.SheetData.Row) < row {
		for i := len(xlsx.SheetData.Row); i < row; i++ {
			xlsx.SheetData.Row = append(xlsx.SheetData.Row, xlsxRow{
				R: i + 1,
			})
		}
		buffer := bytes.Buffer{}
		for ii := 0; ii < row; ii++ {
			start := len(xlsx.SheetData.Row[ii].C)
			if start == 0 {
				for iii := start; iii < cell; iii++ {
					buffer.WriteString(toAlphaString(iii + 1))
					buffer.WriteString(strconv.Itoa(ii + 1))
					xlsx.SheetData.Row[ii].C = append(xlsx.SheetData.Row[ii].C, xlsxC{
						R: buffer.String(),
					})
					buffer.Reset()
				}
			}
		}
	}
	return xlsx
}

// Replace xl/worksheets/sheet%d.xml XML tags to self-closing for compatible Office Excel 2007
func replaceWorkSheetsRelationshipsNameSpace(workbookMarshal string) string {
	oldXmlns := `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`
	newXmlns := `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">`
	workbookMarshal = strings.Replace(workbookMarshal, oldXmlns, newXmlns, -1)
	workbookMarshal = strings.Replace(workbookMarshal, `></sheetPr>`, ` />`, -1)
	workbookMarshal = strings.Replace(workbookMarshal, `></dimension>`, ` />`, -1)
	workbookMarshal = strings.Replace(workbookMarshal, `></selection>`, ` />`, -1)
	workbookMarshal = strings.Replace(workbookMarshal, `></sheetFormatPr>`, ` />`, -1)
	workbookMarshal = strings.Replace(workbookMarshal, `></printOptions>`, ` />`, -1)
	workbookMarshal = strings.Replace(workbookMarshal, `></pageSetup>`, ` />`, -1)
	workbookMarshal = strings.Replace(workbookMarshal, `></pageMargins>`, ` />`, -1)
	workbookMarshal = strings.Replace(workbookMarshal, `></headerFooter>`, ` />`, -1)
	workbookMarshal = strings.Replace(workbookMarshal, `></drawing>`, ` />`, -1)
	return workbookMarshal
}

// Check XML tags and fix discontinuous case, for example:
//
//  <row r="15" spans="1:22" x14ac:dyDescent="0.2">
//      <c r="A15" s="2" />
//      <c r="B15" s="2" />
//      <c r="F15" s="1" />
//      <c r="G15" s="1" />
//  </row>
//
// in this case, we should to change it to
//
//  <row r="15" spans="1:22" x14ac:dyDescent="0.2">
//      <c r="A15" s="2" />
//      <c r="B15" s="2" />
//      <c r="C15" s="2" />
//      <c r="D15" s="2" />
//      <c r="E15" s="2" />
//      <c r="F15" s="1" />
//      <c r="G15" s="1" />
//  </row>
//
func checkRow(xlsx xlsxWorksheet) xlsxWorksheet {
	buffer := bytes.Buffer{}
	for k, v := range xlsx.SheetData.Row {
		lenCol := len(v.C)
		if lenCol < 1 {
			continue
		}
		endR := getColIndex(v.C[lenCol-1].R)
		endRow := getRowIndex(v.C[lenCol-1].R)
		endCol := titleToNumber(endR)
		if lenCol < endCol {
			oldRow := xlsx.SheetData.Row[k].C
			xlsx.SheetData.Row[k].C = xlsx.SheetData.Row[k].C[:0]
			tmp := []xlsxC{}
			for i := 0; i <= endCol; i++ {
				buffer.WriteString(toAlphaString(i + 1))
				buffer.WriteString(strconv.Itoa(endRow))
				tmp = append(tmp, xlsxC{
					R: buffer.String(),
				})
				buffer.Reset()
			}
			xlsx.SheetData.Row[k].C = tmp
			for _, y := range oldRow {
				colAxis := titleToNumber(getColIndex(y.R))
				xlsx.SheetData.Row[k].C[colAxis] = y
			}
		}
	}
	return xlsx
}
