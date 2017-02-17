package excelize

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"io"
	"io/ioutil"
	"os"
	"strconv"
	"strings"
)

// File define a populated XLSX file struct.
type File struct {
	checked    map[string]bool
	XLSX       map[string]string
	Path       string
	SheetCount int
}

// OpenFile take the name of an XLSX file and returns a populated XLSX file
// struct for it.
func OpenFile(filename string) (*File, error) {
	file, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	defer file.Close()
	f, err := OpenReader(file)
	if err != nil {
		return nil, err
	}
	f.Path = filename
	return f, nil
}

// OpenReader take an io.Reader and return a populated XLSX file.
func OpenReader(r io.Reader) (*File, error) {
	b, err := ioutil.ReadAll(r)
	if err != nil {
		return nil, err
	}

	zr, err := zip.NewReader(bytes.NewReader(b), int64(len(b)))
	if err != nil {
		return nil, err
	}

	file, sheetCount, err := ReadZipReader(zr)
	if err != nil {
		return nil, err
	}
	return &File{
		checked:    make(map[string]bool),
		XLSX:       file,
		Path:       "",
		SheetCount: sheetCount,
	}, nil
}

// SetCellValue provides function to set int or string type value of a cell.
func (f *File) SetCellValue(sheet string, axis string, value interface{}) {
	switch t := value.(type) {
	case int:
		f.SetCellInt(sheet, axis, value.(int))
	case int8:
		f.SetCellInt(sheet, axis, int(value.(int8)))
	case int16:
		f.SetCellInt(sheet, axis, int(value.(int16)))
	case int32:
		f.SetCellInt(sheet, axis, int(value.(int32)))
	case int64:
		f.SetCellInt(sheet, axis, int(value.(int64)))
	case float32:
		f.SetCellDefault(sheet, axis, strconv.FormatFloat(float64(value.(float32)), 'f', -1, 32))
	case float64:
		f.SetCellDefault(sheet, axis, strconv.FormatFloat(float64(value.(float64)), 'f', -1, 64))
	case string:
		f.SetCellStr(sheet, axis, t)
	case []byte:
		f.SetCellStr(sheet, axis, string(t))
	default:
		f.SetCellStr(sheet, axis, "")
	}
}

// SetCellInt provides function to set int type value of a cell.
func (f *File) SetCellInt(sheet string, axis string, value int) {
	axis = strings.ToUpper(axis)
	var xlsx xlsxWorksheet
	name := "xl/worksheets/" + strings.ToLower(sheet) + ".xml"
	xml.Unmarshal([]byte(f.readXML(name)), &xlsx)
	if f.checked == nil {
		f.checked = make(map[string]bool)
	}
	ok := f.checked[name]
	if !ok {
		xlsx = checkRow(xlsx)
		f.checked[name] = true
	}

	if xlsx.MergeCells != nil {
		for i := 0; i < len(xlsx.MergeCells.Cells); i++ {
			if checkCellInArea(axis, xlsx.MergeCells.Cells[i].Ref) {
				axis = strings.Split(xlsx.MergeCells.Cells[i].Ref, ":")[0]
			}
		}
	}
	col := string(strings.Map(letterOnlyMapF, axis))
	row, _ := strconv.Atoi(strings.Map(intOnlyMapF, axis))
	xAxis := row - 1
	yAxis := titleToNumber(col)

	rows := xAxis + 1
	cell := yAxis + 1

	xlsx = completeRow(xlsx, rows, cell)
	xlsx = completeCol(xlsx, rows, cell)

	xlsx.SheetData.Row[xAxis].C[yAxis].T = ""
	xlsx.SheetData.Row[xAxis].C[yAxis].V = strconv.Itoa(value)

	output, _ := xml.Marshal(xlsx)
	f.saveFileList(name, replaceWorkSheetsRelationshipsNameSpace(string(output)))
}

// SetCellStr provides function to set string type value of a cell. Total number
// of characters that a cell can contain 32767 characters.
func (f *File) SetCellStr(sheet string, axis string, value string) {
	axis = strings.ToUpper(axis)
	var xlsx xlsxWorksheet
	name := "xl/worksheets/" + strings.ToLower(sheet) + ".xml"
	xml.Unmarshal([]byte(f.readXML(name)), &xlsx)
	if f.checked == nil {
		f.checked = make(map[string]bool)
	}
	ok := f.checked[name]
	if !ok {
		xlsx = checkRow(xlsx)
		f.checked[name] = true
	}
	if xlsx.MergeCells != nil {
		for i := 0; i < len(xlsx.MergeCells.Cells); i++ {
			if checkCellInArea(axis, xlsx.MergeCells.Cells[i].Ref) {
				axis = strings.Split(xlsx.MergeCells.Cells[i].Ref, ":")[0]
			}
		}
	}
	if len(value) > 32767 {
		value = value[0:32767]
	}
	col := string(strings.Map(letterOnlyMapF, axis))
	row, _ := strconv.Atoi(strings.Map(intOnlyMapF, axis))
	xAxis := row - 1
	yAxis := titleToNumber(col)

	rows := xAxis + 1
	cell := yAxis + 1

	xlsx = completeRow(xlsx, rows, cell)
	xlsx = completeCol(xlsx, rows, cell)

	xlsx.SheetData.Row[xAxis].C[yAxis].T = "str"
	xlsx.SheetData.Row[xAxis].C[yAxis].V = value

	output, _ := xml.Marshal(xlsx)
	f.saveFileList(name, replaceWorkSheetsRelationshipsNameSpace(string(output)))
}

// SetCellDefault provides function to set string type value of a cell as
// default format without escaping the cell.
func (f *File) SetCellDefault(sheet string, axis string, value string) {
	axis = strings.ToUpper(axis)
	var xlsx xlsxWorksheet
	name := "xl/worksheets/" + strings.ToLower(sheet) + ".xml"
	xml.Unmarshal([]byte(f.readXML(name)), &xlsx)
	if f.checked == nil {
		f.checked = make(map[string]bool)
	}
	ok := f.checked[name]
	if !ok {
		xlsx = checkRow(xlsx)
		f.checked[name] = true
	}
	if xlsx.MergeCells != nil {
		for i := 0; i < len(xlsx.MergeCells.Cells); i++ {
			if checkCellInArea(axis, xlsx.MergeCells.Cells[i].Ref) {
				axis = strings.Split(xlsx.MergeCells.Cells[i].Ref, ":")[0]
			}
		}
	}
	col := string(strings.Map(letterOnlyMapF, axis))
	row, _ := strconv.Atoi(strings.Map(intOnlyMapF, axis))
	xAxis := row - 1
	yAxis := titleToNumber(col)

	rows := xAxis + 1
	cell := yAxis + 1

	xlsx = completeRow(xlsx, rows, cell)
	xlsx = completeCol(xlsx, rows, cell)

	xlsx.SheetData.Row[xAxis].C[yAxis].T = ""
	xlsx.SheetData.Row[xAxis].C[yAxis].V = value

	output, _ := xml.Marshal(xlsx)
	f.saveFileList(name, replaceWorkSheetsRelationshipsNameSpace(string(output)))
}

// Completion column element tags of XML in a sheet.
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

// Completion row element tags of XML in a sheet.
func completeRow(xlsx xlsxWorksheet, row int, cell int) xlsxWorksheet {
	currentRows := len(xlsx.SheetData.Row)
	if currentRows > 1 {
		lastRow := xlsx.SheetData.Row[currentRows-1].R
		if lastRow >= row {
			row = lastRow
		}
	}
	sheetData := xlsxSheetData{}
	existsRows := map[int]int{}
	for k, v := range xlsx.SheetData.Row {
		existsRows[v.R] = k
	}
	for i := 0; i < row; i++ {
		_, ok := existsRows[i+1]
		if ok {
			sheetData.Row = append(sheetData.Row, xlsx.SheetData.Row[existsRows[i+1]])
			continue
		}
		sheetData.Row = append(sheetData.Row, xlsxRow{
			R: i + 1,
		})
	}
	buffer := bytes.Buffer{}
	for ii := 0; ii < row; ii++ {
		start := len(sheetData.Row[ii].C)
		if start == 0 {
			for iii := start; iii < cell; iii++ {
				buffer.WriteString(toAlphaString(iii + 1))
				buffer.WriteString(strconv.Itoa(ii + 1))
				sheetData.Row[ii].C = append(sheetData.Row[ii].C, xlsxC{
					R: buffer.String(),
				})
				buffer.Reset()
			}
		}
	}
	xlsx.SheetData = sheetData
	return xlsx
}

// Replace xl/worksheets/sheet%d.xml XML tags to self-closing for compatible
// Office Excel 2007.
func replaceWorkSheetsRelationshipsNameSpace(workbookMarshal string) string {
	oldXmlns := `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`
	newXmlns := `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">`
	workbookMarshal = strings.Replace(workbookMarshal, oldXmlns, newXmlns, -1)
	return workbookMarshal
}

// Check XML tags and fix discontinuous case. For example:
//
//    <row r="15" spans="1:22" x14ac:dyDescent="0.2">
//        <c r="A15" s="2" />
//        <c r="B15" s="2" />
//        <c r="F15" s="1" />
//        <c r="G15" s="1" />
//    </row>
//
// in this case, we should to change it to
//
//    <row r="15" spans="1:22" x14ac:dyDescent="0.2">
//        <c r="A15" s="2" />
//        <c r="B15" s="2" />
//        <c r="C15" s="2" />
//        <c r="D15" s="2" />
//        <c r="E15" s="2" />
//        <c r="F15" s="1" />
//        <c r="G15" s="1" />
//    </row>
//
// Noteice: this method could be very slow for large spreadsheets (more than
// 3000 rows one sheet).
func checkRow(xlsx xlsxWorksheet) xlsxWorksheet {
	buffer := bytes.Buffer{}
	for k, v := range xlsx.SheetData.Row {
		lenCol := len(v.C)
		if lenCol < 1 {
			continue
		}
		endR := string(strings.Map(letterOnlyMapF, v.C[lenCol-1].R))
		endRow, _ := strconv.Atoi(strings.Map(intOnlyMapF, v.C[lenCol-1].R))
		endCol := titleToNumber(endR) + 1
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
				colAxis := titleToNumber(string(strings.Map(letterOnlyMapF, y.R)))
				xlsx.SheetData.Row[k].C[colAxis] = y
			}
		}
	}
	return xlsx
}

// UpdateLinkedValue fix linked values within a spreadsheet are not updating in
// Office Excel 2007 and 2010. This function will be remove value tag when met a
// cell have a linked value. Reference
// https://social.technet.microsoft.com/Forums/office/en-US/e16bae1f-6a2c-4325-8013-e989a3479066/excel-2010-linked-cells-not-updating?forum=excel
//
// Notice: after open XLSX file Excel will be update linked value and generate
// new value and will prompt save file or not.
//
// For example:
//
//    <row r="19" spans="2:2">
//        <c r="B19">
//            <f>SUM(Sheet2!D2,Sheet2!D11)</f>
//            <v>100</v>
//         </c>
//    </row>
//
// to
//
//    <row r="19" spans="2:2">
//        <c r="B19">
//            <f>SUM(Sheet2!D2,Sheet2!D11)</f>
//        </c>
//    </row>
//
func (f *File) UpdateLinkedValue() {
	for i := 1; i <= f.SheetCount; i++ {
		var xlsx xlsxWorksheet
		name := "xl/worksheets/sheet" + strconv.Itoa(i) + ".xml"
		xml.Unmarshal([]byte(f.readXML(name)), &xlsx)
		for indexR, row := range xlsx.SheetData.Row {
			for indexC, col := range row.C {
				if col.F != nil && col.V != "" {
					xlsx.SheetData.Row[indexR].C[indexC].V = ""
					xlsx.SheetData.Row[indexR].C[indexC].T = ""
				}
			}
		}
		output, _ := xml.Marshal(xlsx)
		f.saveFileList(name, replaceWorkSheetsRelationshipsNameSpace(string(output)))
	}
}
