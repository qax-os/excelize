package excelize

import (
	"encoding/xml"
	"strconv"
	"strings"
)

// GetRows return all the rows in a sheet, for example:
//
//    rows := xlsx.GetRows("Sheet2")
//    for _, row := range rows {
//        for _, colCell := range row {
//            fmt.Print(colCell, "\t")
//        }
//        fmt.Println()
//    }
//
func (f *File) GetRows(sheet string) [][]string {
	rows := [][]string{}
	name := "xl/worksheets/" + strings.ToLower(sheet) + ".xml"
	decoder := xml.NewDecoder(strings.NewReader(f.readXML(name)))
	d, err := readXMLSST(f)
	if err != nil {
		return rows
	}
	var inElement string
	var r xlsxRow
	var row []string
	tr, tc := f.getTotalRowsCols(sheet)
	for i := 0; i < tr; i++ {
		row = []string{}
		for j := 0; j <= tc; j++ {
			row = append(row, "")
		}
		rows = append(rows, row)
	}
	decoder = xml.NewDecoder(strings.NewReader(f.readXML(name)))
	for {
		token, _ := decoder.Token()
		if token == nil {
			break
		}
		switch startElement := token.(type) {
		case xml.StartElement:
			inElement = startElement.Name.Local
			if inElement == "row" {
				r = xlsxRow{}
				decoder.DecodeElement(&r, &startElement)
				cr := r.R - 1
				for _, colCell := range r.C {
					c := titleToNumber(strings.Map(letterOnlyMapF, colCell.R))
					val, _ := colCell.getValueFrom(f, d)
					rows[cr][c] = val
				}
			}
		default:
		}
	}
	return rows
}

// getTotalRowsCols provides a function to get total columns and rows in a
// sheet.
func (f *File) getTotalRowsCols(sheet string) (int, int) {
	name := "xl/worksheets/" + strings.ToLower(sheet) + ".xml"
	decoder := xml.NewDecoder(strings.NewReader(f.readXML(name)))
	var inElement string
	var r xlsxRow
	var tr, tc int
	for {
		token, _ := decoder.Token()
		if token == nil {
			break
		}
		switch startElement := token.(type) {
		case xml.StartElement:
			inElement = startElement.Name.Local
			if inElement == "row" {
				r = xlsxRow{}
				decoder.DecodeElement(&r, &startElement)
				tr = r.R
				for _, colCell := range r.C {
					col := titleToNumber(strings.Map(letterOnlyMapF, colCell.R))
					if col > tc {
						tc = col
					}
				}
			}
		default:
		}
	}
	return tr, tc
}

// SetRowHeight provides a function to set the height of a single row.
// For example:
//
//    xlsx := excelize.CreateFile()
//    xlsx.SetRowHeight("Sheet1", 0, 50)
//    err := xlsx.Save()
//    if err != nil {
//        fmt.Println(err)
//        os.Exit(1)
//    }
//
func (f *File) SetRowHeight(sheet string, rowIndex int, height float64) {
	xlsx := xlsxWorksheet{}
	name := "xl/worksheets/" + strings.ToLower(sheet) + ".xml"
	xml.Unmarshal([]byte(f.readXML(name)), &xlsx)

	rows := rowIndex + 1
	cells := 0

	completeRow(&xlsx, rows, cells)

	xlsx.SheetData.Row[rowIndex].Ht = strconv.FormatFloat(height, 'f', -1, 64)
	xlsx.SheetData.Row[rowIndex].CustomHeight = true

	output, _ := xml.Marshal(xlsx)
	f.saveFileList(name, replaceWorkSheetsRelationshipsNameSpace(string(output)))
}

// readXMLSST read xmlSST simple function.
func readXMLSST(f *File) (*xlsxSST, error) {
	shardStrings := xlsxSST{}
	err := xml.Unmarshal([]byte(f.readXML("xl/sharedStrings.xml")), &shardStrings)
	return &shardStrings, err
}

// getValueFrom return a value from a column/row cell, this function is inteded
// to be used with for range on rows an argument with the xlsx opened file.
func (xlsx *xlsxC) getValueFrom(f *File, d *xlsxSST) (string, error) {
	switch xlsx.T {
	case "s":
		xlsxSI := 0
		xlsxSI, _ = strconv.Atoi(xlsx.V)
		if len(d.SI[xlsxSI].R) > 0 {
			value := ""
			for _, v := range d.SI[xlsxSI].R {
				value += v.T
			}
			return value, nil
		}
		return d.SI[xlsxSI].T, nil
	case "str":
		return xlsx.V, nil
	default:
		return xlsx.V, nil
	}
}
