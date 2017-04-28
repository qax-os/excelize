package excelize

import (
	"encoding/json"
	"encoding/xml"
	"strconv"
	"strings"
)

// parseFormatTableSet provides function to parse the format settings of the
// table with default value.
func parseFormatTableSet(formatSet string) *formatTable {
	format := formatTable{
		TableStyle:     "",
		ShowRowStripes: true,
	}
	json.Unmarshal([]byte(formatSet), &format)
	return &format
}

// AddTable provides the method to add table in a worksheet by given sheet
// index, coordinate area and format set. For example, create a table of A1:D5
// on Sheet1:
//
//    xlsx.AddTable("Sheet1", "A1", "D5", ``)
//
// Create a table of F2:H6 on Sheet2 with format set:
//
//    xlsx.AddTable("Sheet2", "F2", "H6", `{"table_style":"TableStyleMedium2", "show_first_column":true,"show_last_column":true,"show_row_stripes":false,"show_column_stripes":true}`)
//
// Note that the table at least two lines include string type header. The two
// chart coordinate areas can not have an intersection.
//
// table_style: The built-in table style names
//
//    TableStyleLight1 - TableStyleLight21
//    TableStyleMedium1 - TableStyleMedium28
//    TableStyleDark1 - TableStyleDark11
//
func (f *File) AddTable(sheet, hcell, vcell, format string) {
	formatSet := parseFormatTableSet(format)
	hcell = strings.ToUpper(hcell)
	vcell = strings.ToUpper(vcell)
	// Coordinate conversion, convert C1:B3 to 2,0,1,2.
	hcol := string(strings.Map(letterOnlyMapF, hcell))
	hrow, _ := strconv.Atoi(strings.Map(intOnlyMapF, hcell))
	hyAxis := hrow - 1
	hxAxis := titleToNumber(hcol)

	vcol := string(strings.Map(letterOnlyMapF, vcell))
	vrow, _ := strconv.Atoi(strings.Map(intOnlyMapF, vcell))
	vyAxis := vrow - 1
	vxAxis := titleToNumber(vcol)
	if vxAxis < hxAxis {
		hcell, vcell = vcell, hcell
		vxAxis, hxAxis = hxAxis, vxAxis
	}
	if vyAxis < hyAxis {
		hcell, vcell = vcell, hcell
		vyAxis, hyAxis = hyAxis, vyAxis
	}
	tableID := f.countTables() + 1
	sheetRelationshipsTableXML := "../tables/table" + strconv.Itoa(tableID) + ".xml"
	tableXML := strings.Replace(sheetRelationshipsTableXML, "..", "xl", -1)
	// Add first table for given sheet.
	rID := f.addSheetRelationships(sheet, SourceRelationshipTable, sheetRelationshipsTableXML, "")
	f.addSheetTable(sheet, rID)
	f.addTable(sheet, tableXML, hxAxis, hyAxis, vxAxis, vyAxis, tableID, formatSet)
	f.addTableContentTypePart(tableID)
}

// countTables provides function to get table files count storage in the folder
// xl/tables.
func (f *File) countTables() int {
	count := 0
	for k := range f.XLSX {
		if strings.Contains(k, "xl/tables/table") {
			count++
		}
	}
	return count
}

// addSheetTable provides function to add tablePart element to
// xl/worksheets/sheet%d.xml by given sheet name and relationship index.
func (f *File) addSheetTable(sheet string, rID int) {
	xlsx := f.workSheetReader(sheet)
	table := &xlsxTablePart{
		RID: "rId" + strconv.Itoa(rID),
	}
	if xlsx.TableParts != nil {
		xlsx.TableParts.Count++
		xlsx.TableParts.TableParts = append(xlsx.TableParts.TableParts, table)
	} else {
		xlsx.TableParts = &xlsxTableParts{
			Count:      1,
			TableParts: []*xlsxTablePart{table},
		}
	}

}

// addTable provides function to add table by given sheet index, coordinate area
// and format set.
func (f *File) addTable(sheet, tableXML string, hxAxis, hyAxis, vxAxis, vyAxis, i int, formatSet *formatTable) {
	// Correct the minimum number of rows, the table at least two lines.
	if hyAxis == vyAxis {
		vyAxis++
	}
	// Correct table reference coordinate area, such correct C1:B3 to B1:C3.
	ref := toAlphaString(hxAxis+1) + strconv.Itoa(hyAxis+1) + ":" + toAlphaString(vxAxis+1) + strconv.Itoa(vyAxis+1)
	tableColumn := []*xlsxTableColumn{}
	idx := 0
	for i := hxAxis; i <= vxAxis; i++ {
		idx++
		cell := toAlphaString(i+1) + strconv.Itoa(hyAxis+1)
		name := f.GetCellValue(sheet, cell)
		if _, err := strconv.Atoi(name); err == nil {
			f.SetCellStr(sheet, cell, name)
		}
		if name == "" {
			name = "Column" + strconv.Itoa(idx)
			f.SetCellStr(sheet, cell, name)
		}
		tableColumn = append(tableColumn, &xlsxTableColumn{
			ID:   idx,
			Name: name,
		})
	}
	name := "Table" + strconv.Itoa(i)
	t := xlsxTable{
		XMLNS:       NameSpaceSpreadSheet,
		ID:          i,
		Name:        name,
		DisplayName: name,
		Ref:         ref,
		AutoFilter: &xlsxAutoFilter{
			Ref: ref,
		},
		TableColumns: &xlsxTableColumns{
			Count:       idx,
			TableColumn: tableColumn,
		},
		TableStyleInfo: &xlsxTableStyleInfo{
			Name:              formatSet.TableStyle,
			ShowFirstColumn:   formatSet.ShowFirstColumn,
			ShowLastColumn:    formatSet.ShowLastColumn,
			ShowRowStripes:    formatSet.ShowRowStripes,
			ShowColumnStripes: formatSet.ShowColumnStripes,
		},
	}
	table, _ := xml.Marshal(t)
	f.saveFileList(tableXML, string(table))
}

// addTableContentTypePart provides function to add image part relationships
// in the file [Content_Types].xml by given drawing index.
func (f *File) addTableContentTypePart(index int) {
	f.setContentTypePartImageExtensions()
	content := f.contentTypesReader()
	for _, v := range content.Overrides {
		if v.PartName == "/xl/tables/table"+strconv.Itoa(index)+".xml" {
			return
		}
	}
	content.Overrides = append(content.Overrides, xlsxOverride{
		PartName:    "/xl/tables/table" + strconv.Itoa(index) + ".xml",
		ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
	})
}
