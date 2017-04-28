package excelize

import "encoding/xml"

// xlsxTable directly maps the table element. A table helps organize and provide
// structure to lists of information in a worksheet. Tables have clearly labeled
// columns, rows, and data regions. Tables make it easier for users to sort,
// analyze, format, manage, add, and delete information. This element is the
// root element for a table that is not a single cell XML table.
type xlsxTable struct {
	XMLName              xml.Name            `xml:"table"`
	XMLNS                string              `xml:"xmlns,attr"`
	DataCellStyle        string              `xml:"dataCellStyle,attr,omitempty"`
	DataDxfID            int                 `xml:"dataDxfId,attr,omitempty"`
	DisplayName          string              `xml:"displayName,attr,omitempty"`
	HeaderRowBorderDxfID int                 `xml:"headerRowBorderDxfId,attr,omitempty"`
	HeaderRowCellStyle   string              `xml:"headerRowCellStyle,attr,omitempty"`
	HeaderRowCount       int                 `xml:"headerRowCount,attr,omitempty"`
	HeaderRowDxfID       int                 `xml:"headerRowDxfId,attr,omitempty"`
	ID                   int                 `xml:"id,attr"`
	InsertRow            bool                `xml:"insertRow,attr,omitempty"`
	InsertRowShift       bool                `xml:"insertRowShift,attr,omitempty"`
	Name                 string              `xml:"name,attr"`
	Published            bool                `xml:"published,attr,omitempty"`
	Ref                  string              `xml:"ref,attr"`
	TotalsRowCount       int                 `xml:"totalsRowCount,attr,omitempty"`
	TotalsRowDxfID       int                 `xml:"totalsRowDxfId,attr,omitempty"`
	TotalsRowShown       bool                `xml:"totalsRowShown,attr"`
	AutoFilter           *xlsxAutoFilter     `xml:"autoFilter"`
	TableColumns         *xlsxTableColumns   `xml:"tableColumns"`
	TableStyleInfo       *xlsxTableStyleInfo `xml:"tableStyleInfo"`
}

// xlsxAutoFilter temporarily hides rows based on a filter criteria, which is
// applied column by column to a table of data in the worksheet. This collection
// expresses AutoFilter settings.
type xlsxAutoFilter struct {
	Ref string `xml:"ref,attr"`
}

// xlsxTableColumns directly maps the element representing the collection of all
// table columns for this table.
type xlsxTableColumns struct {
	Count       int                `xml:"count,attr"`
	TableColumn []*xlsxTableColumn `xml:"tableColumn"`
}

// xlsxTableColumn directly maps the element representing a single column for
// this table.
type xlsxTableColumn struct {
	DataCellStyle      string `xml:"dataCellStyle,attr,omitempty"`
	DataDxfID          int    `xml:"dataDxfId,attr,omitempty"`
	HeaderRowCellStyle string `xml:"headerRowCellStyle,attr,omitempty"`
	HeaderRowDxfID     int    `xml:"headerRowDxfId,attr,omitempty"`
	ID                 int    `xml:"id,attr"`
	Name               string `xml:"name,attr"`
	QueryTableFieldID  int    `xml:"queryTableFieldId,attr,omitempty"`
	TotalsRowCellStyle string `xml:"totalsRowCellStyle,attr,omitempty"`
	TotalsRowDxfID     int    `xml:"totalsRowDxfId,attr,omitempty"`
	TotalsRowFunction  string `xml:"totalsRowFunction,attr,omitempty"`
	TotalsRowLabel     string `xml:"totalsRowLabel,attr,omitempty"`
	UniqueName         string `xml:"uniqueName,attr,omitempty"`
}

// xlsxTableStyleInfo directly maps the tableStyleInfo element. This element
// describes which style is used to display this table, and specifies which
// portions of the table have the style applied.
type xlsxTableStyleInfo struct {
	Name              string `xml:"name,attr,omitempty"`
	ShowFirstColumn   bool   `xml:"showFirstColumn,attr"`
	ShowLastColumn    bool   `xml:"showLastColumn,attr"`
	ShowRowStripes    bool   `xml:"showRowStripes,attr"`
	ShowColumnStripes bool   `xml:"showColumnStripes,attr"`
}

// formatTable directly maps the format settings of the table.
type formatTable struct {
	TableStyle        string `json:"table_style"`
	ShowFirstColumn   bool   `json:"show_first_column"`
	ShowLastColumn    bool   `json:"show_last_column"`
	ShowRowStripes    bool   `json:"show_row_stripes"`
	ShowColumnStripes bool   `json:"show_column_stripes"`
}
