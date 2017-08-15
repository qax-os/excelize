package format

// Table directly maps the format settings of the table.
type Table struct {
	TableStyle        string `json:"table_style"`
	ShowFirstColumn   bool   `json:"show_first_column"`
	ShowLastColumn    bool   `json:"show_last_column"`
	ShowRowStripes    bool   `json:"show_row_stripes"`
	ShowColumnStripes bool   `json:"show_column_stripes"`
}

// AutoFilter directly maps the auto filter settings.
type AutoFilter struct {
	Column     string `json:"column"`
	Expression string `json:"expression"`
	FilterList []struct {
		Column string `json:"column"`
		Value  []int  `json:"value"`
	} `json:"filter_list"`
}
