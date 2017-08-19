package format

import (
	"encoding/json"
)

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

// NewTable provides function to parse the format settings of the table with default value.
func NewTable(f interface{})(*Table, error) {
	def := Table{
		TableStyle: "",
		ShowRowStripes: true,
	}

	var s []byte

	switch t := f.(type) {
	case string:
		s = []byte(t)
	case Table:
		s, _ = json.Marshal(t)
	case *Table:
		s, _ = json.Marshal(t)
	default:
		return &def, unknownFormat
	}

	c := &def
	err := json.Unmarshal(s, c)
	return c, err
}

// NewAutoFilter provides function to parse the settings of the auto filter
func NewAutoFilter(f interface{})(*AutoFilter, error) {
	switch t := f.(type) {
	case string:
		fs := AutoFilter{}
		err := json.Unmarshal([]byte(t), &fs)
		return &fs, err
	case AutoFilter:
		return &t, nil
	case *AutoFilter:
		return &(*t), nil
	default:
		return nil, unknownFormat
	}
}