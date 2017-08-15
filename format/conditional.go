package format

import (
	"encoding/json"
)

// Conditional directly maps the conditional format settings of the cells.
type Conditional struct {
	Type         string `json:"type"`
	AboveAverage bool   `json:"above_average"`
	Percent      bool   `json:"percent"`
	Format       int    `json:"format"`
	Criteria     string `json:"criteria"`
	Value        string `json:"value,omitempty"`
	Minimum      string `json:"minimum,omitempty"`
	Maximum      string `json:"maximum,omitempty"`
	MinType      string `json:"min_type,omitempty"`
	MidType      string `json:"mid_type,omitempty"`
	MaxType      string `json:"max_type,omitempty"`
	MinValue     string `json:"min_value,omitempty"`
	MidValue     string `json:"mid_value,omitempty"`
	MaxValue     string `json:"max_value,omitempty"`
	MinColor     string `json:"min_color,omitempty"`
	MidColor     string `json:"mid_color,omitempty"`
	MaxColor     string `json:"max_color,omitempty"`
	MinLength    string `json:"min_length,omitempty"`
	MaxLength    string `json:"max_length,omitempty"`
	MultiRange   string `json:"multi_range,omitempty"`
	BarColor     string `json:"bar_color,omitempty"`
}

// NewConditional provides function to parse the settings of the conditional format
func NewConditional(f interface{})([]*Conditional, error) {
	switch t := f.(type) {
	case string:
		fs := []*Conditional{}
		err := json.Unmarshal([]byte(t), &fs)
		return fs, err
	case Conditional:
		return []*Conditional{&t}, nil
	case *Conditional:
		return []*Conditional{&(*t)}, nil
	case []Conditional:
		fs := []*Conditional{}
		for _, c := range t {
			fs = append(fs, &c)
		}
		return fs, nil
	case []*Conditional:
		return t, nil
	default:
		return nil, unknownFormat
	}
}
