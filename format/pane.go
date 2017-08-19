package format

import (
	"encoding/json"
)

type Pane struct {
	SQRef      string `json:"sqref"`
	ActiveCell string `json:"active_cell"`
	Pane       string `json:"pane"`
}

// Panes directly maps the settings of the panes.
type Panes struct {
	Freeze      bool   `json:"freeze"`
	Split       bool   `json:"split"`
	XSplit      int    `json:"x_split"`
	YSplit      int    `json:"y_split"`
	TopLeftCell string `json:"top_left_cell"`
	ActivePane  string `json:"active_pane"`
	Panes       []Pane `json:"panes"`
}

// NewPanes provides function to parse the panes settings.
func NewPanes(f interface{})(*Panes, error) {
	switch t := f.(type) {
	case string:
		fs := Panes{}
		err := json.Unmarshal([]byte(t), &fs)
		return &fs, err
	case Panes:
		return &t, nil
	case *Panes:
		return &(*t), nil
	default:
		return nil, unknownFormat
	}
}