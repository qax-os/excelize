package format

// Panes directly maps the settings of the panes.
type Panes struct {
	Freeze      bool   `json:"freeze"`
	Split       bool   `json:"split"`
	XSplit      int    `json:"x_split"`
	YSplit      int    `json:"y_split"`
	TopLeftCell string `json:"top_left_cell"`
	ActivePane  string `json:"active_pane"`
	Panes       []struct {
		SQRef      string `json:"sqref"`
		ActiveCell string `json:"active_cell"`
		Pane       string `json:"pane"`
	} `json:"panes"`
}

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

