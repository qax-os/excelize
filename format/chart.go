package format

import (
	"encoding/json"
)

// formatChartAxis directly maps the format settings of the chart axis.
type ChartAxis struct {
	Crossing            string `json:"crossing"`
	MajorTickMark       string `json:"major_tick_mark"`
	MinorTickMark       string `json:"minor_tick_mark"`
	MinorUnitType       string `json:"minor_unit_type"`
	MajorUnit           int    `json:"major_unit"`
	MajorUnitType       string `json:"major_unit_type"`
	DisplayUnits        string `json:"display_units"`
	DisplayUnitsVisible bool   `json:"display_units_visible"`
	DateAxis            bool   `json:"date_axis"`
	NumFormat           string `json:"num_format"`
	NumFont struct {
		Color     string `json:"color"`
		Bold      bool   `json:"bold"`
		Italic    bool   `json:"italic"`
		Underline bool   `json:"underline"`
	} `json:"num_font"`
	NameLayout Layout `json:"name_layout"`
}

// formatChart directly maps the format settings of the chart.
type Chart struct {
	Type           string        `json:"type"`
	Series         []ChartSeries `json:"series"`
	Format         Picture       `json:"format"`
	Legend         ChartLegend   `json:"legend"`
	Title          ChartTitle    `json:"title"`
	XAxis          ChartAxis     `json:"x_axis"`
	YAxis          ChartAxis     `json:"y_axis"`
	Area           ChartArea     `json:"chartarea"`
	Plot           PlotArea      `json:"plotarea"`
	ShowBlanksAs   string `json:"show_blanks_as"`
	ShowHiddenData bool   `json:"show_hidden_data"`
	SetRotation    int    `json:"set_rotation"`
	SetHoleSize    int    `json:"set_hole_size"`
}

type ChartArea struct {
	Border struct {
		None bool `json:"none"`
	} `json:"border"`
	Fill struct {
		Color string `json:"color"`
	} `json:"fill"`
	Pattern struct {
		Pattern string `json:"pattern"`
		FgColor string `json:"fg_color"`
		BgColor string `json:"bg_color"`
	} `json:"pattern"`
}

type PlotArea struct {
	ShowBubbleSize  bool `json:"show_bubble_size"`
	ShowCatName     bool `json:"show_cat_name"`
	ShowLeaderLines bool `json:"show_leader_lines"`
	ShowPercent     bool `json:"show_percent"`
	ShowSerName     bool `json:"show_series_name"`
	ShowVal         bool `json:"show_val"`
	Gradient struct {
		Colors []string `json:"colors"`
	} `json:"gradient"`
	Border struct {
		Color    string `json:"color"`
		Width    int    `json:"width"`
		DashType string `json:"dash_type"`
	} `json:"border"`
	Fill struct {
		Color string `json:"color"`
	} `json:"fill"`
	Layout Layout `json:"layout"`
}

// formatChartLegend directly maps the format settings of the chart legend.
type ChartLegend struct {
	None            bool         `json:"none"`
	DeleteSeries    []int        `json:"delete_series"`
	Font            Font   `json:"font"`
	Layout          Layout `json:"layout"`
	Position        string       `json:"position"`
	ShowLegendEntry bool         `json:"show_legend_entry"`
	ShowLegendKey   bool         `json:"show_legend_key"`
}

// formatChartSeries directly maps the format settings of the chart series.
type ChartSeries struct {
	Name       string `json:"name"`
	Categories string `json:"categories"`
	Values     string `json:"values"`
	Line struct {
		None  bool   `json:"none"`
		Color string `json:"color"`
	} `json:"line"`
	Marker struct {
		Type  string  `json:"type"`
		Size  int     `json:"size,"`
		Width float64 `json:"width"`
		Border struct {
			Color string `json:"color"`
			None  bool   `json:"none"`
		} `json:"border"`
		Fill struct {
			Color string `json:"color"`
			None  bool   `json:"none"`
		} `json:"fill"`
	} `json:"marker"`
}

// ChartTitle directly maps the format settings of the chart title.
type ChartTitle struct {
	None    bool         `json:"none"`
	Name    string       `json:"name"`
	Overlay bool         `json:"overlay"`
	Layout  Layout 		 `json:"layout"`
}

// Layout directly maps the format settings of the element layout.
type Layout struct {
	X      float64 `json:"x"`
	Y      float64 `json:"y"`
	Width  float64 `json:"width"`
	Height float64 `json:"height"`
}

// NewChart provides function to parse the format settings of the
func NewChart(f interface{})(*Chart, error) {
	def := Chart{
		Format: Picture{
			FPrintsWithSheet: true,
			FLocksWithSheet:  false,
			NoChangeAspect:   false,
			OffsetX:          0,
			OffsetY:          0,
			XScale:           1.0,
			YScale:           1.0,
		},
		Legend: ChartLegend{
			Position:      "bottom",
			ShowLegendKey: false,
		},
		Title: ChartTitle{
			Name: " ",
		},
		ShowBlanksAs: "gap",
	}

	var s []byte

	switch t := f.(type) {
	case string:
		s = []byte(t)
	case Chart:
		s, _ = json.Marshal(t)
	case *Chart:
		s, _ = json.Marshal(t)
	default:
		return &def, unknownFormat
	}

	c := &def
	err := json.Unmarshal(s, c)
	return c, err
}
