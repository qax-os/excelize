package format

import (
	"encoding/json"
	"errors"
)

var unknownFormat = errors.New("Unknown format")

type Font struct {
	Bold      bool   `json:"bold"`
	Italic    bool   `json:"italic"`
	Underline string `json:"underline"`
	Family    string `json:"family"`
	Size      int    `json:"size"`
	Color     string `json:"color"`
}

type Border struct {
	Type  string `json:"type"`
	Color string `json:"color"`
	Style int    `json:"style"`
}

type Fill struct {
	Type    string   `json:"type"`
	Pattern int      `json:"pattern"`
	Color   []string `json:"color"`
	Shading int      `json:"shading"`
}

type Alignment struct {
	Horizontal      string `json:"horizontal"`
	Indent          int    `json:"indent"`
	JustifyLastLine bool   `json:"justify_last_line"`
	ReadingOrder    uint64 `json:"reading_order"`
	RelativeIndent  int    `json:"relative_indent"`
	ShrinkToFit     bool   `json:"shrink_to_fit"`
	TextRotation    int    `json:"text_rotation"`
	Vertical        string `json:"vertical"`
	WrapText        bool   `json:"wrap_text"`
}

type Style struct {
	Border        []Border `json:"border"`
	Fill          Fill `json:"fill"`
	Font          Font `json:"font"`
	Alignment     Alignment `json:"alignment"`
	NumFmt        int     `json:"number_format"`
	DecimalPlaces int     `json:"decimal_places"`
	CustomNumFmt  string `json:"custom_number_format"`
	Lang          string  `json:"lang"`
	NegRed        bool    `json:"negred"`
}

func NewStyleSet(f interface{})(*Style, error) {
	switch t := f.(type) {
	case string:
		s := &Style{}
		err := json.Unmarshal([]byte(t), s)
		return s, err
	case Style:
		return &t, nil
	case *Style:
		return &(*t), nil
	default:
		return nil, unknownFormat
	}
}