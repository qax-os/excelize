package format

import (
	"encoding/json"
)

// formatPicture directly maps the format settings of the picture.
type Picture struct {
	FPrintsWithSheet bool    `json:"print_obj"`
	FLocksWithSheet  bool    `json:"locked"`
	NoChangeAspect   bool    `json:"lock_aspect_ratio"`
	OffsetX          int     `json:"x_offset"`
	OffsetY          int     `json:"y_offset"`
	XScale           float64 `json:"x_scale"`
	YScale           float64 `json:"y_scale"`
}

// formatShape directly maps the format settings of the shape.
type Shape struct {
	Type      string                 `json:"type"`
	Width     int                    `json:"width"`
	Height    int                    `json:"height"`
	Format    Picture          `json:"format"`
	Color     ShapeColor       `json:"color"`
	Paragraph []ShapeParagraph `json:"paragraph"`
}

// formatShapeParagraph directly maps the format settings of the paragraph in
// the shape.
type ShapeParagraph struct {
	Font Font `json:"font"`
	Text string     `json:"text"`
}

// formatShapeColor directly maps the color settings of the shape.
type ShapeColor struct {
	Line   string `json:"line"`
	Fill   string `json:"fill"`
	Effect string `json:"effect"`
}

// NewPicture provides function to parse the format settings of the picture with default value.
func NewPicture(f interface{}) (*Picture, error) {
	def := Picture{
		FPrintsWithSheet: true,
		FLocksWithSheet:  false,
		NoChangeAspect:   false,
		OffsetX:          0,
		OffsetY:          0,
		XScale:           1.0,
		YScale:           1.0,
	}

	var s []byte

	switch t := f.(type) {
	case string:
		s = []byte(t)
	case Picture:
		s, _ = json.Marshal(t)
	case *Picture:
		s, _ = json.Marshal(t)
	default:
		return &def, unknownFormat
	}

	c := &def
	err := json.Unmarshal(s, c)
	return c, err
}

// NewShape provides function to parse the format settings of the shape with default value.
func NewShape(f interface{}) (*Shape, error) {
	def := Shape{
		Width:  160,
		Height: 160,
		Format: Picture{
			FPrintsWithSheet: true,
			FLocksWithSheet:  false,
			NoChangeAspect:   false,
			OffsetX:          0,
			OffsetY:          0,
			XScale:           1.0,
			YScale:           1.0,
		},
	}

	var s []byte

	switch t := f.(type) {
	case string:
		s = []byte(t)
	case Shape:
		s, _ = json.Marshal(t)
	case *Shape:
		s, _ = json.Marshal(t)
	default:
		return &def, unknownFormat
	}

	c := &def
	err := json.Unmarshal(s, c)
	return c, err
}
