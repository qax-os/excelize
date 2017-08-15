package format

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
