package excelize

// formatButton directly maps the format settings of the comment.
type formatButton struct {
	Macro       string  `json:"macro"`
	Caption     string  `json:"caption"`
	Width       float64 `json:"width"`
	Height      float64 `json:"height"`
	OffsetX     int     `json:"x_offset"`
	OffsetY     int     `json:"y_offset"`
	ScaleX      float64 `json:"x_scale"`
	ScaleY      float64 `json:"y_scale"`
	Description string  `json:"description"`
}
