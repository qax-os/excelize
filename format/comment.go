package format

// Comment directly maps the format settings of the comment.
type Comment struct {
	Author string `json:"author"`
	Text   string `json:"text"`
}
