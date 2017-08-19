package format

import (
	"encoding/json"
)

// Comment directly maps the format settings of the comment.
type Comment struct {
	Author string `json:"author"`
	Text   string `json:"text"`
}

// NewComment provides function to parse the format settings of the comment with default value.
func NewComment(f interface{})(*Comment, error) {
	def := Comment{
		Author: " ",
		Text: " ",
	}

	var s []byte

	switch t := f.(type) {
	case string:
		s = []byte(t)
	case Comment:
		s, _ = json.Marshal(t)
	case *Comment:
		s, _ = json.Marshal(t)
	default:
		return &def, unknownFormat
	}

	c := &def
	err := json.Unmarshal(s, c)
	return c, err
}
