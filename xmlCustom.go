package excelize

import (
	"encoding/xml"
	"time"
)

// xlsxCustomProperties is the root element of the Custom File Properties part
type xlsxCustomProperties struct {
	XMLName xml.Name             `xml:"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties Properties"`
	Vt      string               `xml:"xmlns:vt,attr"`
	Props   []xlsxCustomProperty `xml:"property"`
}

type xlsxCustomProperty struct {
	FmtID    string         `xml:"fmtid,attr"`
	PID      string         `xml:"pid,attr"`
	Name     string         `xml:"name,attr"`
	Text     *TextValue     `xml:"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes lpwstr,omitempty"`
	Bool     *BoolValue     `xml:"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes bool,omitempty"`
	Number   *NumberValue   `xml:"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes r8,omitempty"`
	DateTime *FileTimeValue `xml:"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes filetime,omitempty"`
}

func (p *xlsxCustomProperty) getPropertyValue() interface{} {
	if p.Text != nil {
		return p.Text.Text
	}
	if p.Bool != nil {
		return p.Bool.Bool
	}
	if p.Number != nil {
		return p.Number.Number
	}
	if p.DateTime != nil {
		// parse date time raw str to time.Time
		timeStr := p.DateTime.DateTime
		parsedTime, err := time.ParseInLocation(time.RFC3339, timeStr, time.Local)
		if err != nil {
			return nil
		}
		return parsedTime
	}

	return nil
}

type TextValue struct {
	Text string `xml:",chardata"`
}

type BoolValue struct {
	Bool bool `xml:",chardata"`
}

type NumberValue struct {
	Number float64 `xml:",chardata"`
}

type FileTimeValue struct {
	DateTime string `xml:",chardata"`
}
