package excelize

import (
	"encoding/xml"
)

// xlsxCustomProperties is the root element of the Custom File Properties part
type xlsxCustomProperties struct {
	XMLName xml.Name             `xml:"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties Properties"`
	Vt      string               `xml:"xmlns:vt,attr"`
	Props   []xlsxCustomProperty `xml:"property"`
}

type xlsxCustomProperty struct {
	FmtID string `xml:"fmtid,attr"`
	PID   string `xml:"pid,attr"`
	Name  string `xml:"name,attr"`
	V     string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes lpwstr"`
}
