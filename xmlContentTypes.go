// Some code of this file reference tealeg/xlsx

package excelize

import (
	"encoding/xml"
)

type xlsxTypes struct {
	XMLName   xml.Name       `xml:"http://schemas.openxmlformats.org/package/2006/content-types Types"`
	Overrides []xlsxOverride `xml:"Override"`
	Defaults  []xlsxDefault  `xml:"Default"`
}

type xlsxOverride struct {
	PartName    string `xml:",attr"`
	ContentType string `xml:",attr"`
}

type xlsxDefault struct {
	Extension   string `xml:",attr"`
	ContentType string `xml:",attr"`
}
