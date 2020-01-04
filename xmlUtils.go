package excelize

import (
	"encoding/xml"
	"regexp"
	"strings"
	"sync"
)

type relationship string

func (r relationship) MarshalXMLAttr(name xml.Name) (xml.Attr, error) {
	return xml.Attr{Name: xml.Name{
		Local: "r:id",
	}, Value: string(r)}, nil
}

var (
	defaultAttrs     []xml.Attr
	defaultAttrsInit sync.Once
)

func getDefaultAttrs() []xml.Attr {
	defaultAttrsInit.Do(func() {
		re := regexp.MustCompile("[A-z0-9:]+=\".+?\"")
		for _, split := range re.FindAll([]byte(templateNamespaceIDMap), -1) {
			k := strings.Split(string(split), "=")[0]
			v := strings.Split(string(split), "=")[1]
			v = v[1 : len(v)-1]
			defaultAttrs = append(defaultAttrs, xml.Attr{
				Name:  xml.Name{Local: k, Space: ""},
				Value: v,
			})
		}
	})
	return defaultAttrs
}
