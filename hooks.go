package excelize

import "encoding/xml"

type NumFmtIdToCodeHookInfo struct {
	XmlAttr map[string][]xml.Attr
	NumFmt  []*xlsxNumFmt
}

// NumFmtIdToCodeHook User-defined mapping function for numFmtId to numFmtCode
// if return ok is false, will use default mapping
type NumFmtIdToCodeHook func(info NumFmtIdToCodeHookInfo, numFmtId int) (numFmtCode string, ok bool)

func runNumFmtIdToCodeHook(file *File, numFmtId int) (numFmtCode string, ok bool) {
	if file == nil || file.options == nil || file.options.NumFmtIdToCodeHook == nil {
		return
	}
	hookInfo := NumFmtIdToCodeHookInfo{XmlAttr: file.xmlAttr}
	if file.Styles != nil && file.Styles.NumFmts != nil {
		hookInfo.NumFmt = file.Styles.NumFmts.NumFmt
	}
	return file.options.NumFmtIdToCodeHook(hookInfo, numFmtId)
}
