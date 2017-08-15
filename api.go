package excelize

import (
	"encoding/xml"
	"github.com/plandem/excelize/format"
)

func (f *File)GetSheet(sheet string)(*Worksheet){
	return &Worksheet{ f, f.workSheetReader(sheet) }
}

func (f *File)NewFormatStyle(fs *format.Style)(Style, error) {
	var cellXfsID, fontID, borderID, fillID int

	s := f.stylesReader()
	numFmtID := setNumFmt(s, fs)

	if fs.Font != (format.Font{}) {
		font, _ := xml.Marshal(setFont(fs))
		s.Fonts.Count++
		s.Fonts.Font = append(s.Fonts.Font, &xlsxFont{
			Font: string(font[6 : len(font)-7]),
		})
		fontID = s.Fonts.Count - 1
	}

	s.Borders.Count++
	s.Borders.Border = append(s.Borders.Border, setBorders(fs))
	borderID = s.Borders.Count - 1

	s.Fills.Count++
	s.Fills.Fill = append(s.Fills.Fill, setFills(fs, true))
	fillID = s.Fills.Count - 1

	applyAlignment, alignment := fs.Alignment != (format.Alignment{}), setAlignment(fs)
	cellXfsID = setCellXfs(s, fontID, numFmtID, fillID, borderID, applyAlignment, alignment)
	return Style(cellXfsID), nil
}