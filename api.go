package excelize

func (f *File)GetSheet(sheet string)(*Worksheet){
	return &Worksheet{ f, f.workSheetReader(sheet) }
}
