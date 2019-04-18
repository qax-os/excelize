package excelize

import "strings"

// GetMergeCells provides a function to get all merged cells from a worksheet currently.
func (f *File) GetMergeCells(sheet string) ([]MergeCell, error) {
	var mergeCells []MergeCell
	xlsx, err := f.workSheetReader(sheet)
	if err != nil {
		return mergeCells, err
	}
	if xlsx.MergeCells != nil {
		mergeCells = make([]MergeCell, 0, len(xlsx.MergeCells.Cells))

		for i := range xlsx.MergeCells.Cells {
			ref := xlsx.MergeCells.Cells[i].Ref
			axis := strings.Split(ref, ":")[0]
			val, _ := f.GetCellValue(sheet, axis)
			mergeCells = append(mergeCells, []string{ref, val})
		}
	}

	return mergeCells, err
}

// MergeCell define a merged cell data.
// It consists of the following structure.
// example: []string{"D4:E10", "cell value"}
type MergeCell []string

// GetCellValue returns merged cell value.
func (m *MergeCell) GetCellValue() string {
	return (*m)[1]
}

// GetStartAxis returns the merge start axis.
// example: "C2"
func (m *MergeCell) GetStartAxis() string {
	axis := strings.Split((*m)[0], ":")
	return axis[0]
}

// GetEndAxis returns the merge end axis.
// example: "D4"
func (m *MergeCell) GetEndAxis() string {
	axis := strings.Split((*m)[0], ":")
	return axis[1]
}