package excelize

import (
	"regexp"
	"strconv"
	"strings"
)

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

// GetCellAxis returns merged cell axis.
func (m *MergeCell) GetCellAxis() string {
	return (*m)[0]
}

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

// GetRangeCells returns axis in the X and Y direction included in the merged cell.
// example:
//    cellsX, cellxY, err := xlsx.GetRangeCells("D4:E10")
// return:
//    ["D4" "E4"] ["D4" "D5" "D6" "D7" "D8" "D9" "D10"] nil
//
// example:
//    cellsX, cellxY, err := xlsx.GetRangeCells("D4")
// return:
//    ["D4"] ["D4"] nil
//
func (f *File) GetRangeCells(cell string) ([]string, []string, error) {
	var cellsX, cellsY []string
	cellsX, err := getRangeCellsX(cell)
	if err != nil {
		return cellsX, cellsY, err
	}
	cellsY = getRangeCellsY(cell)
	return cellsX, cellsY, nil
}

// getRangeCellsX returns X direction included in the merged cell.
// example: getRangeCellsX("D4:E10")
// return: ["D4" "E4"] nil
func getRangeCellsX(cell string) ([]string, error) {
	startX, startY, endX, _ := getCellsRangeParams(cell)

	cells := []string{}
	startXNum, err := ColumnNameToNumber(startX)
	if err != nil {
		return cells, err
	}
	endXNum, err := ColumnNameToNumber(endX)
	if err != nil {
		return cells, err
	}
	for x := startXNum; x <= endXNum; x++ {
		col, err := ColumnNumberToName(x)
		if err != nil {
			return cells, err
		}
		cells = append(cells, getCellPath(col, startY))
	}

	return cells, nil
}

// getRangeCellsY returns Y direction included in the merged cell.
// example: getRangeCellsY("D4:E10")
// return: ["D4" "D5" "D6" "D7" "D8" "D9" "D10"]
func getRangeCellsY(cell string) []string {
	startX, startY, _, endY := getCellsRangeParams(cell)

	cells := []string{}
	cells = append(cells, getCellPath(startX, startY))
	for i := startY + 1; i <= endY; i++ {
		cells = append(cells, getCellPath(startX, i))
	}

	return cells
}

// getCellsRangeParams returns axis in the X and Y direction Start/End of the merged cell.
// example: getCellsRangeParams("D4:E10")
// return: "D" 4 "E" 10
func getCellsRangeParams(cell string) (string, int, string, int) {
	axis := strings.Split(cell, ":")
	axisStart := axis[0]
	axisEnd := axis[0]
	if len(axis) == 2 {
		axisEnd = axis[1]
	}
	startX, startY := getCellXY(axisStart)
	endX, endY := getCellXY(axisEnd)
	return startX, startY, endX, endY
}

func getCellXY(axis string) (string, int) {
	regexStr := regexp.MustCompile("[A-Z]{1,}")
	if !regexStr.MatchString(axis) {
		return "", 0
	}
	X := regexStr.FindString(axis)

	regexNum := regexp.MustCompile("\\d+")
	if !regexNum.MatchString(axis) {
		return "", 0
	}
	col := regexNum.FindString(axis)
	Y, _ := strconv.Atoi(col)

	return X, Y
}

func getCellPath(x string, y int) string {
	return x + strconv.Itoa(y)
}

// searchMergedCell returns merged cell axis.
// or single cell axis.
// example: xlsx.searchMergedCell(mergeCells, "D4")
// return: "D4:E10" or "D4"
func (f *File) searchMergedCell(mergeCells []MergeCell, cell string) string {
	for _, m := range mergeCells {
		axis := m.GetStartAxis()
		if axis == cell {
			return m.GetCellAxis()
		}
	}
	return cell
}
