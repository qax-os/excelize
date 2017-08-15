package excelize

import (
	"strings"
	"strconv"
)

type Worksheet struct {
	f *File
	sheet *xlsxWorksheet
}

//worksheets
func (w *Worksheet)mergeCellsParser(axis string) string {
	axis = strings.ToUpper(axis)
	if w.sheet.MergeCells != nil {
		for i := 0; i < len(w.sheet.MergeCells.Cells); i++ {
			if checkCellInArea(axis, w.sheet.MergeCells.Cells[i].Ref) {
				axis = strings.Split(w.sheet.MergeCells.Cells[i].Ref, ":")[0]
			}
		}
	}

	return axis
}

func (w *Worksheet)axisToIndex(axis string)(int, int) {
	axis = w.mergeCellsParser(axis)
	col := string(strings.Map(letterOnlyMapF, axis))
	row, _ := strconv.Atoi(strings.Map(intOnlyMapF, axis))
	xAxis := row - 1
	yAxis := TitleToNumber(col)

	return xAxis, yAxis
}

func (w *Worksheet)GetRange(fromAxis, toAxis string)(*Range) {
	fromAxis = strings.ToUpper(fromAxis)
	toAxis = strings.ToUpper(toAxis)

	fromXAxis, fromYAxis := w.axisToIndex(fromAxis)
	toXAxis, toYAxis := w.axisToIndex(toAxis)

	if toXAxis < fromXAxis {
		fromAxis, toAxis = toAxis, fromAxis
		toXAxis, fromXAxis = fromXAxis, toXAxis
	}

	if toYAxis < fromYAxis {
		fromAxis, toAxis = toAxis, fromAxis
		toYAxis, fromYAxis = fromYAxis, toYAxis
	}

	completeRow(w.sheet, toYAxis + 1, toXAxis + 1)
	completeCol(w.sheet, toYAxis + 1, toXAxis + 1)

	// Correct the coordinate area, such correct C1:B3 to B1:C3.
	fromAxis = ToAlphaString(fromXAxis) + strconv.Itoa(fromYAxis + 1)
	toAxis = ToAlphaString(toXAxis) + strconv.Itoa(toYAxis + 1)

	return &Range{
		w,
		fromAxis,
		toAxis,
	}
}

func (w *Worksheet)getRow(rowIndex int, cells int )(*Row) {
	rows := rowIndex + 1
	completeRow(w.sheet, rows, cells)

	return &Row{w, w.sheet.SheetData.Row[rows]}
}

func (w *Worksheet)GetRow(rowIndex int)(*Row) {
	return w.getRow(rowIndex, 0)
}

func (w *Worksheet)GetCell(axis string)(*Cell) {
	xAxis, yAxis := w.axisToIndex(axis)

	r := w.getRow(xAxis, yAxis + 1)

	completeCol(w.sheet, xAxis + 1, yAxis + 1)
	c := w.sheet.SheetData.Row[xAxis].C[yAxis]
	r.w.f.prepareCellStyle(r.w.sheet, yAxis, c.S)

	return &Cell{r, c}
}

