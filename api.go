package excelize

import (
	"strings"
	"strconv"
	"time"
	"fmt"
	"encoding/xml"
)

type Worksheet struct {
	f *File
	sheet *xlsxWorksheet
}

type Row struct {
	w *Worksheet
	row *xlsxRow
}

type Cell struct {
	r *Row
	cell *xlsxC
}

func (f *File)GetSheet(sheet string)(Worksheet){
	return Worksheet{ f, f.workSheetReader(sheet) }
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

func (w *Worksheet)getRow(rowIndex int, cells int )(Row) {
	rows := rowIndex + 1
	completeRow(w.sheet, rows, cells)

	return Row{w, w.sheet.SheetData.Row[rows]}
}

func (w *Worksheet)GetRow(rowIndex int)(Row) {
	return w.getRow(rowIndex, 0)
}

func (w *Worksheet)GetCell(axis string)(Cell) {
	xAxis, yAxis := w.axisToIndex(axis)

	r := w.getRow(xAxis + 1, yAxis + 1)

	completeCol(w.sheet, xAxis + 1, yAxis + 1)
	c := w.sheet.SheetData.Row[xAxis].C[yAxis]
	r.w.f.prepareCellStyle(r.w.sheet, yAxis, c.S)

	return Cell{&r, c}
}

//rows
func (r *Row)SetHeight(height float64) {
	r.row.Ht = height
	r.row.CustomHeight = true
}

func (r *Row)SetVisible(visible bool) {
	if visible {
		r.row.Hidden = false
		return
	}

	r.row.Hidden = true
}

func (r *Row) GetVisible() bool {
	return !r.row.Hidden
}

//cells
func (c *Cell) SetInt(value int) {
	c.cell.T = ""
	c.cell.V = strconv.Itoa(value)
}

func (c *Cell) SetStr(value string) {
	if len(value) > 32767 {
		value = value[0:32767]
	}

	// Leading space(s) character detection.
	if len(value) > 0 {
		if value[0] == 32 {
			c.cell.XMLSpace = xml.Attr{
				Name:  xml.Name{Space: NameSpaceXML, Local: "space"},
				Value: "preserve",
			}
		}
	}

	c.cell.T = "str"
	c.cell.V = value
}

func (c *Cell) SetDefault(value string) {
	c.cell.T = ""
	c.cell.V = value
}

func (c *Cell)SetValue(value interface{}) {
	switch t := value.(type) {
	case int:
		c.SetInt(value.(int))
	case int8:
		c.SetInt(int(value.(int8)))
	case int16:
		c.SetInt(int(value.(int16)))
	case int32:
		c.SetInt(int(value.(int32)))
	case int64:
		c.SetInt(int(value.(int64)))
	case float32:
		c.SetDefault(strconv.FormatFloat(float64(value.(float32)), 'f', -1, 32))
	case float64:
		c.SetDefault(strconv.FormatFloat(float64(value.(float64)), 'f', -1, 64))
	case string:
		c.SetStr(t)
	case []byte:
		c.SetStr(string(t))
	case time.Time:
		c.SetDefault(strconv.FormatFloat(float64(timeToExcelTime(timeToUTCTime(value.(time.Time)))), 'f', -1, 64))
		//c.setDefaultTimeStyle(sheet, axis)
	case nil:
		c.SetStr("")
	default:
		c.SetStr(fmt.Sprintf("%v", value))
	}
}
