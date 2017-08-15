package excelize

import (
	"strings"
	"strconv"
	"time"
	"fmt"
	"encoding/xml"
)

type Style int

type Worksheet struct {
	f *File
	sheet *xlsxWorksheet
}

type Row struct {
	w *Worksheet
	row *xlsxRow
}

type Col struct {
	w *Worksheet
	col *xlsxCol
}

type Cell struct {
	//c *Col
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

	r := w.getRow(xAxis, yAxis + 1)

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

func (c *Cell) SetString(value string) {
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

func (c *Cell) SetTime(value time.Time) {
	c.cell.T = ""
	c.cell.V = strconv.FormatFloat(float64(timeToExcelTime(timeToUTCTime(value))), 'f', -1, 64)
	if c.cell.S == 0 {
		style, _ := c.r.w.f.NewStyle(`{"number_format": 22}`)
		c.cell.S = style
	}
}

func (c *Cell) SetFloat(value float64) {
	c.cell.T = ""
	c.cell.V = strconv.FormatFloat(value, 'f', -1, 64)
}

func (c *Cell) SetFloatWithFormat(value float64, format byte) {
	c.cell.T = ""
	c.cell.V = strconv.FormatFloat(value, format, -1, 64)
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
		c.SetFloat(float64(value.(float32)))
	case float64:
		c.SetFloat(float64(value.(float64)))
	case string:
		c.SetString(t)
	case []byte:
		c.SetString(string(t))
	case time.Time:
		c.SetTime(value.(time.Time))
	case nil:
		c.SetString("")
	default:
		c.SetString(fmt.Sprintf("%v", value))
	}
}

func (c *Cell) GetStyle() Style {
	return Style(c.cell.S)
}

func (c *Cell) SetStyle(s Style) {
	c.cell.S = int(s)
}

func (c *Cell) SetFormula(formula string) {
	if c.cell.F != nil {
		c.cell.F.Content = formula
	} else {
		c.cell.F = &xlsxF{
			Content: formula,
		}
	}
}
