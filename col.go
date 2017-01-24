package excelize

import (
	"encoding/xml"
	"math"
	"strconv"
	"strings"
)

// Define the default cell size and EMU unit of measurement.
const (
	defaultColWidthPixels  int = 64
	defaultRowHeightPixels int = 20
	EMU                    int = 9525
)

// SetColWidth provides function to set the width of a single column or multiple
// columns. For example:
//
//    xlsx := excelize.CreateFile()
//    xlsx.SetColWidth("Sheet1", "A", "H", 20)
//    err := xlsx.Save()
//    if err != nil {
//        fmt.Println(err)
//        os.Exit(1)
//    }
//
func (f *File) SetColWidth(sheet, startcol, endcol string, width float64) {
	min := titleToNumber(strings.ToUpper(startcol)) + 1
	max := titleToNumber(strings.ToUpper(endcol)) + 1
	if min > max {
		min, max = max, min
	}
	var xlsx xlsxWorksheet
	name := "xl/worksheets/" + strings.ToLower(sheet) + ".xml"
	xml.Unmarshal([]byte(f.readXML(name)), &xlsx)
	col := xlsxCol{
		Min:         min,
		Max:         max,
		Width:       width,
		CustomWidth: true,
	}
	if xlsx.Cols != nil {
		xlsx.Cols.Col = append(xlsx.Cols.Col, col)
	} else {
		cols := xlsxCols{}
		cols.Col = append(cols.Col, col)
		xlsx.Cols = &cols
	}
	output, _ := xml.Marshal(xlsx)
	f.saveFileList(name, replaceWorkSheetsRelationshipsNameSpace(string(output)))
}

// positionObjectPixels calculate the vertices that define the position of a
// graphical object within the worksheet in pixels.
//
//          +------------+------------+
//          |     A      |      B     |
//    +-----+------------+------------+
//    |     |(x1,y1)     |            |
//    |  1  |(A1)._______|______      |
//    |     |    |              |     |
//    |     |    |              |     |
//    +-----+----|    OBJECT    |-----+
//    |     |    |              |     |
//    |  2  |    |______________.     |
//    |     |            |        (B2)|
//    |     |            |     (x2,y2)|
//    +-----+------------+------------+
//
// Example of an object that covers some of the area from cell A1 to B2.
//
// Based on the width and height of the object we need to calculate 8 vars:
//
//    colStart, rowStart, colEnd, rowEnd, x1, y1, x2, y2.
//
// We also calculate the absolute x and y position of the top left vertex of
// the object. This is required for images.
//
// The width and height of the cells that the object occupies can be
// variable and have to be taken into account.
//
// The values of col_start and row_start are passed in from the calling
// function. The values of col_end and row_end are calculated by
// subtracting the width and height of the object from the width and
// height of the underlying cells.
//
//    colStart        # Col containing upper left corner of object.
//    x1              # Distance to left side of object.
//
//    rowStart        # Row containing top left corner of object.
//    y1              # Distance to top of object.
//
//    colEnd          # Col containing lower right corner of object.
//    x2              # Distance to right side of object.
//
//    rowEnd          # Row containing bottom right corner of object.
//    y2              # Distance to bottom of object.
//
//    width           # Width of object frame.
//    height          # Height of object frame.
//
//    xAbs            # Absolute distance to left side of object.
//    yAbs            # Absolute distance to top side of object.
//
func (f *File) positionObjectPixels(sheet string, colStart, rowStart, x1, y1, width, height int) (int, int, int, int, int, int, int, int) {
	xAbs := 0
	yAbs := 0

	// Calculate the absolute x offset of the top-left vertex.
	for colID := 1; colID <= colStart; colID++ {
		xAbs += f.getColWidth(sheet, colID)
	}
	xAbs += x1

	// Calculate the absolute y offset of the top-left vertex.
	// Store the column change to allow optimisations.
	for rowID := 1; rowID <= rowStart; rowID++ {
		yAbs += f.getRowHeight(sheet, rowID)
	}
	yAbs += y1

	// Adjust start column for offsets that are greater than the col width.
	for x1 >= f.getColWidth(sheet, colStart) {
		x1 -= f.getColWidth(sheet, colStart)
		colStart++
	}

	// Adjust start row for offsets that are greater than the row height.
	for y1 >= f.getRowHeight(sheet, rowStart) {
		y1 -= f.getRowHeight(sheet, rowStart)
		rowStart++
	}

	// Initialise end cell to the same as the start cell.
	colEnd := colStart
	rowEnd := rowStart

	width += x1
	height += y1

	// Subtract the underlying cell widths to find end cell of the object.
	for width >= f.getColWidth(sheet, colEnd) {
		colEnd++
		width -= f.getColWidth(sheet, colEnd)
	}

	// Subtract the underlying cell heights to find end cell of the object.
	for height >= f.getRowHeight(sheet, rowEnd) {
		rowEnd++
		height -= f.getRowHeight(sheet, rowEnd)
	}

	// The end vertices are whatever is left from the width and height.
	x2 := width
	y2 := height
	return colStart, rowStart, xAbs, yAbs, colEnd, rowEnd, x2, y2
}

// getColWidth provides function to get column width in pixels by given sheet
// name and column index.
func (f *File) getColWidth(sheet string, col int) int {
	var xlsx xlsxWorksheet
	name := "xl/worksheets/" + strings.ToLower(sheet) + ".xml"
	xml.Unmarshal([]byte(f.readXML(name)), &xlsx)
	if xlsx.Cols != nil {
		var width float64
		for _, v := range xlsx.Cols.Col {
			if v.Min <= col && col <= v.Max {
				width = v.Width
			}
		}
		if width != 0 {
			return int(convertColWidthToPixels(width))
		}
	}
	// Optimisation for when the column widths haven't changed.
	return defaultColWidthPixels
}

// getRowHeight provides function to get row height in pixels by given sheet
// name and row index.
func (f *File) getRowHeight(sheet string, row int) int {
	var xlsx xlsxWorksheet
	name := "xl/worksheets/" + strings.ToLower(sheet) + ".xml"
	xml.Unmarshal([]byte(f.readXML(name)), &xlsx)
	for _, v := range xlsx.SheetData.Row {
		if v.R == row && v.Ht != "" {
			ht, _ := strconv.ParseFloat(v.Ht, 64)
			return int(convertRowHeightToPixels(ht))
		}
	}
	// Optimisation for when the row heights haven't changed.
	return defaultRowHeightPixels
}

// convertColWidthToPixels provieds function to convert the width of a cell from
// user's units to pixels. Excel rounds the column width to the nearest pixel.
// If the width hasn't been set by the user we use the default value. If the
// column is hidden it has a value of zero.
func convertColWidthToPixels(width float64) float64 {
	var padding float64 = 5
	var pixels float64
	var maxDigitWidth float64 = 7
	if width == 0 {
		return pixels
	}
	if width < 1 {
		pixels = (width * 12) + 0.5
		return math.Ceil(pixels)
	}
	pixels = (width*maxDigitWidth + 0.5) + padding
	return math.Ceil(pixels)
}

// convertRowHeightToPixels provides function to convert the height of a cell
// from user's units to pixels. If the height hasn't been set by the user we use
// the default value. If the row is hidden it has a value of zero.
func convertRowHeightToPixels(height float64) float64 {
	var pixels float64
	if height == 0 {
		return pixels
	}
	pixels = math.Ceil(4.0 / 3.0 * height)
	return pixels
}
