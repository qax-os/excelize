package excelize

import (
	"fmt"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestAddSparkline(t *testing.T) {
	f := prepareSparklineDataset()

	// Set the columns widths to make the output clearer
	style, err := f.NewStyle(`{"font":{"bold":true}}`)
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "A1", "B1", style))
	assert.NoError(t, f.SetSheetViewOptions("Sheet1", 0, ZoomScale(150)))

	f.SetColWidth("Sheet1", "A", "A", 14)
	f.SetColWidth("Sheet1", "B", "B", 50)
	// Headings
	f.SetCellValue("Sheet1", "A1", "Sparkline")
	f.SetCellValue("Sheet1", "B1", "Description")

	f.SetCellValue("Sheet1", "B2", `A default "line" sparkline.`)
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A2"},
		Range:    []string{"Sheet3!A1:J1"},
	}))

	f.SetCellValue("Sheet1", "B3", `A default "column" sparkline.`)
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A3"},
		Range:    []string{"Sheet3!A2:J2"},
		Type:     "column",
	}))

	f.SetCellValue("Sheet1", "B4", `A default "win/loss" sparkline.`)
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A4"},
		Range:    []string{"Sheet3!A3:J3"},
		Type:     "win_loss",
	}))

	f.SetCellValue("Sheet1", "B6", "Line with markers.")
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A6"},
		Range:    []string{"Sheet3!A1:J1"},
		Markers:  true,
	}))

	f.SetCellValue("Sheet1", "B7", "Line with high and low points.")
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A7"},
		Range:    []string{"Sheet3!A1:J1"},
		High:     true,
		Low:      true,
	}))

	f.SetCellValue("Sheet1", "B8", "Line with first and last point markers.")
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A8"},
		Range:    []string{"Sheet3!A1:J1"},
		First:    true,
		Last:     true,
	}))

	f.SetCellValue("Sheet1", "B9", "Line with negative point markers.")
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A9"},
		Range:    []string{"Sheet3!A1:J1"},
		Negative: true,
	}))

	f.SetCellValue("Sheet1", "B10", "Line with axis.")
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A10"},
		Range:    []string{"Sheet3!A1:J1"},
		Axis:     true,
	}))

	f.SetCellValue("Sheet1", "B12", "Column with default style (1).")
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A12"},
		Range:    []string{"Sheet3!A2:J2"},
		Type:     "column",
	}))

	f.SetCellValue("Sheet1", "B13", "Column with style 2.")
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A13"},
		Range:    []string{"Sheet3!A2:J2"},
		Type:     "column",
		Style:    2,
	}))

	f.SetCellValue("Sheet1", "B14", "Column with style 3.")
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A14"},
		Range:    []string{"Sheet3!A2:J2"},
		Type:     "column",
		Style:    3,
	}))

	f.SetCellValue("Sheet1", "B15", "Column with style 4.")
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A15"},
		Range:    []string{"Sheet3!A2:J2"},
		Type:     "column",
		Style:    4,
	}))

	f.SetCellValue("Sheet1", "B16", "Column with style 5.")
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A16"},
		Range:    []string{"Sheet3!A2:J2"},
		Type:     "column",
		Style:    5,
	}))

	f.SetCellValue("Sheet1", "B17", "Column with style 6.")
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A17"},
		Range:    []string{"Sheet3!A2:J2"},
		Type:     "column",
		Style:    6,
	}))

	f.SetCellValue("Sheet1", "B18", "Column with a user defined color.")
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location:    []string{"A18"},
		Range:       []string{"Sheet3!A2:J2"},
		Type:        "column",
		SeriesColor: "#E965E0",
	}))

	f.SetCellValue("Sheet1", "B20", "A win/loss sparkline.")
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A20"},
		Range:    []string{"Sheet3!A3:J3"},
		Type:     "win_loss",
	}))

	f.SetCellValue("Sheet1", "B21", "A win/loss sparkline with negative points highlighted.")
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A21"},
		Range:    []string{"Sheet3!A3:J3"},
		Type:     "win_loss",
		Negative: true,
	}))

	f.SetCellValue("Sheet1", "B23", "A left to right column (the default).")
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A23"},
		Range:    []string{"Sheet3!A4:J4"},
		Type:     "column",
		Style:    20,
	}))

	f.SetCellValue("Sheet1", "B24", "A right to left column.")
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A24"},
		Range:    []string{"Sheet3!A4:J4"},
		Type:     "column",
		Style:    20,
		Reverse:  true,
	}))

	f.SetCellValue("Sheet1", "B25", "Sparkline and text in one cell.")
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A25"},
		Range:    []string{"Sheet3!A4:J4"},
		Type:     "column",
		Style:    20,
	}))
	f.SetCellValue("Sheet1", "A25", "Growth")

	f.SetCellValue("Sheet1", "B27", "A grouped sparkline. Changes are applied to all three.")
	assert.NoError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A27", "A28", "A29"},
		Range:    []string{"Sheet3!A5:J5", "Sheet3!A6:J6", "Sheet3!A7:J7"},
		Markers:  true,
	}))

	// Sheet2 sections
	assert.NoError(t, f.AddSparkline("Sheet2", &SparklineOption{
		Location: []string{"F3"},
		Range:    []string{"Sheet2!A3:E3"},
		Type:     "win_loss",
		Negative: true,
	}))

	assert.NoError(t, f.AddSparkline("Sheet2", &SparklineOption{
		Location: []string{"F1"},
		Range:    []string{"Sheet2!A1:E1"},
		Markers:  true,
	}))

	assert.NoError(t, f.AddSparkline("Sheet2", &SparklineOption{
		Location: []string{"F2"},
		Range:    []string{"Sheet2!A2:E2"},
		Type:     "column",
		Style:    12,
	}))

	assert.NoError(t, f.AddSparkline("Sheet2", &SparklineOption{
		Location: []string{"F3"},
		Range:    []string{"Sheet2!A3:E3"},
		Type:     "win_loss",
		Negative: true,
	}))

	// Save xlsx file by the given path.
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddSparkline.xlsx")))

	// Test error exceptions
	assert.EqualError(t, f.AddSparkline("SheetN", &SparklineOption{
		Location: []string{"F3"},
		Range:    []string{"Sheet2!A3:E3"},
	}), "sheet SheetN is not exist")

	assert.EqualError(t, f.AddSparkline("Sheet1", nil), "parameter is required")

	assert.EqualError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Range: []string{"Sheet2!A3:E3"},
	}), `parameter 'Location' is required`)

	assert.EqualError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"F3"},
	}), `parameter 'Range' is required`)

	assert.EqualError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"F2", "F3"},
		Range:    []string{"Sheet2!A3:E3"},
	}), `must have the same number of 'Location' and 'Range' parameters`)

	assert.EqualError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"F3"},
		Range:    []string{"Sheet2!A3:E3"},
		Type:     "unknown_type",
	}), `parameter 'Type' must be 'line', 'column' or 'win_loss'`)

	assert.EqualError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"F3"},
		Range:    []string{"Sheet2!A3:E3"},
		Style:    -1,
	}), `parameter 'Style' must betweent 0-35`)

	assert.EqualError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"F3"},
		Range:    []string{"Sheet2!A3:E3"},
		Style:    -1,
	}), `parameter 'Style' must betweent 0-35`)

	f.Sheet["xl/worksheets/sheet1.xml"].ExtLst.Ext = `<extLst>
	    <ext x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
	        <x14:sparklineGroups
	            xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
	             <x14:sparklineGroup>
                    </x14:sparklines>
                </x14:sparklineGroup>
	        </x14:sparklineGroups>
	    </ext>
	</extLst>`
	assert.EqualError(t, f.AddSparkline("Sheet1", &SparklineOption{
		Location: []string{"A2"},
		Range:    []string{"Sheet3!A1:J1"},
	}), "XML syntax error on line 6: element <sparklineGroup> closed by </sparklines>")
}

func prepareSparklineDataset() *File {
	f := NewFile()
	sheet2 := [][]int{
		{-2, 2, 3, -1, 0},
		{30, 20, 33, 20, 15},
		{1, -1, -1, 1, -1},
	}
	sheet3 := [][]int{
		{-2, 2, 3, -1, 0, -2, 3, 2, 1, 0},
		{30, 20, 33, 20, 15, 5, 5, 15, 10, 15},
		{1, 1, -1, -1, 1, -1, 1, 1, 1, -1},
		{5, 6, 7, 10, 15, 20, 30, 50, 70, 100},
		{-2, 2, 3, -1, 0, -2, 3, 2, 1, 0},
		{3, -1, 0, -2, 3, 2, 1, 0, 2, 1},
		{0, -2, 3, 2, 1, 0, 1, 2, 3, 1},
	}
	f.NewSheet("Sheet2")
	f.NewSheet("Sheet3")
	for row, data := range sheet2 {
		f.SetSheetRow("Sheet2", fmt.Sprintf("A%d", row+1), &data)
	}
	for row, data := range sheet3 {
		f.SetSheetRow("Sheet3", fmt.Sprintf("A%d", row+1), &data)
	}
	return f
}
