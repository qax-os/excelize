package excelize

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestChartSize(t *testing.T) {
	f := NewFile()
	sheet1 := f.GetSheetName(0)

	categories := map[string]string{
		"A2": "Small",
		"A3": "Normal",
		"A4": "Large",
		"B1": "Apple",
		"C1": "Orange",
		"D1": "Pear",
	}
	for cell, v := range categories {
		assert.NoError(t, f.SetCellValue(sheet1, cell, v))
	}

	values := map[string]int{
		"B2": 2,
		"C2": 3,
		"D2": 3,
		"B3": 5,
		"C3": 2,
		"D3": 4,
		"B4": 6,
		"C4": 7,
		"D4": 8,
	}
	for cell, v := range values {
		assert.NoError(t, f.SetCellValue(sheet1, cell, v))
	}

	assert.NoError(t, f.AddChart("Sheet1", "E4", &Chart{
		Type: Col3DClustered,
		Dimension: ChartDimension{
			Width:  640,
			Height: 480,
		},
		Series: []ChartSeries{
			{Name: "Sheet1!$A$2", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$2:$D$2"},
			{Name: "Sheet1!$A$3", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$3:$D$3"},
			{Name: "Sheet1!$A$4", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$4:$D$4"},
		},
		Title: []RichTextRun{{Text: "3D Clustered Column Chart"}},
	}))

	var buffer bytes.Buffer

	// Save spreadsheet by the given path.
	assert.NoError(t, f.Write(&buffer))

	newFile, err := OpenReader(&buffer)
	assert.NoError(t, err)

	chartsNum := newFile.countCharts()
	if !assert.Equal(t, 1, chartsNum, "Expected 1 chart, actual %d", chartsNum) {
		t.FailNow()
	}

	var (
		workdir decodeWsDr
		anchor  decodeCellAnchor
	)

	content, ok := newFile.Pkg.Load("xl/drawings/drawing1.xml")
	assert.True(t, ok, "Can't open the chart")

	err = xml.Unmarshal(content.([]byte), &workdir)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	err = xml.Unmarshal([]byte("<decodeCellAnchor>"+
		workdir.TwoCellAnchor[0].Content+"</decodeCellAnchor>"), &anchor)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	if !assert.Equal(t, 4, anchor.From.Col, "Expected 'from' column 4") ||
		!assert.Equal(t, 3, anchor.From.Row, "Expected 'from' row 3") {

		t.FailNow()
	}

	if !assert.Equal(t, 14, anchor.To.Col, "Expected 'to' column 14") ||
		!assert.Equal(t, 29, anchor.To.Row, "Expected 'to' row 29") {

		t.FailNow()
	}
}

func TestAddDrawingChart(t *testing.T) {
	f := NewFile()
	assert.EqualError(t, f.addDrawingChart("SheetN", "", "", 0, 0, 0, nil), newCellNameToCoordinatesError("", newInvalidCellNameError("")).Error())

	path := "xl/drawings/drawing1.xml"
	f.Pkg.Store(path, MacintoshCyrillicCharset)
	assert.EqualError(t, f.addDrawingChart("Sheet1", path, "A1", 0, 0, 0, &GraphicOptions{PrintObject: boolPtr(true), Locked: boolPtr(false)}), "XML syntax error on line 1: invalid UTF-8")
}

func TestAddSheetDrawingChart(t *testing.T) {
	f := NewFile()
	path := "xl/drawings/drawing1.xml"
	f.Pkg.Store(path, MacintoshCyrillicCharset)
	assert.EqualError(t, f.addSheetDrawingChart(path, 0, &GraphicOptions{PrintObject: boolPtr(true), Locked: boolPtr(false)}), "XML syntax error on line 1: invalid UTF-8")
}

func TestDeleteDrawing(t *testing.T) {
	f := NewFile()
	path := "xl/drawings/drawing1.xml"
	f.Pkg.Store(path, MacintoshCyrillicCharset)
	_, err := f.deleteDrawing(0, 0, path, "Chart")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	f, err = OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)
	f.Drawings.Store(path, &xlsxWsDr{OneCellAnchor: []*xdrCellAnchor{{
		GraphicFrame: string(MacintoshCyrillicCharset),
	}}})
	_, err = f.deleteDrawing(0, 0, path, "Chart")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	f.Drawings.Store(path, &xlsxWsDr{TwoCellAnchor: []*xdrCellAnchor{{
		GraphicFrame: string(MacintoshCyrillicCharset),
	}}})
	_, err = f.deleteDrawing(0, 0, path, "Chart")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
}

func TestAddChart(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	categories := map[string]string{"A30": "SS", "A31": "S", "A32": "M", "A33": "L", "A34": "LL", "A35": "XL", "A36": "XXL", "A37": "XXXL", "B29": "Apple", "C29": "Orange", "D29": "Pear"}
	values := map[string]int{"B30": 1, "C30": 1, "D30": 1, "B31": 2, "C31": 2, "D31": 2, "B32": 3, "C32": 3, "D32": 3, "B33": 4, "C33": 4, "D33": 4, "B34": 5, "C34": 5, "D34": 5, "B35": 6, "C35": 6, "D35": 6, "B36": 7, "C36": 7, "D36": 7, "B37": 8, "C37": 8, "D37": 8}
	for k, v := range categories {
		assert.NoError(t, f.SetCellValue("Sheet1", k, v))
	}
	for k, v := range values {
		assert.NoError(t, f.SetCellValue("Sheet1", k, v))
	}
	assert.EqualError(t, f.AddChart("Sheet1", "P1", nil), ErrParameterInvalid.Error())

	// Test add chart on not exists worksheet
	assert.EqualError(t, f.AddChart("SheetN", "P1", nil), "sheet SheetN does not exist")
	maximum, minimum, zero := 7.5, 0.5, .0
	series := []ChartSeries{
		{Name: "Sheet1!$A$30", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$30:$D$30"},
		{Name: "Sheet1!$A$31", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$31:$D$31"},
		{Name: "Sheet1!$A$32", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$32:$D$32"},
		{Name: "Sheet1!$A$33", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$33:$D$33"},
		{Name: "Sheet1!$A$34", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$34:$D$34"},
		{Name: "Sheet1!$A$35", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$35:$D$35"},
		{Name: "Sheet1!$A$36", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$36:$D$36"},
		{
			Name: "Sheet1!$A$37", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$37:$D$37",
			Marker: ChartMarker{
				Fill: Fill{Type: "pattern", Color: []string{"FFFF00"}, Pattern: 1},
			},
		},
	}
	series2 := []ChartSeries{
		{
			Name: "Sheet1!$A$30", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$30:$D$30",
			Fill:   Fill{Type: "pattern", Color: []string{"000000"}, Pattern: 1},
			Marker: ChartMarker{Symbol: "none", Size: 10},
		},
		{Name: "Sheet1!$A$31", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$31:$D$31"},
		{Name: "Sheet1!$A$32", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$32:$D$32"},
		{Name: "Sheet1!$A$33", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$33:$D$33"},
		{Name: "Sheet1!$A$34", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$34:$D$34"},
		{Name: "Sheet1!$A$35", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$35:$D$35"},
		{Name: "Sheet1!$A$36", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$36:$D$36"},
		{Name: "Sheet1!$A$37", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$37:$D$37", Line: ChartLine{Width: 0.25}},
	}
	series3 := []ChartSeries{{Name: "Sheet1!$A$30", Categories: "Sheet1!$A$30:$D$37", Values: "Sheet1!$B$30:$B$37"}}
	series4 := []ChartSeries{
		{Name: "Sheet1!$A$30", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$30:$D$30", Sizes: "Sheet1!$B$30:$D$30", DataLabelPosition: ChartDataLabelsPositionAbove},
		{Name: "Sheet1!$A$31", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$31:$D$31", Sizes: "Sheet1!$B$31:$D$31", DataLabelPosition: ChartDataLabelsPositionLeft},
		{Name: "Sheet1!$A$32", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$32:$D$32", Sizes: "Sheet1!$B$32:$D$32", DataLabelPosition: ChartDataLabelsPositionBestFit},
		{Name: "Sheet1!$A$33", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$33:$D$33", Sizes: "Sheet1!$B$33:$D$33", DataLabelPosition: ChartDataLabelsPositionCenter},
		{Name: "Sheet1!$A$34", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$34:$D$34", Sizes: "Sheet1!$B$34:$D$34", DataLabelPosition: ChartDataLabelsPositionInsideBase},
		{Name: "Sheet1!$A$35", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$35:$D$35", Sizes: "Sheet1!$B$35:$D$35", DataLabelPosition: ChartDataLabelsPositionInsideEnd},
		{Name: "Sheet1!$A$36", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$36:$D$36", Sizes: "Sheet1!$B$36:$D$36", DataLabelPosition: ChartDataLabelsPositionOutsideEnd},
		{Name: "Sheet1!$A$37", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$37:$D$37", Sizes: "Sheet1!$B$37:$D$37", DataLabelPosition: ChartDataLabelsPositionRight},
	}
	format := GraphicOptions{
		ScaleX:          defaultDrawingScale,
		ScaleY:          defaultDrawingScale,
		OffsetX:         15,
		OffsetY:         10,
		PrintObject:     boolPtr(true),
		LockAspectRatio: false,
		Locked:          boolPtr(false),
	}
	legend := ChartLegend{Position: "left", ShowLegendKey: false}
	plotArea := ChartPlotArea{
		SecondPlotValues: 3,
		ShowBubbleSize:   true,
		ShowCatName:      true,
		ShowLeaderLines:  false,
		ShowPercent:      true,
		ShowSerName:      true,
		ShowVal:          true,
		Fill:             Fill{Type: "pattern", Pattern: 1},
	}
	for _, c := range []struct {
		sheetName, cell string
		opts            *Chart
	}{
		{sheetName: "Sheet1", cell: "P1", opts: &Chart{Type: Col, Series: series, Format: format, Legend: ChartLegend{Position: "none", ShowLegendKey: true}, Title: []RichTextRun{{Text: "2D Column Chart"}}, PlotArea: plotArea, Border: ChartLine{Type: ChartLineNone}, ShowBlanksAs: "zero", XAxis: ChartAxis{Font: Font{Bold: true, Italic: true, Underline: "dbl", Family: "Times New Roman", Size: 15, Strike: true, Color: "000000"}, Title: []RichTextRun{{Text: "Primary Horizontal Axis Title"}}}, YAxis: ChartAxis{Font: Font{Bold: false, Italic: false, Underline: "sng", Color: "777777"}, Title: []RichTextRun{{Text: "Primary Vertical Axis Title", Font: &Font{Color: "777777", Bold: true, Italic: true, Size: 12}}}}}},
		{sheetName: "Sheet1", cell: "X1", opts: &Chart{Type: ColStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "2D Stacked Column Chart"}}, PlotArea: plotArea, Fill: Fill{Type: "pattern", Pattern: 1}, Border: ChartLine{Type: ChartLineAutomatic}, ShowBlanksAs: "zero", GapWidth: uintPtr(10), Overlap: intPtr(100)}},
		{sheetName: "Sheet1", cell: "P16", opts: &Chart{Type: ColPercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "100% Stacked Column Chart"}}, PlotArea: plotArea, Fill: Fill{Type: "pattern", Color: []string{"EEEEEE"}, Pattern: 1}, Border: ChartLine{Type: ChartLineSolid, Width: 2}, ShowBlanksAs: "zero", XAxis: ChartAxis{Alignment: Alignment{Vertical: "wordArtVertRtl", TextRotation: 0}}}},
		{sheetName: "Sheet1", cell: "X16", opts: &Chart{Type: Col3DClustered, Series: series, Format: format, Legend: ChartLegend{Position: "bottom", ShowLegendKey: false}, Title: []RichTextRun{{Text: "3D Clustered Column Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "P30", opts: &Chart{Type: Col3DStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Stacked Column Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", XAxis: ChartAxis{Alignment: Alignment{Vertical: "vert", TextRotation: 0}}}},
		{sheetName: "Sheet1", cell: "X30", opts: &Chart{Type: Col3DPercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D 100% Stacked Column Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "X45", opts: &Chart{Type: Radar, Series: series, Format: format, Legend: ChartLegend{Position: "top_right", ShowLegendKey: false}, Title: []RichTextRun{{Text: "Radar Chart"}}, PlotArea: plotArea, ShowBlanksAs: "span"}},
		{sheetName: "Sheet1", cell: "AF1", opts: &Chart{Type: Col3DConeStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Cone Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AF16", opts: &Chart{Type: Col3DConeClustered, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Cone Clustered Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AF30", opts: &Chart{Type: Col3DConePercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Cone Percent Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AF45", opts: &Chart{Type: Col3DCone, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Cone Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AN1", opts: &Chart{Type: Col3DPyramidStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Pyramid Percent Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AN16", opts: &Chart{Type: Col3DPyramidClustered, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Pyramid Clustered Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AN30", opts: &Chart{Type: Col3DPyramidPercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Pyramid Percent Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AN45", opts: &Chart{Type: Col3DPyramid, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Pyramid Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AV1", opts: &Chart{Type: Col3DCylinderStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Cylinder Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AV16", opts: &Chart{Type: Col3DCylinderClustered, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Cylinder Clustered Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AV30", opts: &Chart{Type: Col3DCylinderPercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Cylinder Percent Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "AV45", opts: &Chart{Type: Col3DCylinder, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Cylinder Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet1", cell: "P45", opts: &Chart{Type: Col3D, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Column Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", XAxis: ChartAxis{Alignment: Alignment{Vertical: "vert270", TextRotation: 0}}}},
		{sheetName: "Sheet2", cell: "P1", opts: &Chart{Type: Line3D, Series: series2, Format: format, Legend: ChartLegend{Position: "top", ShowLegendKey: false}, Title: []RichTextRun{{Text: "3D Line Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", XAxis: ChartAxis{MajorGridLines: true, MinorGridLines: true, TickLabelSkip: 1, NumFmt: ChartNumFmt{CustomNumFmt: "General"}}, YAxis: ChartAxis{MajorGridLines: true, MinorGridLines: true, MajorUnit: 1, NumFmt: ChartNumFmt{CustomNumFmt: "General"}}}},
		{sheetName: "Sheet2", cell: "X1", opts: &Chart{Type: Scatter, Series: series, Format: format, Legend: ChartLegend{Position: "bottom", ShowLegendKey: false}, Title: []RichTextRun{{Text: "Scatter Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "P16", opts: &Chart{Type: Doughnut, Series: series3, Format: format, Legend: ChartLegend{Position: "right", ShowLegendKey: false}, Title: []RichTextRun{{Text: "Doughnut Chart"}}, PlotArea: ChartPlotArea{ShowBubbleSize: false, ShowCatName: false, ShowLeaderLines: false, ShowPercent: true, ShowSerName: false, ShowVal: false}, ShowBlanksAs: "zero", HoleSize: 30}},
		{sheetName: "Sheet2", cell: "X16", opts: &Chart{Type: Line, Series: series2, Format: format, Legend: ChartLegend{Position: "top", ShowLegendKey: false}, Title: []RichTextRun{{Text: "Line Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", XAxis: ChartAxis{MajorGridLines: true, MinorGridLines: true, TickLabelSkip: 1, TickLabelPosition: ChartTickLabelLow}, YAxis: ChartAxis{MajorGridLines: true, MinorGridLines: true, MajorUnit: 1}}},
		{sheetName: "Sheet2", cell: "P32", opts: &Chart{Type: Pie3D, Series: series3, Format: format, Legend: ChartLegend{Position: "bottom", ShowLegendKey: false}, Title: []RichTextRun{{Text: "3D Column Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "X32", opts: &Chart{Type: Pie, Series: series3, Format: format, Legend: ChartLegend{Position: "bottom", ShowLegendKey: false}, Title: []RichTextRun{{Text: "Pie Chart"}}, PlotArea: ChartPlotArea{ShowBubbleSize: true, ShowCatName: false, ShowLeaderLines: false, ShowPercent: true, ShowSerName: false, ShowVal: false, NumFmt: ChartNumFmt{CustomNumFmt: "0.00%;0;;"}}, ShowBlanksAs: "gap"}},
		// bar series chart
		{sheetName: "Sheet2", cell: "P48", opts: &Chart{Type: Bar, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "2D Clustered Bar Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "X48", opts: &Chart{Type: BarStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "2D Stacked Bar Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "P64", opts: &Chart{Type: BarPercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "2D Stacked 100% Bar Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "X64", opts: &Chart{Type: Bar3DClustered, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Clustered Bar Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "P80", opts: &Chart{Type: Bar3DStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Stacked Bar Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", YAxis: ChartAxis{Maximum: &maximum, Minimum: &minimum}}},
		{sheetName: "Sheet2", cell: "X80", opts: &Chart{Type: Bar3DPercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D 100% Stacked Bar Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", XAxis: ChartAxis{ReverseOrder: true, Secondary: true, Minimum: &zero}, YAxis: ChartAxis{ReverseOrder: true, Minimum: &zero}}},
		// area series chart
		{sheetName: "Sheet2", cell: "AF1", opts: &Chart{Type: Area, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "2D Area Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AN1", opts: &Chart{Type: AreaStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "2D Stacked Area Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AF16", opts: &Chart{Type: AreaPercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "2D 100% Stacked Area Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AN16", opts: &Chart{Type: Area3D, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Area Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AF32", opts: &Chart{Type: Area3DStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Stacked Area Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AN32", opts: &Chart{Type: Area3DPercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D 100% Stacked Area Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		// cylinder series chart
		{sheetName: "Sheet2", cell: "AF48", opts: &Chart{Type: Bar3DCylinderStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Bar Cylinder Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AF64", opts: &Chart{Type: Bar3DCylinderClustered, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Bar Cylinder Clustered Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AF80", opts: &Chart{Type: Bar3DCylinderPercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Bar Cylinder Percent Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		// cone series chart
		{sheetName: "Sheet2", cell: "AN48", opts: &Chart{Type: Bar3DConeStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Bar Cone Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AN64", opts: &Chart{Type: Bar3DConeClustered, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Bar Cone Clustered Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AN80", opts: &Chart{Type: Bar3DConePercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Bar Cone Percent Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AV48", opts: &Chart{Type: Bar3DPyramidStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Bar Pyramid Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AV64", opts: &Chart{Type: Bar3DPyramidClustered, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Bar Pyramid Clustered Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "AV80", opts: &Chart{Type: Bar3DPyramidPercentStacked, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Bar Pyramid Percent Stacked Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		// surface series chart
		{sheetName: "Sheet2", cell: "AV1", opts: &Chart{Type: Surface3D, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Surface Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", YAxis: ChartAxis{MajorGridLines: true}}},
		{sheetName: "Sheet2", cell: "AV16", opts: &Chart{Type: WireframeSurface3D, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "3D Wireframe Surface Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", YAxis: ChartAxis{MajorGridLines: true}}},
		{sheetName: "Sheet2", cell: "AV32", opts: &Chart{Type: Contour, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "Contour Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		{sheetName: "Sheet2", cell: "BD1", opts: &Chart{Type: WireframeContour, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "Wireframe Contour Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}},
		// bubble chart
		{sheetName: "Sheet2", cell: "BD16", opts: &Chart{Type: Bubble, Series: series4, Format: format, Legend: legend, Title: []RichTextRun{{Text: "Bubble Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", BubbleSize: 75}},
		{sheetName: "Sheet2", cell: "BD32", opts: &Chart{Type: Bubble3D, Series: series4, Format: format, Legend: legend, Title: []RichTextRun{{Text: "Bubble 3D Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", XAxis: ChartAxis{MajorGridLines: true}, YAxis: ChartAxis{MajorGridLines: true}}},
		// pie of pie chart
		{sheetName: "Sheet2", cell: "BD48", opts: &Chart{Type: PieOfPie, Series: series3, Format: format, Legend: legend, Title: []RichTextRun{{Text: "Pie of Pie Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", XAxis: ChartAxis{MajorGridLines: true}, YAxis: ChartAxis{MajorGridLines: true}}},
		// bar of pie chart
		{sheetName: "Sheet2", cell: "BD64", opts: &Chart{Type: BarOfPie, Series: series3, Format: format, Legend: legend, Title: []RichTextRun{{Text: "Bar of Pie Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", XAxis: ChartAxis{MajorGridLines: true}, YAxis: ChartAxis{MajorGridLines: true}}},
	} {
		assert.NoError(t, f.AddChart(c.sheetName, c.cell, c.opts))
	}
	// combo chart
	_, err = f.NewSheet("Combo Charts")
	assert.NoError(t, err)
	clusteredColumnCombo := [][]interface{}{
		{"A1", Line, "Clustered Column - Line Chart"},
		{"I1", Doughnut, "Clustered Column - Doughnut Chart"},
	}
	for _, props := range clusteredColumnCombo {
		assert.NoError(t, f.AddChart("Combo Charts", props[0].(string), &Chart{Type: Col, Series: series[:4], Format: format, Legend: legend, Title: []RichTextRun{{Text: props[2].(string)}}, PlotArea: ChartPlotArea{ShowBubbleSize: true, ShowCatName: false, ShowLeaderLines: false, ShowPercent: true, ShowSerName: true, ShowVal: true}}, &Chart{Type: props[1].(ChartType), Series: series[4:], Format: format, Legend: legend, PlotArea: ChartPlotArea{ShowBubbleSize: true, ShowCatName: false, ShowLeaderLines: false, ShowPercent: true, ShowSerName: true, ShowVal: true}, YAxis: ChartAxis{Secondary: true}}))
	}
	stackedAreaCombo := map[string][]interface{}{
		"A16": {Line, "Stacked Area - Line Chart"},
		"I16": {Doughnut, "Stacked Area - Doughnut Chart"},
	}
	for axis, props := range stackedAreaCombo {
		assert.NoError(t, f.AddChart("Combo Charts", axis, &Chart{Type: AreaStacked, Series: series[:4], Format: format, Legend: legend, Title: []RichTextRun{{Text: props[1].(string)}}, PlotArea: ChartPlotArea{ShowBubbleSize: true, ShowCatName: false, ShowLeaderLines: false, ShowPercent: true, ShowSerName: true, ShowVal: true}}, &Chart{Type: props[0].(ChartType), Series: series[4:], Format: format, Legend: legend, PlotArea: ChartPlotArea{ShowBubbleSize: true, ShowCatName: false, ShowLeaderLines: false, ShowPercent: true, ShowSerName: true, ShowVal: true}}))
	}
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddChart.xlsx")))
	// Test with invalid sheet name
	assert.EqualError(t, f.AddChart("Sheet:1", "A1", &Chart{Type: Col, Series: series[:1]}), ErrSheetNameInvalid.Error())
	// Test with illegal cell reference
	assert.EqualError(t, f.AddChart("Sheet2", "A", &Chart{Type: Col, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "2D Column Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	// Test with unsupported chart type
	assert.EqualError(t, f.AddChart("Sheet2", "BD32", &Chart{Type: 0x37, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "Bubble 3D Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}), newUnsupportedChartType(0x37).Error())
	// Test add combo chart with invalid format set
	assert.EqualError(t, f.AddChart("Sheet2", "BD32", &Chart{Type: Col, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "2D Column Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}, nil), ErrParameterInvalid.Error())
	// Test add combo chart with unsupported chart type
	assert.EqualError(t, f.AddChart("Sheet2", "BD64", &Chart{Type: BarOfPie, Series: []ChartSeries{{Name: "Sheet1!$A$30", Categories: "Sheet1!$A$30:$D$37", Values: "Sheet1!$B$30:$B$37"}}, Format: format, Legend: legend, Title: []RichTextRun{{Text: "Bar of Pie Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", XAxis: ChartAxis{MajorGridLines: true}, YAxis: ChartAxis{MajorGridLines: true}}, &Chart{Type: 0x37, Series: []ChartSeries{{Name: "Sheet1!$A$30", Categories: "Sheet1!$A$30:$D$37", Values: "Sheet1!$B$30:$B$37"}}, Format: format, Legend: legend, Title: []RichTextRun{{Text: "Bar of Pie Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero", XAxis: ChartAxis{MajorGridLines: true}, YAxis: ChartAxis{MajorGridLines: true}}), newUnsupportedChartType(0x37).Error())
	assert.NoError(t, f.Close())

	// Test add chart with unsupported charset content types.
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, f.AddChart("Sheet1", "P1", &Chart{Type: Col, Series: []ChartSeries{{Name: "Sheet1!$A$30", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$30:$D$30"}}, Title: []RichTextRun{{Text: "2D Column Chart"}}}), "XML syntax error on line 1: invalid UTF-8")
}

func TestAddChartSheet(t *testing.T) {
	categories := map[string]string{"A2": "Small", "A3": "Normal", "A4": "Large", "B1": "Apple", "C1": "Orange", "D1": "Pear"}
	values := map[string]int{"B2": 2, "C2": 3, "D2": 3, "B3": 5, "C3": 2, "D3": 4, "B4": 6, "C4": 7, "D4": 8}
	f := NewFile()
	for k, v := range categories {
		assert.NoError(t, f.SetCellValue("Sheet1", k, v))
	}
	for k, v := range values {
		assert.NoError(t, f.SetCellValue("Sheet1", k, v))
	}
	series := []ChartSeries{
		{Name: "Sheet1!$A$2", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$2:$D$2"},
		{Name: "Sheet1!$A$3", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$3:$D$3"},
		{Name: "Sheet1!$A$4", Categories: "Sheet1!$B$1:$D$1", Values: "Sheet1!$B$4:$D$4"},
	}
	assert.NoError(t, f.AddChartSheet("Chart1", &Chart{Type: Col3DClustered, Series: series, Title: []RichTextRun{{Text: "Fruit 3D Clustered Column Chart"}}}))
	// Test set the chartsheet as active sheet
	var sheetIdx int
	for idx, sheetName := range f.GetSheetList() {
		if sheetName != "Chart1" {
			continue
		}
		sheetIdx = idx
	}
	f.SetActiveSheet(sheetIdx)

	// Test cell value on chartsheet
	assert.EqualError(t, f.SetCellValue("Chart1", "A1", true), "sheet Chart1 is not a worksheet")
	// Test add chartsheet on already existing name sheet

	assert.EqualError(t, f.AddChartSheet("Sheet1", &Chart{Type: Col3DClustered, Series: series, Title: []RichTextRun{{Text: "Fruit 3D Clustered Column Chart"}}}), ErrExistsSheet.Error())
	// Test add chartsheet with invalid sheet name
	assert.EqualError(t, f.AddChartSheet("Sheet:1", nil, &Chart{Type: Col3DClustered, Series: series, Title: []RichTextRun{{Text: "Fruit 3D Clustered Column Chart"}}}), ErrSheetNameInvalid.Error())
	// Test with unsupported chart type
	assert.EqualError(t, f.AddChartSheet("Chart2", &Chart{Type: 0x37, Series: series, Title: []RichTextRun{{Text: "Fruit 3D Clustered Column Chart"}}}), newUnsupportedChartType(0x37).Error())

	assert.NoError(t, f.UpdateLinkedValue())

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddChartSheet.xlsx")))
	// Test add chart sheet with unsupported charset content types
	f = NewFile()
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, f.AddChartSheet("Chart4", &Chart{Type: Col, Series: []ChartSeries{{Name: "Sheet1!$A$30", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$30:$D$30"}}, Title: []RichTextRun{{Text: "2D Column Chart"}}}), "XML syntax error on line 1: invalid UTF-8")
}

func TestDeleteChart(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)
	assert.NoError(t, f.DeleteChart("Sheet1", "A1"))
	series := []ChartSeries{
		{Name: "Sheet1!$A$30", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$30:$D$30"},
		{Name: "Sheet1!$A$31", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$31:$D$31"},
		{Name: "Sheet1!$A$32", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$32:$D$32"},
		{Name: "Sheet1!$A$33", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$33:$D$33"},
		{Name: "Sheet1!$A$34", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$34:$D$34"},
		{Name: "Sheet1!$A$35", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$35:$D$35"},
		{Name: "Sheet1!$A$36", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$36:$D$36"},
		{Name: "Sheet1!$A$37", Categories: "Sheet1!$B$29:$D$29", Values: "Sheet1!$B$37:$D$37"},
	}
	format := GraphicOptions{
		ScaleX:          defaultDrawingScale,
		ScaleY:          defaultDrawingScale,
		OffsetX:         15,
		OffsetY:         10,
		PrintObject:     boolPtr(true),
		LockAspectRatio: false,
		Locked:          boolPtr(false),
	}
	legend := ChartLegend{Position: "left", ShowLegendKey: false}
	plotArea := ChartPlotArea{
		ShowBubbleSize:  true,
		ShowCatName:     true,
		ShowLeaderLines: false,
		ShowPercent:     true,
		ShowSerName:     true,
		ShowVal:         true,
	}
	assert.NoError(t, f.AddChart("Sheet1", "P1", &Chart{Type: Col, Series: series, Format: format, Legend: legend, Title: []RichTextRun{{Text: "2D Column Chart"}}, PlotArea: plotArea, ShowBlanksAs: "zero"}))
	assert.NoError(t, f.DeleteChart("Sheet1", "P1"))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestDeleteChart.xlsx")))
	// Test delete chart with invalid sheet name
	assert.EqualError(t, f.DeleteChart("Sheet:1", "P1"), ErrSheetNameInvalid.Error())
	// Test delete chart on not exists worksheet
	assert.EqualError(t, f.DeleteChart("SheetN", "A1"), "sheet SheetN does not exist")
	// Test delete chart with invalid coordinates
	assert.EqualError(t, f.DeleteChart("Sheet1", ""), newCellNameToCoordinatesError("", newInvalidCellNameError("")).Error())
	// Test delete chart on no chart worksheet
	assert.NoError(t, NewFile().DeleteChart("Sheet1", "A1"))
	assert.NoError(t, f.Close())
}

func TestChartWithLogarithmicBase(t *testing.T) {
	// Create test workbook with data
	f := NewFile()
	sheet1 := f.GetSheetName(0)
	categories := map[string]float64{
		"A1":  1,
		"A2":  2,
		"A3":  3,
		"A4":  4,
		"A5":  5,
		"A6":  6,
		"A7":  7,
		"A8":  8,
		"A9":  9,
		"A10": 10,
		"B1":  0.1,
		"B2":  1,
		"B3":  2,
		"B4":  3,
		"B5":  20,
		"B6":  30,
		"B7":  100,
		"B8":  500,
		"B9":  700,
		"B10": 5000,
	}
	for cell, v := range categories {
		assert.NoError(t, f.SetCellValue(sheet1, cell, v))
	}
	series := []ChartSeries{{Name: "value", Categories: "Sheet1!$A$1:$A$19", Values: "Sheet1!$B$1:$B$10"}}
	dimension := []uint{640, 480, 320, 240}
	for _, c := range []struct {
		cell string
		opts *Chart
	}{
		{cell: "C1", opts: &Chart{Type: Line, Dimension: ChartDimension{Width: dimension[0], Height: dimension[1]}, Series: series, Title: []RichTextRun{{Text: "Line chart without log scaling"}}}},
		{cell: "M1", opts: &Chart{Type: Line, Dimension: ChartDimension{Width: dimension[0], Height: dimension[1]}, Series: series, Title: []RichTextRun{{Text: "Line chart with log 10.5 scaling"}}, YAxis: ChartAxis{LogBase: 10.5}}},
		{cell: "A25", opts: &Chart{Type: Line, Dimension: ChartDimension{Width: dimension[2], Height: dimension[3]}, Series: series, Title: []RichTextRun{{Text: "Line chart with log 1.9 scaling"}}, YAxis: ChartAxis{LogBase: 1.9}}},
		{cell: "F25", opts: &Chart{Type: Line, Dimension: ChartDimension{Width: dimension[2], Height: dimension[3]}, Series: series, Title: []RichTextRun{{Text: "Line chart with log 2 scaling"}}, YAxis: ChartAxis{LogBase: 2}}},
		{cell: "K25", opts: &Chart{Type: Line, Dimension: ChartDimension{Width: dimension[2], Height: dimension[3]}, Series: series, Title: []RichTextRun{{Text: "Line chart with log 1000.1 scaling"}}, YAxis: ChartAxis{LogBase: 1000.1}}},
		{cell: "P25", opts: &Chart{Type: Line, Dimension: ChartDimension{Width: dimension[2], Height: dimension[3]}, Series: series, Title: []RichTextRun{{Text: "Line chart with log 1000 scaling"}}, YAxis: ChartAxis{LogBase: 1000}}},
	} {
		// Add two chart, one without and one with log scaling
		assert.NoError(t, f.AddChart(sheet1, c.cell, c.opts))
	}

	// Export workbook for human confirmation
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestChartWithLogarithmicBase10.xlsx")))

	// Write the workbook to a buffer
	var buffer bytes.Buffer
	assert.NoError(t, f.Write(&buffer))

	// Read back the workbook from the buffer
	newFile, err := OpenReader(&buffer)
	assert.NoError(t, err)

	// Check the number of charts
	expectedChartsCount := 6
	chartsNum := newFile.countCharts()
	if !assert.Equal(t, expectedChartsCount, chartsNum,
		"Expected %d charts, actual %d", expectedChartsCount, chartsNum) {
		t.FailNow()
	}

	chartSpaces := make([]xlsxChartSpace, expectedChartsCount)
	type xmlChartContent []byte
	xmlCharts := make([]xmlChartContent, expectedChartsCount)
	expectedChartsLogBase := []float64{0, 10.5, 0, 2, 0, 1000}
	var (
		drawingML interface{}
		ok        bool
	)
	for i := 0; i < expectedChartsCount; i++ {
		chartPath := fmt.Sprintf("xl/charts/chart%d.xml", i+1)
		if drawingML, ok = newFile.Pkg.Load(chartPath); ok {
			xmlCharts[i] = drawingML.([]byte)
		}
		assert.True(t, ok, "Can't open the %s", chartPath)

		err = xml.Unmarshal(xmlCharts[i], &chartSpaces[i])
		if !assert.NoError(t, err) {
			t.FailNow()
		}

		chartLogBasePtr := chartSpaces[i].Chart.PlotArea.ValAx[0].Scaling.LogBase
		if expectedChartsLogBase[i] == 0 {
			if !assert.Nil(t, chartLogBasePtr, "LogBase is not nil") {
				t.FailNow()
			}
		} else {
			if !assert.NotNil(t, chartLogBasePtr, "LogBase is nil") {
				t.FailNow()
			}
			if !assert.Equal(t, expectedChartsLogBase[i], *(chartLogBasePtr.Val),
				"Expected log base to %f, actual %f", expectedChartsLogBase[i], *(chartLogBasePtr.Val)) {
				t.FailNow()
			}
		}
	}
}
