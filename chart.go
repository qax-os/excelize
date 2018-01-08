package excelize

import (
	"encoding/json"
	"encoding/xml"
	"strconv"
	"strings"
)

// This section defines the currently supported chart types.
const (
	Bar                 = "bar"
	BarStacked          = "barStacked"
	Bar3D               = "bar3D"
	Bar3DColumn         = "bar3DColumn"
	Bar3DPercentStacked = "bar3DPercentStacked"
	Doughnut            = "doughnut"
	Line                = "line"
	Pie                 = "pie"
	Pie3D               = "pie3D"
	Radar               = "radar"
	Scatter             = "scatter"
)

// This section defines the default value of chart properties.
var (
	chartView3DRotX = map[string]int{
		Bar:                 0,
		BarStacked:          0,
		Bar3D:               15,
		Bar3DColumn:         15,
		Bar3DPercentStacked: 15,
		Doughnut:            0,
		Line:                0,
		Pie:                 0,
		Pie3D:               30,
		Radar:               0,
		Scatter:             0,
	}
	chartView3DRotY = map[string]int{
		Bar:                 0,
		BarStacked:          0,
		Bar3D:               20,
		Bar3DColumn:         20,
		Bar3DPercentStacked: 20,
		Doughnut:            0,
		Line:                0,
		Pie:                 0,
		Pie3D:               0,
		Radar:               0,
		Scatter:             0,
	}
	chartView3DDepthPercent = map[string]int{
		Bar:                 100,
		BarStacked:          100,
		Bar3D:               100,
		Bar3DColumn:         100,
		Bar3DPercentStacked: 100,
		Doughnut:            100,
		Line:                100,
		Pie:                 100,
		Pie3D:               100,
		Radar:               100,
		Scatter:             100,
	}
	chartView3DRAngAx = map[string]int{
		Bar:                 0,
		BarStacked:          0,
		Bar3D:               1,
		Bar3DColumn:         1,
		Bar3DPercentStacked: 1,
		Doughnut:            0,
		Line:                0,
		Pie:                 0,
		Pie3D:               0,
		Radar:               0,
		Scatter:             0,
	}
	chartLegendPosition = map[string]string{
		"bottom":    "b",
		"left":      "l",
		"right":     "r",
		"top":       "t",
		"top_right": "tr",
	}
	chartValAxNumFmtFormatCode = map[string]string{
		Bar:                 "General",
		BarStacked:          "General",
		Bar3D:               "General",
		Bar3DColumn:         "General",
		Bar3DPercentStacked: "0%",
		Doughnut:            "General",
		Line:                "General",
		Pie:                 "General",
		Pie3D:               "General",
		Radar:               "General",
		Scatter:             "General",
	}
	plotAreaChartGrouping = map[string]string{
		Bar:                 "clustered",
		BarStacked:          "stacked",
		Bar3D:               "clustered",
		Bar3DColumn:         "standard",
		Bar3DPercentStacked: "percentStacked",
		Line:                "standard",
	}
)

// parseFormatChartSet provides function to parse the format settings of the
// chart with default value.
func parseFormatChartSet(formatSet string) *formatChart {
	format := formatChart{
		Format: formatPicture{
			FPrintsWithSheet: true,
			FLocksWithSheet:  false,
			NoChangeAspect:   false,
			OffsetX:          0,
			OffsetY:          0,
			XScale:           1.0,
			YScale:           1.0,
		},
		Legend: formatChartLegend{
			Position:      "bottom",
			ShowLegendKey: false,
		},
		Title: formatChartTitle{
			Name: " ",
		},
		ShowBlanksAs: "gap",
	}
	json.Unmarshal([]byte(formatSet), &format)
	return &format
}

// AddChart provides the method to add chart in a sheet by given chart format
// set (such as offset, scale, aspect ratio setting and print settings) and
// properties set. For example, create 3D bar chart with data
// Sheet1!$A$29:$D$32:
//
//    package main
//
//    import (
//        "fmt"
//
//        "github.com/360EntSecGroup-Skylar/excelize"
//    )
//
//    func main() {
//        categories := map[string]string{"A2": "Small", "A3": "Normal", "A4": "Large", "B1": "Apple", "C1": "Orange", "D1": "Pear"}
//        values := map[string]int{"B2": 2, "C2": 3, "D2": 3, "B3": 5, "C3": 2, "D3": 4, "B4": 6, "C4": 7, "D4": 8}
//        xlsx := excelize.NewFile()
//        for k, v := range categories {
//            xlsx.SetCellValue("Sheet1", k, v)
//        }
//        for k, v := range values {
//            xlsx.SetCellValue("Sheet1", k, v)
//        }
//        xlsx.AddChart("SHEET1", "F2", `{"type":"bar3D","series":[{"name":"=Sheet1!$A$30","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$30:$D$30"},{"name":"=Sheet1!$A$31","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$31:$D$31"},{"name":"=Sheet1!$A$32","categories":"=Sheet1!$B$29:$D$29","values":"=Sheet1!$B$32:$D$32"}],"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},"legend":{"position":"bottom","show_legend_key":false},"title":{"name":"Fruit Line Chart"},"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":true,"show_val":true},"show_blanks_as":"zero"}`)
//        // Save xlsx file by the given path.
//        err := xlsx.SaveAs("./Workbook.xlsx")
//        if err != nil {
//            fmt.Println(err)
//        }
//    }
//
// The following shows the type of chart supported by excelize:
//
//     Type                | Chart
//    ---------------------+---------------------------
//     bar                 | bar chart
//     barStacked          | stacked bar chart
//     bar3D               | 3D bar chart
//	   bar3DColumn 		   | 3D column bar chart
// 	   bar3DPercentStacked | 3D 100% stacked bar chart
//     doughnut            | doughnut chart
//     line                | line chart
//     pie                 | pie chart
//     pie3D               | 3D pie chart
//     radar               | radar chart
//     scatter             | scatter chart
//
// In Excel a chart series is a collection of information that defines which data is plotted such as values, axis labels and formatting.
//
// The series options that can be set are:
//
//    name
//    categories
//    values
//
// name: Set the name for the series. The name is displayed in the chart legend and in the formula bar. The name property is optional and if it isn't supplied it will default to Series 1..n. The name can also be a formula such as =Sheet1!$A$1
//
// categories: This sets the chart category labels. The category is more or less the same as the X axis. In most chart types the categories property is optional and the chart will just assume a sequential series from 1..n.
//
// values: This is the most important property of a series and is the only mandatory option for every chart object. This option links the chart with the worksheet data that it displays.
//
// Set properties of the chart legend. The options that can be set are:
//
//    position
//    show_legend_key
//
// position: Set the position of the chart legend. The default legend position is right. The available positions are:
//
//    top
//    bottom
//    left
//    right
//    top_right
//
// show_legend_key: Set the legend keys shall be shown in data labels. The default value is false.
//
// Set properties of the chart title. The properties that can be set are:
//
//    title
//
// name: Set the name (title) for the chart. The name is displayed above the chart. The name can also be a formula such as =Sheet1!$A$1 or a list with a sheetname. The name property is optional. The default is to have no chart title.
//
// Specifies how blank cells are plotted on the chart by show_blanks_as. The default value is gap. The options that can be set are:
//
//    gap
//    span
//    zero
//
// gap: Specifies that blank values shall be left as a gap.
//
// sapn: Specifies that blank values shall be spanned with a line.
//
// zero: Specifies that blank values shall be treated as zero.
//
// Set chart offset, scale, aspect ratio setting and print settings by format, same as function AddPicture.
//
// Set the position of the chart plot area by plotarea. The properties that can be set are:
//
//    show_bubble_size
//    show_cat_name
//    show_leader_lines
//    show_percent
//    show_series_name
//    show_val
//
// show_bubble_size: Specifies the bubble size shall be shown in a data label. The show_bubble_size property is optional. The default value is false.
//
// show_cat_name: Specifies that the category name shall be shown in the data label. The show_cat_name property is optional. The default value is true.
//
// show_leader_lines: Specifies leader lines shall be shown for data labels. The show_leader_lines property is optional. The default value is false.
//
// show_percent: Specifies that the percentage shall be shown in a data label. The show_percent property is optional. The default value is false.
//
// show_series_name: Specifies that the series name shall be shown in a data label. The show_series_name property is optional. The default value is false.
//
// show_val: Specifies that the value shall be shown in a data label. The show_val property is optional. The default value is false.
//
func (f *File) AddChart(sheet, cell, format string) {
	formatSet := parseFormatChartSet(format)
	// Read sheet data.
	xlsx := f.workSheetReader(sheet)
	// Add first picture for given sheet, create xl/drawings/ and xl/drawings/_rels/ folder.
	drawingID := f.countDrawings() + 1
	chartID := f.countCharts() + 1
	drawingXML := "xl/drawings/drawing" + strconv.Itoa(drawingID) + ".xml"
	drawingID, drawingXML = f.prepareDrawing(xlsx, drawingID, sheet, drawingXML)
	drawingRID := f.addDrawingRelationships(drawingID, SourceRelationshipChart, "../charts/chart"+strconv.Itoa(chartID)+".xml")
	f.addDrawingChart(sheet, drawingXML, cell, 480, 290, drawingRID, &formatSet.Format)
	f.addChart(formatSet)
	f.addContentTypePart(chartID, "chart")
	f.addContentTypePart(drawingID, "drawings")
}

// countCharts provides function to get chart files count storage in the
// folder xl/charts.
func (f *File) countCharts() int {
	count := 0
	for k := range f.XLSX {
		if strings.Contains(k, "xl/charts/chart") {
			count++
		}
	}
	return count
}

// prepareDrawing provides function to prepare drawing ID and XML by given
// drawingID, worksheet name and default drawingXML.
func (f *File) prepareDrawing(xlsx *xlsxWorksheet, drawingID int, sheet, drawingXML string) (int, string) {
	sheetRelationshipsDrawingXML := "../drawings/drawing" + strconv.Itoa(drawingID) + ".xml"
	if xlsx.Drawing != nil {
		// The worksheet already has a picture or chart relationships, use the relationships drawing ../drawings/drawing%d.xml.
		sheetRelationshipsDrawingXML = f.getSheetRelationshipsTargetByID(sheet, xlsx.Drawing.RID)
		drawingID, _ = strconv.Atoi(strings.TrimSuffix(strings.TrimPrefix(sheetRelationshipsDrawingXML, "../drawings/drawing"), ".xml"))
		drawingXML = strings.Replace(sheetRelationshipsDrawingXML, "..", "xl", -1)
	} else {
		// Add first picture for given sheet.
		rID := f.addSheetRelationships(sheet, SourceRelationshipDrawingML, sheetRelationshipsDrawingXML, "")
		f.addSheetDrawing(sheet, rID)
	}
	return drawingID, drawingXML
}

// addChart provides function to create chart as xl/charts/chart%d.xml by given
// format sets.
func (f *File) addChart(formatSet *formatChart) {
	count := f.countCharts()
	xlsxChartSpace := xlsxChartSpace{
		XMLNSc:         NameSpaceDrawingMLChart,
		XMLNSa:         NameSpaceDrawingML,
		XMLNSr:         SourceRelationship,
		XMLNSc16r2:     SourceRelationshipChart201506,
		Date1904:       &attrValBool{Val: false},
		Lang:           &attrValString{Val: "en-US"},
		RoundedCorners: &attrValBool{Val: false},
		Chart: cChart{
			Title: &cTitle{
				Tx: cTx{
					Rich: &cRich{
						P: aP{
							PPr: &aPPr{
								DefRPr: aRPr{
									Kern:   1200,
									Strike: "noStrike",
									U:      "none",
									Sz:     1400,
									SolidFill: &aSolidFill{
										SchemeClr: &aSchemeClr{
											Val: "tx1",
											LumMod: &attrValInt{
												Val: 65000,
											},
											LumOff: &attrValInt{
												Val: 35000,
											},
										},
									},
									Ea: &aEa{
										Typeface: "+mn-ea",
									},
									Cs: &aCs{
										Typeface: "+mn-cs",
									},
									Latin: &aLatin{
										Typeface: "+mn-lt",
									},
								},
							},
							R: &aR{
								RPr: aRPr{
									Lang:    "en-US",
									AltLang: "en-US",
								},
								T: formatSet.Title.Name,
							},
						},
					},
				},
				TxPr: cTxPr{
					P: aP{
						PPr: &aPPr{
							DefRPr: aRPr{
								Kern:   1200,
								U:      "none",
								Sz:     14000,
								Strike: "noStrike",
							},
						},
						EndParaRPr: &aEndParaRPr{
							Lang: "en-US",
						},
					},
				},
			},
			View3D: &cView3D{
				RotX:         &attrValInt{Val: chartView3DRotX[formatSet.Type]},
				RotY:         &attrValInt{Val: chartView3DRotY[formatSet.Type]},
				DepthPercent: &attrValInt{Val: chartView3DDepthPercent[formatSet.Type]},
				RAngAx:       &attrValInt{Val: chartView3DRAngAx[formatSet.Type]},
			},
			Floor: &cThicknessSpPr{
				Thickness: &attrValInt{Val: 0},
			},
			SideWall: &cThicknessSpPr{
				Thickness: &attrValInt{Val: 0},
			},
			BackWall: &cThicknessSpPr{
				Thickness: &attrValInt{Val: 0},
			},
			PlotArea: &cPlotArea{},
			Legend: &cLegend{
				LegendPos: &attrValString{Val: chartLegendPosition[formatSet.Legend.Position]},
				Overlay:   &attrValBool{Val: false},
			},

			PlotVisOnly:      &attrValBool{Val: false},
			DispBlanksAs:     &attrValString{Val: formatSet.ShowBlanksAs},
			ShowDLblsOverMax: &attrValBool{Val: false},
		},
		SpPr: &cSpPr{
			SolidFill: &aSolidFill{
				SchemeClr: &aSchemeClr{Val: "bg1"},
			},
			Ln: &aLn{
				W:    9525,
				Cap:  "flat",
				Cmpd: "sng",
				Algn: "ctr",
				SolidFill: &aSolidFill{
					SchemeClr: &aSchemeClr{Val: "tx1",
						LumMod: &attrValInt{
							Val: 15000,
						},
						LumOff: &attrValInt{
							Val: 85000,
						},
					},
				},
			},
		},
		PrintSettings: &cPrintSettings{
			PageMargins: &cPageMargins{
				B:      0.75,
				L:      0.7,
				R:      0.7,
				T:      0.7,
				Header: 0.3,
				Footer: 0.3,
			},
		},
	}
	plotAreaFunc := map[string]func(*formatChart) *cPlotArea{
		Bar:                 f.drawBarChart,
		BarStacked:          f.drawBarChart,
		Bar3D:               f.drawBarChart,
		Bar3DColumn:         f.drawBarChart,
		Bar3DPercentStacked: f.drawBarChart,
		Doughnut:            f.drawDoughnutChart,
		Line:                f.drawLineChart,
		Pie3D:               f.drawPie3DChart,
		Pie:                 f.drawPieChart,
		Radar:               f.drawRadarChart,
		Scatter:             f.drawScatterChart,
	}
	xlsxChartSpace.Chart.PlotArea = plotAreaFunc[formatSet.Type](formatSet)

	chart, _ := xml.Marshal(xlsxChartSpace)
	media := "xl/charts/chart" + strconv.Itoa(count+1) + ".xml"
	f.saveFileList(media, string(chart))
}

// drawBarChart provides function to draw the c:plotArea element for bar, barStacked and
// bar3D chart by given format sets.
func (f *File) drawBarChart(formatSet *formatChart) *cPlotArea {
	c := cCharts{
		BarDir: &attrValString{
			Val: "col",
		},
		Grouping: &attrValString{
			Val: "clustered",
		},
		VaryColors: &attrValBool{
			Val: true,
		},
		Ser:   f.drawChartSeries(formatSet),
		DLbls: f.drawChartDLbls(formatSet),
		AxID: []*attrValInt{
			{Val: 754001152},
			{Val: 753999904},
		},
	}
	c.Grouping.Val = plotAreaChartGrouping[formatSet.Type]
	if formatSet.Type == "barStacked" {
		c.Overlap = &attrValInt{Val: 100}
	}
	catAx := f.drawPlotAreaCatAx(formatSet)
	valAx := f.drawPlotAreaValAx(formatSet)
	charts := map[string]*cPlotArea{
		"bar": {
			BarChart: &c,
			CatAx:    catAx,
			ValAx:    valAx,
		},
		"barStacked": {
			BarChart: &c,
			CatAx:    catAx,
			ValAx:    valAx,
		},
		"bar3D": {
			Bar3DChart: &c,
			CatAx:      catAx,
			ValAx:      valAx,
		},
		"bar3DColumn": {
			Bar3DChart: &c,
			CatAx:      catAx,
			ValAx:      valAx,
		},
		"bar3DPercentStacked": {
			Bar3DChart: &c,
			CatAx:      catAx,
			ValAx:      valAx,
		},
	}
	return charts[formatSet.Type]
}

// drawDoughnutChart provides function to draw the c:plotArea element for
// doughnut chart by given format sets.
func (f *File) drawDoughnutChart(formatSet *formatChart) *cPlotArea {
	return &cPlotArea{
		DoughnutChart: &cCharts{
			VaryColors: &attrValBool{
				Val: true,
			},
			Ser:      f.drawChartSeries(formatSet),
			HoleSize: &attrValInt{Val: 75},
		},
	}
}

// drawLineChart provides function to draw the c:plotArea element for line chart
// by given format sets.
func (f *File) drawLineChart(formatSet *formatChart) *cPlotArea {
	return &cPlotArea{
		LineChart: &cCharts{
			Grouping: &attrValString{
				Val: plotAreaChartGrouping[formatSet.Type],
			},
			VaryColors: &attrValBool{
				Val: false,
			},
			Ser:   f.drawChartSeries(formatSet),
			DLbls: f.drawChartDLbls(formatSet),
			Smooth: &attrValBool{
				Val: false,
			},
			AxID: []*attrValInt{
				{Val: 754001152},
				{Val: 753999904},
			},
		},
		CatAx: f.drawPlotAreaCatAx(formatSet),
		ValAx: f.drawPlotAreaValAx(formatSet),
	}
}

// drawPieChart provides function to draw the c:plotArea element for pie chart
// by given format sets.
func (f *File) drawPieChart(formatSet *formatChart) *cPlotArea {
	return &cPlotArea{
		PieChart: &cCharts{
			VaryColors: &attrValBool{
				Val: true,
			},
			Ser: f.drawChartSeries(formatSet),
		},
	}
}

// drawPie3DChart provides function to draw the c:plotArea element for 3D pie
// chart by given format sets.
func (f *File) drawPie3DChart(formatSet *formatChart) *cPlotArea {
	return &cPlotArea{
		Pie3DChart: &cCharts{
			VaryColors: &attrValBool{
				Val: true,
			},
			Ser: f.drawChartSeries(formatSet),
		},
	}
}

// drawRadarChart provides function to draw the c:plotArea element for radar
// chart by given format sets.
func (f *File) drawRadarChart(formatSet *formatChart) *cPlotArea {
	return &cPlotArea{
		RadarChart: &cCharts{
			RadarStyle: &attrValString{
				Val: "marker",
			},
			VaryColors: &attrValBool{
				Val: false,
			},
			Ser:   f.drawChartSeries(formatSet),
			DLbls: f.drawChartDLbls(formatSet),
			AxID: []*attrValInt{
				{Val: 754001152},
				{Val: 753999904},
			},
		},
		CatAx: f.drawPlotAreaCatAx(formatSet),
		ValAx: f.drawPlotAreaValAx(formatSet),
	}
}

// drawScatterChart provides function to draw the c:plotArea element for scatter
// chart by given format sets.
func (f *File) drawScatterChart(formatSet *formatChart) *cPlotArea {
	return &cPlotArea{
		ScatterChart: &cCharts{
			ScatterStyle: &attrValString{
				Val: "smoothMarker", // line,lineMarker,marker,none,smooth,smoothMarker
			},
			VaryColors: &attrValBool{
				Val: false,
			},
			Ser:   f.drawChartSeries(formatSet),
			DLbls: f.drawChartDLbls(formatSet),
			AxID: []*attrValInt{
				{Val: 754001152},
				{Val: 753999904},
			},
		},
		CatAx: f.drawPlotAreaCatAx(formatSet),
		ValAx: f.drawPlotAreaValAx(formatSet),
	}
}

// drawChartSeries provides function to draw the c:ser element by given format
// sets.
func (f *File) drawChartSeries(formatSet *formatChart) *[]cSer {
	ser := []cSer{}
	for k := range formatSet.Series {
		ser = append(ser, cSer{
			IDx:   &attrValInt{Val: k},
			Order: &attrValInt{Val: k},
			Tx: &cTx{
				StrRef: &cStrRef{
					F: formatSet.Series[k].Name,
				},
			},
			SpPr:   f.drawChartSeriesSpPr(k, formatSet),
			Marker: f.drawChartSeriesMarker(k, formatSet),
			DPt:    f.drawChartSeriesDPt(k, formatSet),
			DLbls:  f.drawChartSeriesDLbls(formatSet),
			Cat:    f.drawChartSeriesCat(formatSet.Series[k], formatSet),
			Val:    f.drawChartSeriesVal(formatSet.Series[k], formatSet),
			XVal:   f.drawChartSeriesXVal(formatSet.Series[k], formatSet),
			YVal:   f.drawChartSeriesYVal(formatSet.Series[k], formatSet),
		})
	}
	return &ser
}

// drawChartSeriesSpPr provides function to draw the c:spPr element by given
// format sets.
func (f *File) drawChartSeriesSpPr(i int, formatSet *formatChart) *cSpPr {
	spPrScatter := &cSpPr{
		Ln: &aLn{
			W:      25400,
			NoFill: " ",
		},
	}
	spPrLine := &cSpPr{
		Ln: &aLn{
			W:   25400,
			Cap: "rnd", // rnd, sq, flat
			SolidFill: &aSolidFill{
				SchemeClr: &aSchemeClr{Val: "accent" + strconv.Itoa(i+1)},
			},
		},
	}
	chartSeriesSpPr := map[string]*cSpPr{Bar: nil, BarStacked: nil, Bar3D: nil, Bar3DColumn: nil, Bar3DPercentStacked: nil, Doughnut: nil, Line: spPrLine, Pie: nil, Pie3D: nil, Radar: nil, Scatter: spPrScatter}
	return chartSeriesSpPr[formatSet.Type]
}

// drawChartSeriesDPt provides function to draw the c:dPt element by given data
// index and format sets.
func (f *File) drawChartSeriesDPt(i int, formatSet *formatChart) []*cDPt {
	dpt := []*cDPt{{
		IDx:      &attrValInt{Val: i},
		Bubble3D: &attrValBool{Val: false},
		SpPr: &cSpPr{
			SolidFill: &aSolidFill{
				SchemeClr: &aSchemeClr{Val: "accent" + strconv.Itoa(i+1)},
			},
			Ln: &aLn{
				W:   25400,
				Cap: "rnd",
				SolidFill: &aSolidFill{
					SchemeClr: &aSchemeClr{Val: "lt" + strconv.Itoa(i+1)},
				},
			},
			Sp3D: &aSp3D{
				ContourW: 25400,
				ContourClr: &aContourClr{
					SchemeClr: &aSchemeClr{Val: "lt" + strconv.Itoa(i+1)},
				},
			},
		},
	}}
	chartSeriesDPt := map[string][]*cDPt{Bar: nil, BarStacked: nil, Bar3D: nil, Bar3DColumn: nil, Bar3DPercentStacked: nil, Doughnut: nil, Line: nil, Pie: dpt, Pie3D: dpt, Radar: nil, Scatter: nil}
	return chartSeriesDPt[formatSet.Type]
}

// drawChartSeriesCat provides function to draw the c:cat element by given chart
// series and format sets.
func (f *File) drawChartSeriesCat(v formatChartSeries, formatSet *formatChart) *cCat {
	cat := &cCat{
		StrRef: &cStrRef{
			F: v.Categories,
		},
	}
	chartSeriesCat := map[string]*cCat{Bar: cat, BarStacked: cat, Bar3D: cat, Bar3DColumn: cat, Bar3DPercentStacked: cat, Doughnut: cat, Line: cat, Pie: cat, Pie3D: cat, Radar: cat, Scatter: nil}
	return chartSeriesCat[formatSet.Type]
}

// drawChartSeriesVal provides function to draw the c:val element by given chart
// series and format sets.
func (f *File) drawChartSeriesVal(v formatChartSeries, formatSet *formatChart) *cVal {
	val := &cVal{
		NumRef: &cNumRef{
			F: v.Values,
		},
	}
	chartSeriesVal := map[string]*cVal{Bar: val, BarStacked: val, Bar3D: val, Bar3DColumn: val, Bar3DPercentStacked: val, Doughnut: val, Line: val, Pie: val, Pie3D: val, Radar: val, Scatter: nil}
	return chartSeriesVal[formatSet.Type]
}

// drawChartSeriesMarker provides function to draw the c:marker element by given
// data index and format sets.
func (f *File) drawChartSeriesMarker(i int, formatSet *formatChart) *cMarker {
	marker := &cMarker{
		Symbol: &attrValString{Val: "circle"},
		Size:   &attrValInt{Val: 5},
		SpPr: &cSpPr{
			SolidFill: &aSolidFill{
				SchemeClr: &aSchemeClr{
					Val: "accent" + strconv.Itoa(i+1),
				},
			},
			Ln: &aLn{
				W: 9252,
				SolidFill: &aSolidFill{
					SchemeClr: &aSchemeClr{
						Val: "accent" + strconv.Itoa(i+1),
					},
				},
			},
		},
	}
	chartSeriesMarker := map[string]*cMarker{Bar: nil, BarStacked: nil, Bar3D: nil, Bar3DColumn: nil, Bar3DPercentStacked: nil, Doughnut: nil, Line: nil, Pie: nil, Pie3D: nil, Radar: nil, Scatter: marker}
	return chartSeriesMarker[formatSet.Type]
}

// drawChartSeriesXVal provides function to draw the c:xVal element by given
// chart series and format sets.
func (f *File) drawChartSeriesXVal(v formatChartSeries, formatSet *formatChart) *cCat {
	cat := &cCat{
		StrRef: &cStrRef{
			F: v.Categories,
		},
	}
	chartSeriesXVal := map[string]*cCat{Bar: nil, BarStacked: nil, Bar3D: nil, Bar3DColumn: nil, Bar3DPercentStacked: nil, Doughnut: nil, Line: nil, Pie: nil, Pie3D: nil, Radar: nil, Scatter: cat}
	return chartSeriesXVal[formatSet.Type]
}

// drawChartSeriesYVal provides function to draw the c:yVal element by given
// chart series and format sets.
func (f *File) drawChartSeriesYVal(v formatChartSeries, formatSet *formatChart) *cVal {
	val := &cVal{
		NumRef: &cNumRef{
			F: v.Values,
		},
	}
	chartSeriesYVal := map[string]*cVal{Bar: nil, BarStacked: nil, Bar3D: nil, Bar3DColumn: nil, Bar3DPercentStacked: nil, Doughnut: nil, Line: nil, Pie: nil, Pie3D: nil, Radar: nil, Scatter: val}
	return chartSeriesYVal[formatSet.Type]
}

// drawChartDLbls provides function to draw the c:dLbls element by given format
// sets.
func (f *File) drawChartDLbls(formatSet *formatChart) *cDLbls {
	return &cDLbls{
		ShowLegendKey:   &attrValBool{Val: formatSet.Legend.ShowLegendKey},
		ShowVal:         &attrValBool{Val: formatSet.Plotarea.ShowVal},
		ShowCatName:     &attrValBool{Val: formatSet.Plotarea.ShowCatName},
		ShowSerName:     &attrValBool{Val: formatSet.Plotarea.ShowSerName},
		ShowBubbleSize:  &attrValBool{Val: formatSet.Plotarea.ShowBubbleSize},
		ShowPercent:     &attrValBool{Val: formatSet.Plotarea.ShowPercent},
		ShowLeaderLines: &attrValBool{Val: formatSet.Plotarea.ShowLeaderLines},
	}
}

// drawChartSeriesDLbls provides function to draw the c:dLbls element by given
// format sets.
func (f *File) drawChartSeriesDLbls(formatSet *formatChart) *cDLbls {
	dLbls := f.drawChartDLbls(formatSet)
	chartSeriesDLbls := map[string]*cDLbls{Bar: dLbls, BarStacked: dLbls, Bar3D: dLbls, Bar3DColumn: dLbls, Bar3DPercentStacked: dLbls, Doughnut: dLbls, Line: dLbls, Pie: dLbls, Pie3D: dLbls, Radar: dLbls, Scatter: nil}
	return chartSeriesDLbls[formatSet.Type]
}

// drawPlotAreaCatAx provides function to draw the c:catAx element.
func (f *File) drawPlotAreaCatAx(formatSet *formatChart) []*cAxs {
	return []*cAxs{
		{
			AxID: &attrValInt{Val: 754001152},
			Scaling: &cScaling{
				Orientation: &attrValString{Val: "minMax"},
			},
			Delete: &attrValBool{Val: false},
			AxPos:  &attrValString{Val: "b"},
			NumFmt: &cNumFmt{
				FormatCode:   "General",
				SourceLinked: true,
			},
			MajorTickMark: &attrValString{Val: "none"},
			MinorTickMark: &attrValString{Val: "none"},
			TickLblPos:    &attrValString{Val: "nextTo"},
			SpPr:          f.drawPlotAreaSpPr(),
			TxPr:          f.drawPlotAreaTxPr(),
			CrossAx:       &attrValInt{Val: 753999904},
			Crosses:       &attrValString{Val: "autoZero"},
			Auto:          &attrValBool{Val: true},
			LblAlgn:       &attrValString{Val: "ctr"},
			LblOffset:     &attrValInt{Val: 100},
			NoMultiLvlLbl: &attrValBool{Val: false},
		},
	}
}

// drawPlotAreaValAx provides function to draw the c:valAx element.
func (f *File) drawPlotAreaValAx(formatSet *formatChart) []*cAxs {
	return []*cAxs{
		{
			AxID: &attrValInt{Val: 753999904},
			Scaling: &cScaling{
				Orientation: &attrValString{Val: "minMax"},
			},
			Delete: &attrValBool{Val: false},
			AxPos:  &attrValString{Val: "l"},
			NumFmt: &cNumFmt{
				FormatCode:   chartValAxNumFmtFormatCode[formatSet.Type],
				SourceLinked: true,
			},
			MajorTickMark: &attrValString{Val: "none"},
			MinorTickMark: &attrValString{Val: "none"},
			TickLblPos:    &attrValString{Val: "nextTo"},
			SpPr:          f.drawPlotAreaSpPr(),
			TxPr:          f.drawPlotAreaTxPr(),
			CrossAx:       &attrValInt{Val: 754001152},
			Crosses:       &attrValString{Val: "autoZero"},
			CrossBetween:  &attrValString{Val: "between"},
		},
	}
}

// drawPlotAreaSpPr provides function to draw the c:spPr element.
func (f *File) drawPlotAreaSpPr() *cSpPr {
	return &cSpPr{
		Ln: &aLn{
			W:    9525,
			Cap:  "flat",
			Cmpd: "sng",
			Algn: "ctr",
			SolidFill: &aSolidFill{
				SchemeClr: &aSchemeClr{
					Val:    "tx1",
					LumMod: &attrValInt{Val: 15000},
					LumOff: &attrValInt{Val: 85000},
				},
			},
		},
	}
}

// drawPlotAreaTxPr provides function to draw the c:txPr element.
func (f *File) drawPlotAreaTxPr() *cTxPr {
	return &cTxPr{
		BodyPr: aBodyPr{
			Rot:              -60000000,
			SpcFirstLastPara: true,
			VertOverflow:     "ellipsis",
			Vert:             "horz",
			Wrap:             "square",
			Anchor:           "ctr",
			AnchorCtr:        true,
		},
		P: aP{
			PPr: &aPPr{
				DefRPr: aRPr{
					Sz:       900,
					B:        false,
					I:        false,
					U:        "none",
					Strike:   "noStrike",
					Kern:     1200,
					Baseline: 0,
					SolidFill: &aSolidFill{
						SchemeClr: &aSchemeClr{
							Val:    "tx1",
							LumMod: &attrValInt{Val: 15000},
							LumOff: &attrValInt{Val: 85000},
						},
					},
					Latin: &aLatin{Typeface: "+mn-lt"},
					Ea:    &aEa{Typeface: "+mn-ea"},
					Cs:    &aCs{Typeface: "+mn-cs"},
				},
			},
			EndParaRPr: &aEndParaRPr{Lang: "en-US"},
		},
	}
}

// drawingParser provides function to parse drawingXML. In order to solve the
// problem that the label structure is changed after serialization and
// deserialization, two different structures: decodeWsDr and encodeWsDr are
// defined.
func (f *File) drawingParser(drawingXML string, content *xlsxWsDr) int {
	cNvPrID := 1
	_, ok := f.XLSX[drawingXML]
	if ok { // Append Model
		decodeWsDr := decodeWsDr{}
		xml.Unmarshal([]byte(f.readXML(drawingXML)), &decodeWsDr)
		content.R = decodeWsDr.R
		cNvPrID = len(decodeWsDr.OneCellAnchor) + len(decodeWsDr.TwoCellAnchor) + 1
		for _, v := range decodeWsDr.OneCellAnchor {
			content.OneCellAnchor = append(content.OneCellAnchor, &xdrCellAnchor{
				EditAs:       v.EditAs,
				GraphicFrame: v.Content,
			})
		}
		for _, v := range decodeWsDr.TwoCellAnchor {
			content.TwoCellAnchor = append(content.TwoCellAnchor, &xdrCellAnchor{
				EditAs:       v.EditAs,
				GraphicFrame: v.Content,
			})
		}
	}
	return cNvPrID
}

// addDrawingChart provides function to add chart graphic frame by given sheet,
// drawingXML, cell, width, height, relationship index and format sets.
func (f *File) addDrawingChart(sheet, drawingXML, cell string, width, height, rID int, formatSet *formatPicture) {
	cell = strings.ToUpper(cell)
	fromCol := string(strings.Map(letterOnlyMapF, cell))
	fromRow, _ := strconv.Atoi(strings.Map(intOnlyMapF, cell))
	row := fromRow - 1
	col := TitleToNumber(fromCol)
	width = int(float64(width) * formatSet.XScale)
	height = int(float64(height) * formatSet.YScale)
	colStart, rowStart, _, _, colEnd, rowEnd, x2, y2 := f.positionObjectPixels(sheet, col, row, formatSet.OffsetX, formatSet.OffsetY, width, height)
	content := xlsxWsDr{}
	content.A = NameSpaceDrawingML
	content.Xdr = NameSpaceDrawingMLSpreadSheet
	cNvPrID := f.drawingParser(drawingXML, &content)
	twoCellAnchor := xdrCellAnchor{}
	twoCellAnchor.EditAs = "oneCell"
	from := xlsxFrom{}
	from.Col = colStart
	from.ColOff = formatSet.OffsetX * EMU
	from.Row = rowStart
	from.RowOff = formatSet.OffsetY * EMU
	to := xlsxTo{}
	to.Col = colEnd
	to.ColOff = x2 * EMU
	to.Row = rowEnd
	to.RowOff = y2 * EMU
	twoCellAnchor.From = &from
	twoCellAnchor.To = &to

	graphicFrame := xlsxGraphicFrame{
		NvGraphicFramePr: xlsxNvGraphicFramePr{
			CNvPr: &xlsxCNvPr{
				ID:   f.countCharts() + f.countMedia() + 1,
				Name: "Chart " + strconv.Itoa(cNvPrID),
			},
		},
		Graphic: &xlsxGraphic{
			GraphicData: &xlsxGraphicData{
				URI: NameSpaceDrawingMLChart,
				Chart: &xlsxChart{
					C:   NameSpaceDrawingMLChart,
					R:   SourceRelationship,
					RID: "rId" + strconv.Itoa(rID),
				},
			},
		},
	}
	graphic, _ := xml.Marshal(graphicFrame)
	twoCellAnchor.GraphicFrame = string(graphic)
	twoCellAnchor.ClientData = &xdrClientData{
		FLocksWithSheet:  formatSet.FLocksWithSheet,
		FPrintsWithSheet: formatSet.FPrintsWithSheet,
	}
	content.TwoCellAnchor = append(content.TwoCellAnchor, &twoCellAnchor)
	output, _ := xml.Marshal(content)
	f.saveFileList(drawingXML, string(output))
}
