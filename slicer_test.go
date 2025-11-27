package excelize

import (
	"fmt"
	"math/rand"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestSlicer(t *testing.T) {
	f := NewFile()
	disable, colName := false, "_!@#$%^&*()-+=|\\/<>"
	assert.NoError(t, f.SetCellValue("Sheet1", "B1", colName))
	// Create table in a worksheet
	assert.NoError(t, f.AddTable("Sheet1", &Table{
		Name:  "Table1",
		Range: "A1:D5",
	}))
	assert.NoError(t, f.AddSlicer("Sheet1", &SlicerOptions{
		Name:       "Column1",
		Cell:       "E1",
		TableSheet: "Sheet1",
		TableName:  "Table1",
		Caption:    "Column1",
	}))
	assert.NoError(t, f.AddSlicer("Sheet1", &SlicerOptions{
		Name:       "Column1",
		Cell:       "I1",
		TableSheet: "Sheet1",
		TableName:  "Table1",
		Caption:    "Column1",
	}))
	assert.NoError(t, f.AddSlicer("Sheet1", &SlicerOptions{
		Name:          colName,
		Cell:          "M1",
		TableSheet:    "Sheet1",
		TableName:     "Table1",
		Caption:       colName,
		Macro:         "Button1_Click",
		Width:         200,
		Height:        200,
		DisplayHeader: &disable,
		ItemDesc:      true,
	}))
	// Test get table slicers
	slicers, err := f.GetSlicers("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, "Column1", slicers[0].Name)
	assert.Equal(t, "E1", slicers[0].Cell)
	assert.Equal(t, "Sheet1", slicers[0].TableSheet)
	assert.Equal(t, "Table1", slicers[0].TableName)
	assert.Equal(t, "Column1", slicers[0].Caption)
	assert.Equal(t, "Column1 1", slicers[1].Name)
	assert.Equal(t, "I1", slicers[1].Cell)
	assert.Equal(t, "Sheet1", slicers[1].TableSheet)
	assert.Equal(t, "Table1", slicers[1].TableName)
	assert.Equal(t, "Column1", slicers[1].Caption)
	assert.Equal(t, colName, slicers[2].Name)
	assert.Equal(t, "M1", slicers[2].Cell)
	assert.Equal(t, "Sheet1", slicers[2].TableSheet)
	assert.Equal(t, "Table1", slicers[2].TableName)
	assert.Equal(t, colName, slicers[2].Caption)
	assert.Equal(t, "Button1_Click", slicers[2].Macro)
	assert.False(t, *slicers[2].DisplayHeader)
	assert.True(t, slicers[2].ItemDesc)
	// Test create two pivot tables in a new worksheet
	_, err = f.NewSheet("Sheet2")
	assert.NoError(t, err)
	// Create some data in a sheet
	month := []string{"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
	year := []int{2017, 2018, 2019}
	types := []string{"Meat", "Dairy", "Beverages", "Produce"}
	region := []string{"East", "West", "North", "South"}
	assert.NoError(t, f.SetSheetRow("Sheet2", "A1", &[]string{"Month", "Year", "Type", "Sales", "Region"}))
	for row := 2; row < 32; row++ {
		assert.NoError(t, f.SetCellValue("Sheet2", fmt.Sprintf("A%d", row), month[rand.Intn(12)]))
		assert.NoError(t, f.SetCellValue("Sheet2", fmt.Sprintf("B%d", row), year[rand.Intn(3)]))
		assert.NoError(t, f.SetCellValue("Sheet2", fmt.Sprintf("C%d", row), types[rand.Intn(4)]))
		assert.NoError(t, f.SetCellValue("Sheet2", fmt.Sprintf("D%d", row), rand.Intn(5000)))
		assert.NoError(t, f.SetCellValue("Sheet2", fmt.Sprintf("E%d", row), region[rand.Intn(4)]))
	}
	assert.NoError(t, f.AddPivotTable(&PivotTableOptions{
		DataRange:           "Sheet2!A1:E31",
		PivotTableRange:     "Sheet2!G2:M34",
		Name:                "PivotTable1",
		Rows:                []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Filter:              []PivotTableField{{Data: "Region"}},
		Columns:             []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:                []PivotTableField{{Data: "Sales", Subtotal: "Sum", Name: "Summarize by Sum"}},
		RowGrandTotals:      true,
		ColGrandTotals:      true,
		ShowDrill:           true,
		ShowRowHeaders:      true,
		ShowColHeaders:      true,
		ShowLastColumn:      true,
		ShowError:           true,
		PivotTableStyleName: "PivotStyleLight16",
	}))
	assert.NoError(t, f.AddPivotTable(&PivotTableOptions{
		DataRange:       "Sheet2!A1:E31",
		PivotTableRange: "Sheet2!U34:O2",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales", Subtotal: "Average", Name: "Summarize by Average"}},
		RowGrandTotals:  true,
		ColGrandTotals:  true,
		ShowDrill:       true,
		ShowRowHeaders:  true,
		ShowColHeaders:  true,
		ShowLastColumn:  true,
	}))
	// Test add a pivot table slicer
	assert.NoError(t, f.AddSlicer("Sheet2", &SlicerOptions{
		Name:       "Month",
		Cell:       "G42",
		TableSheet: "Sheet2",
		TableName:  "PivotTable1",
		Caption:    "Month",
	}))
	// Test add a pivot table slicer with duplicate field name
	assert.NoError(t, f.AddSlicer("Sheet2", &SlicerOptions{
		Name:       "Month",
		Cell:       "K42",
		TableSheet: "Sheet2",
		TableName:  "PivotTable1",
		Caption:    "Month",
	}))
	// Test add a pivot table slicer for another pivot table in a worksheet
	assert.NoError(t, f.AddSlicer("Sheet2", &SlicerOptions{
		Name:       "Region",
		Cell:       "O42",
		TableSheet: "Sheet2",
		TableName:  "PivotTable2",
		Caption:    "Region",
		ItemDesc:   true,
	}))
	// Test get pivot table slicers
	slicers, err = f.GetSlicers("Sheet2")
	assert.NoError(t, err)
	assert.Equal(t, "Month", slicers[0].Name)
	assert.Equal(t, "G42", slicers[0].Cell)
	assert.Equal(t, "Sheet2", slicers[0].TableSheet)
	assert.Equal(t, "PivotTable1", slicers[0].TableName)
	assert.Equal(t, "Month", slicers[0].Caption)
	assert.Equal(t, "Month 1", slicers[1].Name)
	assert.Equal(t, "K42", slicers[1].Cell)
	assert.Equal(t, "Sheet2", slicers[1].TableSheet)
	assert.Equal(t, "PivotTable1", slicers[1].TableName)
	assert.Equal(t, "Month", slicers[1].Caption)
	assert.Equal(t, "Region", slicers[2].Name)
	assert.Equal(t, "O42", slicers[2].Cell)
	assert.Equal(t, "Sheet2", slicers[2].TableSheet)
	assert.Equal(t, "PivotTable2", slicers[2].TableName)
	assert.Equal(t, "Region", slicers[2].Caption)
	assert.True(t, slicers[2].ItemDesc)
	// Test add a table slicer with empty slicer options
	assert.Equal(t, ErrParameterRequired, f.AddSlicer("Sheet1", nil))
	// Test add a table slicer with invalid slicer options
	for _, opts := range []*SlicerOptions{
		{Cell: "Q1", TableSheet: "Sheet1", TableName: "Table1"},
		{Name: "Column", Cell: "Q1", TableSheet: "Sheet1"},
		{Name: "Column", TableSheet: "Sheet1", TableName: "Table1"},
	} {
		assert.Equal(t, ErrParameterInvalid, f.AddSlicer("Sheet1", opts))
	}
	// Test add a table slicer with not exist worksheet
	assert.EqualError(t, f.AddSlicer("SheetN", &SlicerOptions{
		Name:       "Column2",
		Cell:       "Q1",
		TableSheet: "SheetN",
		TableName:  "Table1",
	}), "sheet SheetN does not exist")
	// Test add a table slicer with not exist table name
	assert.Equal(t, newNoExistTableError("Table2"), f.AddSlicer("Sheet1", &SlicerOptions{
		Name:       "Column2",
		Cell:       "Q1",
		TableSheet: "Sheet1",
		TableName:  "Table2",
	}))
	// Test add a table slicer with invalid slicer name
	assert.Equal(t, newInvalidSlicerNameError("Column6"), f.AddSlicer("Sheet1", &SlicerOptions{
		Name:       "Column6",
		Cell:       "Q1",
		TableSheet: "Sheet1",
		TableName:  "Table1",
	}))
	workbookPath := filepath.Join("test", "TestAddSlicer.xlsm")
	file, err := os.ReadFile(filepath.Join("test", "vbaProject.bin"))
	assert.NoError(t, err)
	assert.NoError(t, f.AddVBAProject(file))
	assert.NoError(t, f.SaveAs(workbookPath))
	assert.NoError(t, f.Close())

	// Test add a pivot table slicer with unsupported charset pivot table
	f, err = OpenFile(workbookPath)
	assert.NoError(t, err)
	f.Pkg.Store("xl/pivotTables/pivotTable2.xml", MacintoshCyrillicCharset)
	assert.EqualError(t, f.AddSlicer("Sheet2", &SlicerOptions{
		Name:       "Month",
		Cell:       "G42",
		TableSheet: "Sheet2",
		TableName:  "PivotTable1",
		Caption:    "Month",
	}), "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())

	// Test open a workbook and get already exist slicers
	f, err = OpenFile(workbookPath)
	assert.NoError(t, err)
	slicers, err = f.GetSlicers("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, "Column1", slicers[0].Name)
	assert.Equal(t, "E1", slicers[0].Cell)
	assert.Equal(t, "Sheet1", slicers[0].TableSheet)
	assert.Equal(t, "Table1", slicers[0].TableName)
	assert.Equal(t, "Column1", slicers[0].Caption)
	assert.Equal(t, "Column1 1", slicers[1].Name)
	assert.Equal(t, "I1", slicers[1].Cell)
	assert.Equal(t, "Sheet1", slicers[1].TableSheet)
	assert.Equal(t, "Table1", slicers[1].TableName)
	assert.Equal(t, "Column1", slicers[1].Caption)
	assert.Equal(t, colName, slicers[2].Name)
	assert.Equal(t, "M1", slicers[2].Cell)
	assert.Equal(t, "Sheet1", slicers[2].TableSheet)
	assert.Equal(t, "Table1", slicers[2].TableName)
	assert.Equal(t, colName, slicers[2].Caption)
	assert.Equal(t, "Button1_Click", slicers[2].Macro)
	assert.False(t, *slicers[2].DisplayHeader)
	assert.True(t, slicers[2].ItemDesc)
	slicers, err = f.GetSlicers("Sheet2")
	assert.NoError(t, err)
	assert.Equal(t, "Month", slicers[0].Name)
	assert.Equal(t, "G42", slicers[0].Cell)
	assert.Equal(t, "Sheet2", slicers[0].TableSheet)
	assert.Equal(t, "PivotTable1", slicers[0].TableName)
	assert.Equal(t, "Month", slicers[0].Caption)
	assert.Equal(t, "Month 1", slicers[1].Name)
	assert.Equal(t, "K42", slicers[1].Cell)
	assert.Equal(t, "Sheet2", slicers[1].TableSheet)
	assert.Equal(t, "PivotTable1", slicers[1].TableName)
	assert.Equal(t, "Month", slicers[1].Caption)
	assert.Equal(t, "Region", slicers[2].Name)
	assert.Equal(t, "O42", slicers[2].Cell)
	assert.Equal(t, "Sheet2", slicers[2].TableSheet)
	assert.Equal(t, "PivotTable2", slicers[2].TableName)
	assert.Equal(t, "Region", slicers[2].Caption)
	assert.True(t, slicers[2].ItemDesc)

	// Test add a pivot table slicer with workbook which contains timeline
	f, err = OpenFile(workbookPath)
	assert.NoError(t, err)
	f.Pkg.Store("xl/timelines/timeline1.xml", []byte(fmt.Sprintf(`<timelines xmlns="%s"><timeline name="a"/></timelines>`, NameSpaceSpreadSheetX15.Value)))
	assert.NoError(t, f.AddSlicer("Sheet2", &SlicerOptions{
		Name:       "Month",
		Cell:       "G42",
		TableSheet: "Sheet2",
		TableName:  "PivotTable1",
		Caption:    "Month",
	}))
	assert.NoError(t, f.Close())

	// Test add a pivot table slicer with unsupported charset timeline
	f, err = OpenFile(workbookPath)
	assert.NoError(t, err)
	f.Pkg.Store("xl/timelines/timeline1.xml", MacintoshCyrillicCharset)
	assert.NoError(t, f.AddSlicer("Sheet2", &SlicerOptions{
		Name:       "Month",
		Cell:       "G42",
		TableSheet: "Sheet2",
		TableName:  "PivotTable1",
		Caption:    "Month",
	}))
	assert.NoError(t, f.Close())

	// Test add a table slicer with invalid worksheet extension list
	f = NewFile()
	assert.NoError(t, f.AddTable("Sheet1", &Table{
		Name:  "Table1",
		Range: "A1:D5",
	}))
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).ExtLst = &xlsxExtLst{Ext: "<>"}
	assert.Error(t, f.AddSlicer("Sheet1", &SlicerOptions{
		Name:       "Column1",
		Cell:       "E1",
		TableSheet: "Sheet1",
		TableName:  "Table1",
	}))
	assert.NoError(t, f.Close())

	// Test add a table slicer with unsupported charset slicer
	f = NewFile()
	assert.NoError(t, f.AddTable("Sheet1", &Table{
		Name:  "Table1",
		Range: "A1:D5",
	}))
	f.Pkg.Store("xl/slicers/slicer2.xml", MacintoshCyrillicCharset)
	assert.EqualError(t, f.AddSlicer("Sheet1", &SlicerOptions{
		Name:       "Column1",
		Cell:       "E1",
		TableName:  "Table1",
		TableSheet: "Sheet1",
	}), "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())

	// Test add a table slicer with read workbook error
	f = NewFile()
	assert.NoError(t, f.AddTable("Sheet1", &Table{
		Name:  "Table1",
		Range: "A1:D5",
	}))
	f.WorkBook.ExtLst = &xlsxExtLst{Ext: "<>"}
	assert.Error(t, f.AddSlicer("Sheet1", &SlicerOptions{
		Name:       "Column1",
		Cell:       "E1",
		TableName:  "Table1",
		TableSheet: "Sheet1",
	}))
	assert.NoError(t, f.Close())

	// Test add a table slicer with unsupported charset content types
	f = NewFile()
	assert.NoError(t, f.AddTable("Sheet1", &Table{
		Name:  "Table1",
		Range: "A1:D5",
	}))
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, f.AddSlicer("Sheet1", &SlicerOptions{
		Name:       "Column1",
		Cell:       "E1",
		TableName:  "Table1",
		TableSheet: "Sheet1",
	}), "XML syntax error on line 1: invalid UTF-8")
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, f.addSlicer(0, xlsxSlicer{}), "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())

	f = NewFile()
	// Create table in a worksheet
	assert.NoError(t, f.AddTable("Sheet1", &Table{
		Name:  "Table1",
		Range: "A1:D5",
	}))
	f.Pkg.Store("xl/drawings/drawing2.xml", MacintoshCyrillicCharset)
	assert.EqualError(t, f.AddSlicer("Sheet1", &SlicerOptions{
		Name:       "Column1",
		Cell:       "E1",
		TableSheet: "Sheet1",
		TableName:  "Table1",
		Caption:    "Column1",
	}), "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())

	f = NewFile()
	// Test get sheet slicers without slicer
	slicers, err = f.GetSlicers("Sheet1")
	assert.NoError(t, err)
	assert.Empty(t, slicers)
	// Test get sheet slicers with not exist worksheet name
	_, err = f.GetSlicers("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	assert.NoError(t, f.Close())

	f, err = OpenFile(workbookPath)
	assert.NoError(t, err)
	// Test get sheet slicers with unsupported charset slicer cache
	f.Pkg.Store("xl/slicerCaches/slicerCache1.xml", MacintoshCyrillicCharset)
	_, err = f.GetSlicers("Sheet1")
	assert.NoError(t, err)
	// Test get sheet slicers with unsupported charset slicer
	f.Pkg.Store("xl/slicers/slicer1.xml", MacintoshCyrillicCharset)
	_, err = f.GetSlicers("Sheet1")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	// Test get sheet slicers with invalid worksheet extension list
	ws, ok = f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).ExtLst.Ext = "<>"
	_, err = f.GetSlicers("Sheet1")
	assert.Error(t, err)
	assert.NoError(t, f.Close())

	f, err = OpenFile(workbookPath)
	assert.NoError(t, err)
	// Test get sheet slicers without slicer cache
	f.Pkg.Range(func(k, v interface{}) bool {
		if strings.Contains(k.(string), "xl/slicerCaches/slicerCache") {
			f.Pkg.Delete(k.(string))
		}
		return true
	})
	slicers, err = f.GetSlicers("Sheet1")
	assert.NoError(t, err)
	assert.Empty(t, slicers)
	assert.NoError(t, f.Close())
	// Test open a workbook and get sheet slicer with invalid cell reference in the drawing part
	f, err = OpenFile(workbookPath)
	assert.NoError(t, err)
	f.Pkg.Store("xl/drawings/drawing1.xml", []byte(fmt.Sprintf(`<wsDr xmlns="%s"><twoCellAnchor><from><col>-1</col><row>-1</row></from><mc:AlternateContent><mc:Choice xmlns:sle15="%s"><graphicFrame><nvGraphicFramePr><cNvPr id="2" name="Column1"/></nvGraphicFramePr></graphicFrame></mc:Choice></mc:AlternateContent></twoCellAnchor></wsDr>`, NameSpaceDrawingMLSpreadSheet.Value, NameSpaceDrawingMLSlicerX15.Value)))
	_, err = f.GetSlicers("Sheet1")
	assert.Equal(t, newCoordinatesToCellNameError(0, 0), err)
	// Test get sheet slicer without slicer shape in the drawing part
	f.Drawings.Delete("xl/drawings/drawing1.xml")
	f.Pkg.Store("xl/drawings/drawing1.xml", []byte(fmt.Sprintf(`<wsDr xmlns="%s"><twoCellAnchor/></wsDr>`, NameSpaceDrawingMLSpreadSheet.Value)))
	_, err = f.GetSlicers("Sheet1")
	assert.NoError(t, err)
	f.Drawings.Delete("xl/drawings/drawing1.xml")
	// Test get sheet slicers with unsupported charset drawing part
	f.Pkg.Store("xl/drawings/drawing1.xml", MacintoshCyrillicCharset)
	_, err = f.GetSlicers("Sheet1")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	// Test get sheet slicers with unsupported charset table
	f.Pkg.Store("xl/tables/table1.xml", MacintoshCyrillicCharset)
	_, err = f.GetSlicers("Sheet1")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	// Test get sheet slicers with unsupported charset pivot table
	f.Pkg.Store("xl/pivotTables/pivotTable1.xml", MacintoshCyrillicCharset)
	_, err = f.GetSlicers("Sheet2")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())

	// Test create a workbook and get sheet slicer with invalid cell reference in the drawing part
	f = NewFile()
	assert.NoError(t, f.AddTable("Sheet1", &Table{
		Name:  "Table1",
		Range: "A1:D5",
	}))
	assert.NoError(t, f.AddSlicer("Sheet1", &SlicerOptions{
		Name:       "Column1",
		Cell:       "E1",
		TableSheet: "Sheet1",
		TableName:  "Table1",
		Caption:    "Column1",
	}))
	drawing, ok := f.Drawings.Load("xl/drawings/drawing1.xml")
	assert.True(t, ok)
	drawing.(*xlsxWsDr).TwoCellAnchor[0].From = &xlsxFrom{Col: -1, Row: -1}
	_, err = f.GetSlicers("Sheet1")
	assert.Equal(t, newCoordinatesToCellNameError(0, 0), err)
	assert.NoError(t, f.Close())

	// Test open a workbook and delete slicers
	f, err = OpenFile(workbookPath)
	assert.NoError(t, err)
	for _, name := range []string{colName, "Column1 1", "Column1"} {
		assert.NoError(t, f.DeleteSlicer(name))
	}
	for _, name := range []string{"Month", "Month 1", "Region"} {
		assert.NoError(t, f.DeleteSlicer(name))
	}
	// Test delete slicer with no exits slicer name
	assert.Equal(t, newNoExistSlicerError("x"), f.DeleteSlicer("x"))
	assert.NoError(t, f.Close())

	// Test open a workbook and delete sheet slicer with unsupported charset slicer cache
	f, err = OpenFile(workbookPath)
	assert.NoError(t, err)
	f.Pkg.Store("xl/slicers/slicer1.xml", MacintoshCyrillicCharset)
	assert.EqualError(t, f.DeleteSlicer("Column1"), "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())
}

func TestAddSheetSlicer(t *testing.T) {
	f := NewFile()
	// Test add sheet slicer with not exist worksheet name
	_, err := f.addSheetSlicer("SheetN", ExtURISlicerListX15)
	assert.EqualError(t, err, "sheet SheetN does not exist")
	assert.NoError(t, f.Close())
}

func TestAddSheetTableSlicer(t *testing.T) {
	f := NewFile()
	// Test add sheet table slicer with invalid worksheet extension
	assert.Error(t, f.addSheetTableSlicer(&xlsxWorksheet{ExtLst: &xlsxExtLst{Ext: "<>"}}, 0, ExtURISlicerListX15))
	// Test add sheet table slicer with existing worksheet extension
	assert.NoError(t, f.addSheetTableSlicer(&xlsxWorksheet{ExtLst: &xlsxExtLst{Ext: fmt.Sprintf("<ext uri=\"%s\"></ext>", ExtURITimelineRefs)}}, 1, ExtURISlicerListX15))
	assert.NoError(t, f.Close())
}

func TestSetSlicerCache(t *testing.T) {
	f := NewFile()
	f.Pkg.Store("xl/slicerCaches/slicerCache1.xml", MacintoshCyrillicCharset)
	_, err := f.setSlicerCache(1, &SlicerOptions{}, &Table{}, nil)
	assert.NoError(t, err)
	assert.NoError(t, f.Close())

	f = NewFile()

	f.Pkg.Store("xl/slicerCaches/slicerCache2.xml", []byte(fmt.Sprintf(`<slicerCacheDefinition xmlns="%s" name="Slicer2" sourceName="B1"><extLst><ext uri="%s"/></extLst></slicerCacheDefinition>`, NameSpaceSpreadSheetX14.Value, ExtURISlicerCacheDefinition)))
	_, err = f.setSlicerCache(1, &SlicerOptions{}, &Table{}, nil)
	assert.NoError(t, err)
	assert.NoError(t, f.Close())

	f = NewFile()
	f.Pkg.Store("xl/slicerCaches/slicerCache2.xml", []byte(fmt.Sprintf(`<slicerCacheDefinition xmlns="%s" name="Slicer1" sourceName="B1"><extLst><ext uri="%s"/></extLst></slicerCacheDefinition>`, NameSpaceSpreadSheetX14.Value, ExtURISlicerCacheDefinition)))
	_, err = f.setSlicerCache(1, &SlicerOptions{}, &Table{}, nil)
	assert.NoError(t, err)
	assert.NoError(t, f.Close())

	f = NewFile()
	f.Pkg.Store("xl/slicerCaches/slicerCache2.xml", []byte(fmt.Sprintf(`<slicerCacheDefinition xmlns="%s" name="Slicer1" sourceName="B1"><extLst><ext uri="%s"><tableSlicerCache tableId="1" column="2"/></ext></extLst></slicerCacheDefinition>`, NameSpaceSpreadSheetX14.Value, ExtURISlicerCacheDefinition)))
	_, err = f.setSlicerCache(1, &SlicerOptions{}, &Table{tID: 1}, nil)
	assert.NoError(t, err)
	assert.NoError(t, f.Close())

	f = NewFile()
	f.Pkg.Store("xl/slicerCaches/slicerCache2.xml", []byte(fmt.Sprintf(`<slicerCacheDefinition xmlns="%s" name="Slicer1" sourceName="B1"></slicerCacheDefinition>`, NameSpaceSpreadSheetX14.Value)))
	_, err = f.setSlicerCache(1, &SlicerOptions{}, &Table{tID: 1}, nil)
	assert.NoError(t, err)
	assert.NoError(t, f.Close())
}

func TestDeleteSlicer(t *testing.T) {
	f, slicerXML := NewFile(), "xl/slicers/slicer1.xml"
	assert.NoError(t, f.AddTable("Sheet1", &Table{
		Name:  "Table1",
		Range: "A1:D5",
	}))
	assert.NoError(t, f.AddSlicer("Sheet1", &SlicerOptions{
		Name:       "Column1",
		Cell:       "E1",
		TableSheet: "Sheet1",
		TableName:  "Table1",
		Caption:    "Column1",
	}))
	// Test delete sheet slicers with invalid worksheet extension list
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).ExtLst.Ext = "<>"
	assert.Error(t, f.deleteSlicer(SlicerOptions{
		slicerXML:       slicerXML,
		slicerSheetName: "Sheet1",
		Name:            "Column1",
	}))
	// Test delete slicer with unsupported charset worksheet
	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", MacintoshCyrillicCharset)
	assert.EqualError(t, f.deleteSlicer(SlicerOptions{
		slicerXML:       slicerXML,
		slicerSheetName: "Sheet1",
		Name:            "Column1",
	}), "XML syntax error on line 1: invalid UTF-8")
	// Test delete slicer with unsupported charset slicer
	f.Pkg.Store(slicerXML, MacintoshCyrillicCharset)
	assert.EqualError(t, f.deleteSlicer(SlicerOptions{slicerXML: slicerXML}), "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())
}

func TestDeleteSlicerCache(t *testing.T) {
	f := NewFile()
	// Test delete slicer cache with unsupported charset workbook
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	assert.EqualError(t, f.deleteSlicerCache(nil, SlicerOptions{}), "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())
}

func TestAddSlicerCache(t *testing.T) {
	f := NewFile()
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, f.addSlicerCache("Slicer1", 0, &SlicerOptions{}, &Table{}, nil), "XML syntax error on line 1: invalid UTF-8")
	// Test add a pivot table cache slicer with unsupported charset
	pivotCacheXML := "xl/pivotCache/pivotCacheDefinition1.xml"
	f.Pkg.Store(pivotCacheXML, MacintoshCyrillicCharset)
	assert.EqualError(t, f.addSlicerCache("Slicer1", 0, &SlicerOptions{}, nil,
		&PivotTableOptions{pivotCacheXML: pivotCacheXML}), "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())
}

func TestAddDrawingSlicer(t *testing.T) {
	f := NewFile()
	// Test add a drawing slicer with not exist worksheet
	assert.EqualError(t, f.addDrawingSlicer("SheetN", "Column2", NameSpaceDrawingMLSlicerX15, &SlicerOptions{
		Name:       "Column2",
		Cell:       "Q1",
		TableSheet: "SheetN",
		TableName:  "Table1",
	}), "sheet SheetN does not exist")
	// Test add a drawing slicer with invalid cell reference
	assert.EqualError(t, f.addDrawingSlicer("Sheet1", "Column2", NameSpaceDrawingMLSlicerX15, &SlicerOptions{
		Name:       "Column2",
		Cell:       "A",
		TableSheet: "Sheet1",
		TableName:  "Table1",
	}), "cannot convert cell \"A\" to coordinates: invalid cell name \"A\"")
	assert.NoError(t, f.Close())
}

func TestAddWorkbookSlicerCache(t *testing.T) {
	// Test add a workbook slicer cache with unsupported charset workbook
	f := NewFile()
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	assert.EqualError(t, f.addWorkbookSlicerCache(1, ExtURISlicerCachesX15), "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())
}

func TestGenSlicerCacheName(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetDefinedName(&DefinedName{Name: "Slicer_Column_1", RefersTo: formulaErrorNA}))
	assert.Equal(t, "Slicer_Column_11", f.genSlicerCacheName("Column 1"))
	assert.NoError(t, f.Close())
}

func TestAddPivotCacheSlicer(t *testing.T) {
	f := NewFile()
	pivotCacheXML := "xl/pivotCache/pivotCacheDefinition1.xml"
	// Test add a pivot table cache slicer with existing extension list
	f.Pkg.Store(pivotCacheXML, []byte(fmt.Sprintf(`<pivotCacheDefinition xmlns="%s"><extLst><ext uri="%s"><x14:pivotCacheDefinition pivotCacheId="1"/></ext></extLst></pivotCacheDefinition>`, NameSpaceSpreadSheet.Value, ExtURIPivotCacheDefinition)))
	_, err := f.addPivotCacheSlicer(&PivotTableOptions{
		pivotCacheXML: pivotCacheXML,
	})
	assert.NoError(t, err)
}
