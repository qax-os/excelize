package excelize

import (
	"fmt"
	"math/rand"
	"os"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestAddSlicer(t *testing.T) {
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
	// Test create two pivot tables in a new worksheet
	_, err := f.NewSheet("Sheet2")
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
	_, err := f.setSlicerCache("Sheet1", 1, &SlicerOptions{}, &Table{}, nil)
	assert.NoError(t, err)
	assert.NoError(t, f.Close())

	f = NewFile()

	f.Pkg.Store("xl/slicerCaches/slicerCache2.xml", []byte(fmt.Sprintf(`<slicerCacheDefinition xmlns="%s" name="Slicer2" sourceName="B1"><extLst><ext uri="%s"/></extLst></slicerCacheDefinition>`, NameSpaceSpreadSheetX14.Value, ExtURISlicerCacheDefinition)))
	_, err = f.setSlicerCache("Sheet1", 1, &SlicerOptions{}, &Table{}, nil)
	assert.NoError(t, err)
	assert.NoError(t, f.Close())

	f = NewFile()
	f.Pkg.Store("xl/slicerCaches/slicerCache2.xml", []byte(fmt.Sprintf(`<slicerCacheDefinition xmlns="%s" name="Slicer1" sourceName="B1"><extLst><ext uri="%s"/></extLst></slicerCacheDefinition>`, NameSpaceSpreadSheetX14.Value, ExtURISlicerCacheDefinition)))
	_, err = f.setSlicerCache("Sheet1", 1, &SlicerOptions{}, &Table{}, nil)
	assert.NoError(t, err)
	assert.NoError(t, f.Close())

	f = NewFile()
	f.Pkg.Store("xl/slicerCaches/slicerCache2.xml", []byte(fmt.Sprintf(`<slicerCacheDefinition xmlns="%s" name="Slicer1" sourceName="B1"><extLst><ext uri="%s"><tableSlicerCache tableId="1" column="2"/></ext></extLst></slicerCacheDefinition>`, NameSpaceSpreadSheetX14.Value, ExtURISlicerCacheDefinition)))
	_, err = f.setSlicerCache("Sheet1", 1, &SlicerOptions{}, &Table{tID: 1}, nil)
	assert.NoError(t, err)
	assert.NoError(t, f.Close())

	f = NewFile()
	f.Pkg.Store("xl/slicerCaches/slicerCache2.xml", []byte(fmt.Sprintf(`<slicerCacheDefinition xmlns="%s" name="Slicer1" sourceName="B1"></slicerCacheDefinition>`, NameSpaceSpreadSheetX14.Value)))
	_, err = f.setSlicerCache("Sheet1", 1, &SlicerOptions{}, &Table{tID: 1}, nil)
	assert.NoError(t, err)
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
	f.Pkg.Store("xl/workbook.xml", MacintoshCyrillicCharset)
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
