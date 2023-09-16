package excelize

import (
	"fmt"
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
		Name:    "Column1",
		Table:   "Table1",
		Cell:    "E1",
		Caption: "Column1",
	}))
	assert.NoError(t, f.AddSlicer("Sheet1", &SlicerOptions{
		Name:    "Column1",
		Table:   "Table1",
		Cell:    "I1",
		Caption: "Column1",
	}))
	assert.NoError(t, f.AddSlicer("Sheet1", &SlicerOptions{
		Name:          colName,
		Table:         "Table1",
		Cell:          "M1",
		Caption:       colName,
		Macro:         "Button1_Click",
		Width:         200,
		Height:        200,
		DisplayHeader: &disable,
		ItemDesc:      true,
	}))
	// Test add a table slicer with empty slicer options
	assert.Equal(t, ErrParameterRequired, f.AddSlicer("Sheet1", nil))
	// Test add a table slicer with invalid slicer options
	for _, opts := range []*SlicerOptions{
		{Table: "Table1", Cell: "Q1"},
		{Name: "Column", Cell: "Q1"},
		{Name: "Column", Table: "Table1"},
	} {
		assert.Equal(t, ErrParameterInvalid, f.AddSlicer("Sheet1", opts))
	}
	// Test add a table slicer with not exist worksheet
	assert.EqualError(t, f.AddSlicer("SheetN", &SlicerOptions{
		Name:  "Column2",
		Table: "Table1",
		Cell:  "Q1",
	}), "sheet SheetN does not exist")
	// Test add a table slicer with not exist table name
	assert.Equal(t, newNoExistTableError("Table2"), f.AddSlicer("Sheet1", &SlicerOptions{
		Name:  "Column2",
		Table: "Table2",
		Cell:  "Q1",
	}))
	// Test add a table slicer with invalid slicer name
	assert.Equal(t, newInvalidSlicerNameError("Column6"), f.AddSlicer("Sheet1", &SlicerOptions{
		Name:  "Column6",
		Table: "Table1",
		Cell:  "Q1",
	}))
	file, err := os.ReadFile(filepath.Join("test", "vbaProject.bin"))
	assert.NoError(t, err)
	assert.NoError(t, f.AddVBAProject(file))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddSlicer.xlsm")))
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
		Name:  "Column1",
		Table: "Table1",
		Cell:  "E1",
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
		Name:  "Column1",
		Table: "Table1",
		Cell:  "E1",
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
		Name:  "Column1",
		Table: "Table1",
		Cell:  "E1",
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
		Name:  "Column1",
		Table: "Table1",
		Cell:  "E1",
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
		Name:    "Column1",
		Table:   "Table1",
		Cell:    "E1",
		Caption: "Column1",
	}), "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())
}

func TestAddSheetSlicer(t *testing.T) {
	f := NewFile()
	// Test add sheet slicer with not exist worksheet name
	_, err := f.addSheetSlicer("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	assert.NoError(t, f.Close())
}

func TestAddSheetTableSlicer(t *testing.T) {
	f := NewFile()
	// Test add sheet table slicer with invalid worksheet extension
	assert.Error(t, f.addSheetTableSlicer(&xlsxWorksheet{ExtLst: &xlsxExtLst{Ext: "<>"}}, 0))
	// Test add sheet table slicer with existing worksheet extension
	assert.NoError(t, f.addSheetTableSlicer(&xlsxWorksheet{ExtLst: &xlsxExtLst{Ext: fmt.Sprintf("<ext uri=\"%s\"></ext>", ExtURITimelineRefs)}}, 1))
	assert.NoError(t, f.Close())
}

func TestSetSlicerCache(t *testing.T) {
	f := NewFile()
	f.Pkg.Store("xl/slicerCaches/slicerCache1.xml", MacintoshCyrillicCharset)
	_, err := f.setSlicerCache(1, &SlicerOptions{}, &Table{})
	assert.NoError(t, err)
	assert.NoError(t, f.Close())

	f = NewFile()

	f.Pkg.Store("xl/slicerCaches/slicerCache2.xml", []byte(fmt.Sprintf(`<slicerCacheDefinition xmlns="%s" name="Slicer2" sourceName="B1"><extLst><ext uri="%s"/></extLst></slicerCacheDefinition>`, NameSpaceSpreadSheetX14.Value, ExtURISlicerCacheDefinition)))
	_, err = f.setSlicerCache(1, &SlicerOptions{}, &Table{})
	assert.NoError(t, err)
	assert.NoError(t, f.Close())

	f = NewFile()
	f.Pkg.Store("xl/slicerCaches/slicerCache2.xml", []byte(fmt.Sprintf(`<slicerCacheDefinition xmlns="%s" name="Slicer1" sourceName="B1"><extLst><ext uri="%s"/></extLst></slicerCacheDefinition>`, NameSpaceSpreadSheetX14.Value, ExtURISlicerCacheDefinition)))
	_, err = f.setSlicerCache(1, &SlicerOptions{}, &Table{})
	assert.NoError(t, err)
	assert.NoError(t, f.Close())

	f = NewFile()
	f.Pkg.Store("xl/slicerCaches/slicerCache2.xml", []byte(fmt.Sprintf(`<slicerCacheDefinition xmlns="%s" name="Slicer1" sourceName="B1"><extLst><ext uri="%s"><tableSlicerCache tableId="1" column="2"/></ext></extLst></slicerCacheDefinition>`, NameSpaceSpreadSheetX14.Value, ExtURISlicerCacheDefinition)))
	_, err = f.setSlicerCache(1, &SlicerOptions{}, &Table{tID: 1})
	assert.NoError(t, err)
	assert.NoError(t, f.Close())

	f = NewFile()
	f.Pkg.Store("xl/slicerCaches/slicerCache2.xml", []byte(fmt.Sprintf(`<slicerCacheDefinition xmlns="%s" name="Slicer1" sourceName="B1"></slicerCacheDefinition>`, NameSpaceSpreadSheetX14.Value)))
	_, err = f.setSlicerCache(1, &SlicerOptions{}, &Table{tID: 1})
	assert.NoError(t, err)
	assert.NoError(t, f.Close())
}

func TestAddSlicerCache(t *testing.T) {
	f := NewFile()
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, f.addSlicerCache("Slicer1", 0, &SlicerOptions{}, &Table{}), "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())
}

func TestAddDrawingSlicer(t *testing.T) {
	f := NewFile()
	// Test add a drawing slicer with not exist worksheet
	_, err := f.addDrawingSlicer("SheetN", &SlicerOptions{
		Name:  "Column2",
		Table: "Table1",
		Cell:  "Q1",
	})
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Test add a drawing slicer with invalid cell reference
	_, err = f.addDrawingSlicer("Sheet1", &SlicerOptions{
		Name:  "Column2",
		Table: "Table1",
		Cell:  "A",
	})
	assert.EqualError(t, err, "cannot convert cell \"A\" to coordinates: invalid cell name \"A\"")
	assert.NoError(t, f.Close())
}

func TestAddWorkbookSlicerCache(t *testing.T) {
	// Test add a workbook slicer cache with with unsupported charset workbook
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
