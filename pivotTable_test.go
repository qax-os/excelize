package excelize

import (
	"fmt"
	"math/rand"
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestPivotTable(t *testing.T) {
	f := NewFile()
	// Create some data in a sheet
	month := []string{"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
	year := []int{2017, 2018, 2019}
	types := []string{"Meat", "Dairy", "Beverages", "Produce"}
	region := []string{"East", "West", "North", "South"}
	assert.NoError(t, f.SetSheetRow("Sheet1", "A1", &[]string{"Month", "Year", "Type", "Sales", "Region"}))
	for row := 2; row < 32; row++ {
		assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), month[rand.Intn(12)]))
		assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("B%d", row), year[rand.Intn(3)]))
		assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("C%d", row), types[rand.Intn(4)]))
		assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("D%d", row), rand.Intn(5000)))
		assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("E%d", row), region[rand.Intn(4)]))
	}
	expected := &PivotTableOptions{
		pivotTableXML:       "xl/pivotTables/pivotTable1.xml",
		pivotCacheXML:       "xl/pivotCache/pivotCacheDefinition1.xml",
		DataRange:           "Sheet1!A1:E31",
		PivotTableRange:     "Sheet1!G2:M34",
		Name:                "PivotTable1",
		Rows:                []PivotTableField{{Data: "Month", ShowAll: true, DefaultSubtotal: true}, {Data: "Year"}},
		Filter:              []PivotTableField{{Data: "Region"}},
		Columns:             []PivotTableField{{Data: "Type", ShowAll: true, InsertBlankRow: true, DefaultSubtotal: true}},
		Data:                []PivotTableField{{Data: "Sales", Subtotal: "Sum", Name: "Summarize by Sum", NumFmt: 38}},
		RowGrandTotals:      true,
		ColGrandTotals:      true,
		ShowDrill:           true,
		ClassicLayout:       true,
		ShowError:           true,
		ShowRowHeaders:      true,
		ShowColHeaders:      true,
		ShowLastColumn:      true,
		FieldPrintTitles:    true,
		ItemPrintTitles:     true,
		PivotTableStyleName: "PivotStyleLight16",
	}
	assert.NoError(t, f.AddPivotTable(expected))
	// Test get pivot table
	pivotTables, err := f.GetPivotTables("Sheet1")
	assert.NoError(t, err)
	assert.Len(t, pivotTables, 1)
	assert.Equal(t, *expected, pivotTables[0])
	// Use different order of coordinate tests
	assert.NoError(t, f.AddPivotTable(&PivotTableOptions{
		DataRange:       "Sheet1!A1:E31",
		PivotTableRange: "Sheet1!U34:O2",
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
	// Test get pivot table with default style name
	pivotTables, err = f.GetPivotTables("Sheet1")
	assert.NoError(t, err)
	assert.Len(t, pivotTables, 2)
	assert.Equal(t, "PivotStyleLight16", pivotTables[1].PivotTableStyleName)

	assert.NoError(t, f.AddPivotTable(&PivotTableOptions{
		DataRange:       "Sheet1!A1:E31",
		PivotTableRange: "Sheet1!W2:AC34",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Region"}},
		Data:            []PivotTableField{{Data: "Sales", Subtotal: "Count", Name: "Summarize by Count"}},
		RowGrandTotals:  true,
		ColGrandTotals:  true,
		ShowDrill:       true,
		ShowRowHeaders:  true,
		ShowColHeaders:  true,
		ShowLastColumn:  true,
	}))
	assert.NoError(t, f.AddPivotTable(&PivotTableOptions{
		DataRange:       "Sheet1!A1:E31",
		PivotTableRange: "Sheet1!G42:W55",
		Rows:            []PivotTableField{{Data: "Month"}},
		Columns:         []PivotTableField{{Data: "Region", DefaultSubtotal: true}, {Data: "Year"}},
		Data:            []PivotTableField{{Data: "Sales", Subtotal: "CountNums", Name: "Summarize by CountNums"}},
		RowGrandTotals:  true,
		ColGrandTotals:  true,
		ShowDrill:       true,
		ShowRowHeaders:  true,
		ShowColHeaders:  true,
		ShowLastColumn:  true,
	}))
	assert.NoError(t, f.AddPivotTable(&PivotTableOptions{
		DataRange:       "Sheet1!A1:E31",
		PivotTableRange: "Sheet1!AE2:AG33",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Data:            []PivotTableField{{Data: "Sales", Subtotal: "Max", Name: "Summarize by Max"}, {Data: "Sales", Subtotal: "Average", Name: "Average of Sales"}},
		RowGrandTotals:  true,
		ColGrandTotals:  true,
		ShowDrill:       true,
		ShowRowHeaders:  true,
		ShowColHeaders:  true,
		ShowLastColumn:  true,
	}))
	// Create pivot table with empty subtotal field name and specified style
	assert.NoError(t, f.AddPivotTable(&PivotTableOptions{
		DataRange:           "Sheet1!A1:E31",
		PivotTableRange:     "Sheet1!AJ2:AP135",
		Rows:                []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Filter:              []PivotTableField{{Data: "Region"}},
		Columns:             []PivotTableField{},
		Data:                []PivotTableField{{Subtotal: "Sum", Name: "Summarize by Sum"}},
		RowGrandTotals:      true,
		ColGrandTotals:      true,
		ShowDrill:           true,
		ShowRowHeaders:      true,
		ShowColHeaders:      true,
		ShowLastColumn:      true,
		PivotTableStyleName: "PivotStyleLight19",
	}))
	_, err = f.NewSheet("Sheet2")
	assert.NoError(t, err)
	assert.NoError(t, f.AddPivotTable(&PivotTableOptions{
		DataRange:       "Sheet1!A1:E31",
		PivotTableRange: "Sheet2!A1:AN17",
		Rows:            []PivotTableField{{Data: "Month"}},
		Columns:         []PivotTableField{{Data: "Region", DefaultSubtotal: true}, {Data: "Type", DefaultSubtotal: true}, {Data: "Year"}},
		Data:            []PivotTableField{{Data: "Sales", Subtotal: "Min", Name: "Summarize by Min", NumFmt: 32}},
		RowGrandTotals:  true,
		ColGrandTotals:  true,
		ShowDrill:       true,
		ShowRowHeaders:  true,
		ShowColHeaders:  true,
		ShowLastColumn:  true,
	}))

	// Test get pivot table with across worksheet data range
	pivotTables, err = f.GetPivotTables("Sheet2")
	assert.NoError(t, err)
	assert.Len(t, pivotTables, 1)
	assert.Equal(t, "Sheet1!A1:E31", pivotTables[0].DataRange)

	assert.NoError(t, f.AddPivotTable(&PivotTableOptions{
		DataRange:       "Sheet1!A1:E31",
		PivotTableRange: "Sheet2!A20:AR60",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Type"}},
		Columns:         []PivotTableField{{Data: "Region", DefaultSubtotal: true}, {Data: "Year"}},
		Data:            []PivotTableField{{Data: "Sales", Subtotal: "Product", Name: "Summarize by Product", NumFmt: 32}},
		RowGrandTotals:  true,
		ColGrandTotals:  true,
		ShowDrill:       true,
		ShowRowHeaders:  true,
		ShowColHeaders:  true,
		ShowLastColumn:  true,
	}))
	// Create pivot table with many data, many rows, many cols and defined name
	assert.NoError(t, f.SetDefinedName(&DefinedName{
		Name:     "dataRange",
		RefersTo: "Sheet1!A1:E31",
		Comment:  "Pivot Table Data Range",
		Scope:    "Sheet2",
	}))
	assert.NoError(t, f.AddPivotTable(&PivotTableOptions{
		DataRange:       "dataRange",
		PivotTableRange: "Sheet2!A65:AJ100",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Region", DefaultSubtotal: true}, {Data: "Type"}},
		Data:            []PivotTableField{{Data: "Sales", Subtotal: "Sum", Name: "Sum of Sales", NumFmt: -1}, {Data: "Sales", Subtotal: "Average", Name: "Average of Sales", NumFmt: 38}},
		RowGrandTotals:  true,
		ColGrandTotals:  true,
		ShowDrill:       true,
		ShowRowHeaders:  true,
		ShowColHeaders:  true,
		ShowLastColumn:  true,
	}))

	// Test empty pivot table options
	assert.Equal(t, ErrParameterRequired, f.AddPivotTable(nil))
	// Test add pivot table with custom name which exceeds the max characters limit
	assert.Equal(t, ErrNameLength, f.AddPivotTable(&PivotTableOptions{
		DataRange:       "dataRange",
		PivotTableRange: "Sheet2!A65:AJ100",
		Name:            strings.Repeat("c", MaxFieldLength+1),
	}))
	// Test invalid data range
	assert.Equal(t, newPivotTableDataRangeError("parameter is invalid"), f.AddPivotTable(&PivotTableOptions{
		DataRange:       "Sheet1!A1:A1",
		PivotTableRange: "Sheet1!U34:O2",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales"}},
	}))
	// Test the data range of the worksheet that is not declared
	assert.Equal(t, newPivotTableDataRangeError("parameter is invalid"), f.AddPivotTable(&PivotTableOptions{
		DataRange:       "A1:E31",
		PivotTableRange: "Sheet1!U34:O2",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales"}},
	}))
	// Test the worksheet declared in the data range does not exist
	assert.Equal(t, ErrSheetNotExist{"SheetN"}, f.AddPivotTable(&PivotTableOptions{
		DataRange:       "SheetN!A1:E31",
		PivotTableRange: "Sheet1!U34:O2",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales"}},
	}))
	// Test the pivot table range of the worksheet that is not declared
	assert.Equal(t, newPivotTableRangeError("parameter is invalid"), f.AddPivotTable(&PivotTableOptions{
		DataRange:       "Sheet1!A1:E31",
		PivotTableRange: "U34:O2",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales"}},
	}))
	// Test the worksheet declared in the pivot table range does not exist
	assert.Equal(t, ErrSheetNotExist{"SheetN"}, f.AddPivotTable(&PivotTableOptions{
		DataRange:       "Sheet1!A1:E31",
		PivotTableRange: "SheetN!U34:O2",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales"}},
	}))
	// Test not exists worksheet in data range
	assert.Equal(t, ErrSheetNotExist{"SheetN"}, f.AddPivotTable(&PivotTableOptions{
		DataRange:       "SheetN!A1:E31",
		PivotTableRange: "Sheet1!U34:O2",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales"}},
	}))
	// Test invalid row number in data range
	assert.Equal(t, newPivotTableDataRangeError(newCellNameToCoordinatesError("A0", newInvalidCellNameError("A0")).Error()), f.AddPivotTable(&PivotTableOptions{
		DataRange:       "Sheet1!A0:E31",
		PivotTableRange: "Sheet1!U34:O2",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales"}},
	}))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddPivotTable1.xlsx")))
	// Test with field names that exceed the length limit and invalid subtotal
	assert.NoError(t, f.AddPivotTable(&PivotTableOptions{
		DataRange:       "Sheet1!A1:E31",
		PivotTableRange: "Sheet1!G2:M34",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales", Subtotal: "-", Name: strings.Repeat("s", MaxFieldLength+1)}},
	}))
	// Test delete pivot table
	pivotTables, err = f.GetPivotTables("Sheet1")
	assert.Len(t, pivotTables, 7)
	assert.NoError(t, err)
	assert.NoError(t, f.DeletePivotTable("Sheet1", "PivotTable1"))
	pivotTables, err = f.GetPivotTables("Sheet1")
	assert.Len(t, pivotTables, 6)
	assert.NoError(t, err)

	// Test add pivot table with invalid sheet name
	assert.Error(t, f.AddPivotTable(&PivotTableOptions{
		DataRange:       "Sheet:1!A1:E31",
		PivotTableRange: "Sheet:1!G2:M34",
		Rows:            []PivotTableField{{Data: "Year"}},
	}), ErrSheetNameInvalid)
	// Test add pivot table with enable ClassicLayout and CompactData in the same time
	assert.Error(t, f.AddPivotTable(&PivotTableOptions{
		DataRange:       "Sheet1!A1:E31",
		PivotTableRange: "Sheet1!G2:M34",
		CompactData:     true,
		ClassicLayout:   true,
	}), ErrPivotTableClassicLayout)
	// Test delete pivot table with not exists worksheet
	assert.EqualError(t, f.DeletePivotTable("SheetN", "PivotTable1"), "sheet SheetN does not exist")
	// Test delete pivot table with not exists pivot table name
	assert.EqualError(t, f.DeletePivotTable("Sheet1", "PivotTableN"), "table PivotTableN does not exist")
	// Test adjust range with invalid range
	_, _, err = f.adjustRange("")
	assert.Error(t, err, ErrParameterRequired)
	// Test adjust range with incorrect range
	_, _, err = f.adjustRange("sheet1!")
	assert.EqualError(t, err, "parameter is invalid")
	// Test get table fields order with empty data range
	_, err = f.getTableFieldsOrder(&PivotTableOptions{})
	assert.EqualError(t, err, `parameter 'DataRange' parsing error: parameter is required`)
	// Test add pivot cache with empty data range
	assert.EqualError(t, f.addPivotCache(&PivotTableOptions{}), "parameter 'DataRange' parsing error: parameter is required")
	// Test add pivot table with empty options
	assert.EqualError(t, f.addPivotTable(0, 0, &PivotTableOptions{}), "parameter 'PivotTableRange' parsing error: parameter is required")
	// Test add pivot table with invalid data range
	assert.EqualError(t, f.addPivotTable(0, 0, &PivotTableOptions{}), "parameter 'PivotTableRange' parsing error: parameter is required")
	// Test add pivot fields with empty data range
	assert.EqualError(t, f.addPivotFields(nil, &PivotTableOptions{
		DataRange:       "A1:E31",
		PivotTableRange: "Sheet1!U34:O2",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales"}},
	}), `parameter 'DataRange' parsing error: parameter is invalid`)
	// Test get pivot fields index with empty data range
	_, err = f.getPivotFieldsIndex([]PivotTableField{}, &PivotTableOptions{})
	assert.EqualError(t, err, `parameter 'DataRange' parsing error: parameter is required`)
	// Test add pivot table with unsupported charset content types.
	f = NewFile()
	assert.NoError(t, f.SetSheetRow("Sheet1", "A1", &[]string{"Month", "Year", "Type", "Sales", "Region"}))
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, f.AddPivotTable(&PivotTableOptions{
		DataRange:       "Sheet1!A1:E31",
		PivotTableRange: "Sheet1!G2:M34",
		Rows:            []PivotTableField{{Data: "Year"}},
	}), "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())

	// Test get pivot table without pivot table
	f = NewFile()
	pivotTables, err = f.GetPivotTables("Sheet1")
	assert.NoError(t, err)
	assert.Len(t, pivotTables, 0)
	// Test get pivot table with not exists worksheet
	_, err = f.GetPivotTables("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Test get pivot table with unsupported charset worksheet relationships
	f.Pkg.Store("xl/worksheets/_rels/sheet1.xml.rels", MacintoshCyrillicCharset)
	_, err = f.GetPivotTables("Sheet1")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())
	// Test get pivot table with unsupported charset pivot cache definition
	f, err = OpenFile(filepath.Join("test", "TestAddPivotTable1.xlsx"))
	assert.NoError(t, err)
	f.Pkg.Store("xl/pivotCache/pivotCacheDefinition1.xml", MacintoshCyrillicCharset)
	_, err = f.GetPivotTables("Sheet1")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())
	// Test get pivot table with unsupported charset pivot table relationships
	f, err = OpenFile(filepath.Join("test", "TestAddPivotTable1.xlsx"))
	assert.NoError(t, err)
	f.Pkg.Store("xl/pivotTables/_rels/pivotTable1.xml.rels", MacintoshCyrillicCharset)
	_, err = f.GetPivotTables("Sheet1")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())
	// Test get pivot table with unsupported charset pivot table
	f, err = OpenFile(filepath.Join("test", "TestAddPivotTable1.xlsx"))
	assert.NoError(t, err)
	f.Pkg.Store("xl/pivotTables/pivotTable1.xml", MacintoshCyrillicCharset)
	_, err = f.GetPivotTables("Sheet1")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	_, err = f.getPivotTables()
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())
}

func TestPivotTableDataRange(t *testing.T) {
	f := NewFile()
	// Create table in a worksheet
	assert.NoError(t, f.AddTable("Sheet1", &Table{
		Name:  "Table1",
		Range: "A1:D5",
	}))
	for row := 2; row < 6; row++ {
		assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), rand.Intn(10)))
		assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("B%d", row), rand.Intn(10)))
		assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("C%d", row), rand.Intn(10)))
		assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("D%d", row), rand.Intn(10)))
	}
	// Test add pivot table with table data range
	opts := PivotTableOptions{
		DataRange:           "Table1",
		PivotTableRange:     "Sheet1!G2:K7",
		Rows:                []PivotTableField{{Data: "Column1"}},
		Columns:             []PivotTableField{{Data: "Column2"}},
		RowGrandTotals:      true,
		ColGrandTotals:      true,
		ShowDrill:           true,
		ShowRowHeaders:      true,
		ShowColHeaders:      true,
		ShowLastColumn:      true,
		ShowError:           true,
		PivotTableStyleName: "PivotStyleLight16",
	}
	assert.NoError(t, f.AddPivotTable(&opts))
	assert.NoError(t, f.DeletePivotTable("Sheet1", "PivotTable1"))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddPivotTable2.xlsx")))
	assert.NoError(t, f.Close())

	assert.NoError(t, f.AddPivotTable(&opts))

	// Test delete pivot table with unsupported table relationships charset
	f.Pkg.Store("xl/tables/table1.xml", MacintoshCyrillicCharset)
	assert.EqualError(t, f.DeletePivotTable("Sheet1", "PivotTable1"), "XML syntax error on line 1: invalid UTF-8")

	// Test delete pivot table with unsupported worksheet relationships charset
	f.Relationships.Delete("xl/worksheets/_rels/sheet1.xml.rels")
	f.Pkg.Store("xl/worksheets/_rels/sheet1.xml.rels", MacintoshCyrillicCharset)
	assert.EqualError(t, f.DeletePivotTable("Sheet1", "PivotTable1"), "XML syntax error on line 1: invalid UTF-8")

	// Test delete pivot table without worksheet relationships
	f.Relationships.Delete("xl/worksheets/_rels/sheet1.xml.rels")
	f.Pkg.Delete("xl/worksheets/_rels/sheet1.xml.rels")
	assert.EqualError(t, f.DeletePivotTable("Sheet1", "PivotTable1"), "table PivotTable1 does not exist")

	t.Run("data_range_with_empty_column", func(t *testing.T) {
		// Test add pivot table with data range doesn't organized as a list with labeled columns
		f := NewFile()
		// Create some data in a sheet
		month := []string{"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
		types := []string{"Meat", "Dairy", "Beverages", "Produce"}
		assert.NoError(t, f.SetSheetRow("Sheet1", "A1", &[]string{"Month", "", "Type"}))
		for row := 2; row < 32; row++ {
			assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), month[rand.Intn(12)]))
			assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("C%d", row), types[rand.Intn(4)]))
		}
		assert.Equal(t, newPivotTableDataRangeError("parameter is invalid"), f.AddPivotTable(&PivotTableOptions{
			DataRange:       "Sheet1!A1:E31",
			PivotTableRange: "Sheet1!G2:M34",
			Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}},
			Data:            []PivotTableField{{Data: "Type"}},
		}))
	})
}

func TestParseFormatPivotTableSet(t *testing.T) {
	f := NewFile()
	// Create table in a worksheet
	assert.NoError(t, f.AddTable("Sheet1", &Table{
		Name:  "Table1",
		Range: "A1:D5",
	}))
	// Test parse format pivot table options with unsupported table relationships charset
	f.Pkg.Store("xl/tables/table1.xml", MacintoshCyrillicCharset)
	_, _, err := f.parseFormatPivotTableSet(&PivotTableOptions{
		DataRange:       "Table1",
		PivotTableRange: "Sheet1!G2:K7",
		Rows:            []PivotTableField{{Data: "Column1"}},
	})
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
}

func TestAddPivotRowFields(t *testing.T) {
	f := NewFile()
	// Test invalid data range
	assert.EqualError(t, f.addPivotRowFields(&xlsxPivotTableDefinition{}, &PivotTableOptions{
		DataRange: "Sheet1!A1:A1",
	}), `parameter 'DataRange' parsing error: parameter is invalid`)
}

func TestAddPivotPageFields(t *testing.T) {
	f := NewFile()
	// Test invalid data range
	assert.EqualError(t, f.addPivotPageFields(&xlsxPivotTableDefinition{}, &PivotTableOptions{
		DataRange: "Sheet1!A1:A1",
	}), `parameter 'DataRange' parsing error: parameter is invalid`)
}

func TestAddPivotDataFields(t *testing.T) {
	f := NewFile()
	// Test invalid data range
	assert.EqualError(t, f.addPivotDataFields(&xlsxPivotTableDefinition{}, &PivotTableOptions{
		DataRange: "Sheet1!A1:A1",
	}), `parameter 'DataRange' parsing error: parameter is invalid`)
}

func TestAddPivotColFields(t *testing.T) {
	f := NewFile()
	// Test invalid data range
	assert.EqualError(t, f.addPivotColFields(&xlsxPivotTableDefinition{}, &PivotTableOptions{
		DataRange: "Sheet1!A1:A1",
		Columns:   []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
	}), `parameter 'DataRange' parsing error: parameter is invalid`)
}

func TestGetPivotFieldsOrder(t *testing.T) {
	f := NewFile()
	// Test get table fields order with not exist worksheet
	_, err := f.getTableFieldsOrder(&PivotTableOptions{DataRange: "SheetN!A1:E31"})
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Create table in a worksheet
	assert.NoError(t, f.AddTable("Sheet1", &Table{
		Name:  "Table1",
		Range: "A1:D5",
	}))
	// Test get table fields order with unsupported table relationships charset
	f.Pkg.Store("xl/tables/table1.xml", MacintoshCyrillicCharset)
	_, err = f.getTableFieldsOrder(&PivotTableOptions{DataRange: "Table"})
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
}

func TestGetPivotTableFieldName(t *testing.T) {
	f := NewFile()
	assert.Empty(t, f.getPivotTableFieldName("-", []PivotTableField{}))
}

func TestGetPivotTableFieldOptions(t *testing.T) {
	f := NewFile()
	_, ok := f.getPivotTableFieldOptions("-", []PivotTableField{})
	assert.False(t, ok)
}

func TestGenPivotCacheDefinitionID(t *testing.T) {
	f := NewFile()
	// Test generate pivot table cache definition ID with unsupported charset
	f.Pkg.Store("xl/pivotCache/pivotCacheDefinition1.xml", MacintoshCyrillicCharset)
	assert.Equal(t, 1, f.genPivotCacheDefinitionID())
	assert.NoError(t, f.Close())
}

func TestDeleteWorkbookPivotCache(t *testing.T) {
	f := NewFile()
	// Test delete workbook pivot table cache with unsupported workbook charset
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	assert.EqualError(t, f.deleteWorkbookPivotCache(PivotTableOptions{pivotCacheXML: "pivotCache/pivotCacheDefinition1.xml"}), "XML syntax error on line 1: invalid UTF-8")

	// Test delete workbook pivot table cache with unsupported workbook relationships charset
	f.Relationships.Delete("xl/_rels/workbook.xml.rels")
	f.Pkg.Store("xl/_rels/workbook.xml.rels", MacintoshCyrillicCharset)
	assert.EqualError(t, f.deleteWorkbookPivotCache(PivotTableOptions{pivotCacheXML: "pivotCache/pivotCacheDefinition1.xml"}), "XML syntax error on line 1: invalid UTF-8")
}
