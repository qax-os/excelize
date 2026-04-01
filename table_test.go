package excelize

import (
	"fmt"
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestAddTable(t *testing.T) {
	f, err := prepareTestBook1()
	assert.NoError(t, err)
	assert.NoError(t, f.AddTable("Sheet1", &Table{Range: "B26:A21"}))
	assert.NoError(t, f.AddTable("Sheet2", &Table{
		Range:             "A2:B5",
		Name:              "table",
		StyleName:         "TableStyleMedium2",
		ShowColumnStripes: true,
		ShowFirstColumn:   true,
		ShowLastColumn:    true,
		ShowRowStripes:    boolPtr(true),
	}))
	assert.NoError(t, f.AddTable("Sheet2", &Table{
		Range:         "D1:D11",
		ShowHeaderRow: boolPtr(false),
	}))
	assert.NoError(t, f.AddTable("Sheet2", &Table{Range: "F1:F1", StyleName: "TableStyleMedium8"}))
	// Test get tables in worksheet
	tables, err := f.GetTables("Sheet2")
	assert.Len(t, tables, 3)
	assert.NoError(t, err)

	// Test add table with already exist table name
	assert.Equal(t, f.AddTable("Sheet2", &Table{Name: "Table1"}), ErrExistsTableName)
	// Test add table with invalid table options
	assert.Equal(t, f.AddTable("Sheet1", nil), ErrParameterInvalid)
	// Test add table in not exist worksheet
	assert.EqualError(t, f.AddTable("SheetN", &Table{Range: "B26:A21"}), "sheet SheetN does not exist")
	// Test add table with illegal cell reference
	assert.Equal(t, f.AddTable("Sheet1", &Table{Range: "A:B1"}), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")))
	assert.Equal(t, f.AddTable("Sheet1", &Table{Range: "A1:B"}), newCellNameToCoordinatesError("B", newInvalidCellNameError("B")))

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddTable.xlsx")))

	// Test add table with invalid sheet name
	assert.Equal(t, ErrSheetNameInvalid, f.AddTable("Sheet:1", &Table{Range: "B26:A21"}))
	// Test addTable with illegal cell reference
	f = NewFile()
	assert.Equal(t, newCoordinatesToCellNameError(0, 0), f.addTable("sheet1", "", 0, 0, 0, 0, 0, nil))
	assert.Equal(t, newCoordinatesToCellNameError(0, 0), f.addTable("sheet1", "", 1, 1, 0, 0, 0, nil))
	// Test set defined name and add table with invalid name
	for _, cases := range []struct {
		name string
		err  error
	}{
		{name: "1Table", err: newInvalidNameError("1Table")},
		{name: "-Table", err: newInvalidNameError("-Table")},
		{name: "'Table", err: newInvalidNameError("'Table")},
		{name: "Table 1", err: newInvalidNameError("Table 1")},
		{name: "A&B", err: newInvalidNameError("A&B")},
		{name: "_1Table'", err: newInvalidNameError("_1Table'")},
		{name: "\u0f5f\u0fb3\u0f0b\u0f21", err: newInvalidNameError("\u0f5f\u0fb3\u0f0b\u0f21")},
		{name: strings.Repeat("c", MaxFieldLength+1), err: ErrNameLength},
	} {
		assert.Equal(t, cases.err, f.AddTable("Sheet1", &Table{
			Range: "A1:B2",
			Name:  cases.name,
		}))
		assert.Equal(t, cases.err, f.SetDefinedName(&DefinedName{
			Name: cases.name, RefersTo: "Sheet1!$A$2:$D$5",
		}))
	}
	// Test check duplicate table name with unsupported charset table parts
	f = NewFile()
	f.Pkg.Store("xl/tables/table1.xml", MacintoshCyrillicCharset)
	assert.NoError(t, f.AddTable("Sheet1", &Table{Range: "A1:B2"}))
	assert.NoError(t, f.Close())
	f = NewFile()
	// Test add table with workbook with single cells parts
	f.Pkg.Store("xl/tables/tableSingleCells1.xml", []byte("<singleXmlCells><singleXmlCell id=\"2\" r=\"A1\" connectionId=\"2\" /></singleXmlCells>"))
	assert.NoError(t, f.AddTable("Sheet1", &Table{Range: "A1:B2"}))
	// Test add table with workbook with unsupported charset single cells parts
	f.Pkg.Store("xl/tables/tableSingleCells1.xml", MacintoshCyrillicCharset)
	assert.NoError(t, f.AddTable("Sheet1", &Table{Range: "A1:B2"}))
	assert.NoError(t, f.Close())
}

func TestGetTables(t *testing.T) {
	f := NewFile()
	// Test get tables in none table worksheet
	tables, err := f.GetTables("Sheet1")
	assert.Len(t, tables, 0)
	assert.NoError(t, err)
	// Test get tables in not exist worksheet
	_, err = f.GetTables("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Test adjust table with unsupported charset
	assert.NoError(t, f.AddTable("Sheet1", &Table{Range: "B26:A21"}))
	f.Pkg.Store("xl/tables/table1.xml", MacintoshCyrillicCharset)
	_, err = f.GetTables("Sheet1")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	// Test adjust table with no exist table parts
	f.Pkg.Delete("xl/tables/table1.xml")
	tables, err = f.GetTables("Sheet1")
	assert.Len(t, tables, 0)
	assert.NoError(t, err)
}

func TestDeleteTable(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.AddTable("Sheet1", &Table{Range: "A1:B4", Name: "Table1"}))
	assert.NoError(t, f.AddTable("Sheet1", &Table{Range: "B26:A21", Name: "Table2"}))
	assert.NoError(t, f.DeleteTable("Table2"))
	assert.NoError(t, f.DeleteTable("Table1"))
	// Test delete table with invalid table name
	assert.Equal(t, newInvalidNameError("Table 1"), f.DeleteTable("Table 1"))
	// Test delete table with no exist table name
	assert.Equal(t, newNoExistTableError("Table"), f.DeleteTable("Table"))
	// Test delete table with unsupported charset
	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", MacintoshCyrillicCharset)
	assert.EqualError(t, f.DeleteTable("Table1"), "XML syntax error on line 1: invalid UTF-8")
	// Test delete table without deleting table header
	f = NewFile()
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "Date"))
	assert.NoError(t, f.SetCellValue("Sheet1", "B1", "Values"))
	assert.NoError(t, f.UpdateLinkedValue())
	assert.NoError(t, f.AddTable("Sheet1", &Table{Range: "A1:B2", Name: "Table1"}))
	assert.NoError(t, f.DeleteTable("Table1"))
	val, err := f.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "Date", val)
	val, err = f.GetCellValue("Sheet1", "B1")
	assert.NoError(t, err)
	assert.Equal(t, "Values", val)
}

func TestSetTableColumns(t *testing.T) {
	f := NewFile()
	assert.Equal(t, newCoordinatesToCellNameError(1, 0), f.setTableColumns("Sheet1", true, 1, 0, 1, nil))
}

func TestAutoFilter(t *testing.T) {
	outFile := filepath.Join("test", "TestAutoFilter%d.xlsx")
	f, err := prepareTestBook1()
	assert.NoError(t, err)
	for i, opts := range [][]AutoFilterOptions{
		{},
		{{Column: "B", Expression: ""}},
		{{Column: "B", Expression: "x != blanks"}},
		{{Column: "B", Expression: "x == blanks"}},
		{{Column: "B", Expression: "x != nonblanks"}},
		{{Column: "B", Expression: "x == nonblanks"}},
		{{Column: "B", Expression: "x <= 1 and x >= 2"}},
		{{Column: "B", Expression: "x == 1 or x == 2"}},
		{{Column: "B", Expression: "x == 1 or x == 2*"}},
	} {
		t.Run(fmt.Sprintf("Expression%d", i+1), func(t *testing.T) {
			assert.NoError(t, f.AutoFilter("Sheet1", "D4:B1", opts))
			assert.NoError(t, f.SaveAs(fmt.Sprintf(outFile, i+1)))
		})
	}

	// Test add auto filter with invalid sheet name
	assert.Equal(t, ErrSheetNameInvalid, f.AutoFilter("Sheet:1", "A1:B1", nil))
	// Test add auto filter with illegal cell reference
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), f.AutoFilter("Sheet1", "A:B1", nil))
	assert.Equal(t, newCellNameToCoordinatesError("B", newInvalidCellNameError("B")), f.AutoFilter("Sheet1", "A1:B", nil))
	// Test add auto filter with unsupported charset workbook
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	assert.EqualError(t, f.AutoFilter("Sheet1", "D4:B1", nil), "XML syntax error on line 1: invalid UTF-8")
	// Test add auto filter with empty local sheet ID
	f = NewFile()
	f.WorkBook = &xlsxWorkbook{DefinedNames: &xlsxDefinedNames{DefinedName: []xlsxDefinedName{{Name: builtInDefinedNames[3], Hidden: true}}}}
	assert.NoError(t, f.AutoFilter("Sheet1", "A1:B1", nil))
}

func TestAutoFilterError(t *testing.T) {
	outFile := filepath.Join("test", "TestAutoFilterError%d.xlsx")
	f, err := prepareTestBook1()
	assert.NoError(t, err)
	for i, opts := range [][]AutoFilterOptions{
		{{Column: "B", Expression: "x <= 1 and x >= blanks"}},
		{{Column: "B", Expression: "x -- y or x == *2*"}},
		{{Column: "B", Expression: "x != y or x ? *2"}},
		{{Column: "B", Expression: "x -- y o r x == *2"}},
		{{Column: "B", Expression: "x -- y"}},
		{{Column: "A", Expression: "x -- y"}},
	} {
		t.Run(fmt.Sprintf("Expression%d", i+1), func(t *testing.T) {
			if assert.Error(t, f.AutoFilter("Sheet2", "D4:B1", opts)) {
				assert.NoError(t, f.SaveAs(fmt.Sprintf(outFile, i+1)))
			}
		})
	}

	assert.Equal(t, ErrSheetNotExist{"SheetN"}, f.autoFilter("SheetN", "A1", 1, 1, []AutoFilterOptions{{
		Column:     "A",
		Expression: "",
	}}))
	assert.Equal(t, newInvalidColumnNameError("-"), f.autoFilter("Sheet1", "A1", 1, 1, []AutoFilterOptions{{
		Column:     "-",
		Expression: "-",
	}}))
	assert.Equal(t, newInvalidAutoFilterColumnError("A"), f.autoFilter("Sheet1", "A1", 1, 100, []AutoFilterOptions{{
		Column:     "A",
		Expression: "-",
	}}))
	assert.Equal(t, newInvalidAutoFilterExpError("-"), f.autoFilter("Sheet1", "A1", 1, 1, []AutoFilterOptions{{
		Column:     "A",
		Expression: "-",
	}}))
}

func TestParseFilterTokens(t *testing.T) {
	f := NewFile()
	// Test with unknown operator
	_, _, err := f.parseFilterTokens("", []string{"", "!"})
	assert.EqualError(t, err, "unknown operator: !")
	// Test invalid operator in context
	_, _, err = f.parseFilterTokens("", []string{"", "<", "x != blanks"})
	assert.Equal(t, newInvalidAutoFilterOperatorError("<", ""), err)
}
