package excelize

import (
	"fmt"
	"math/rand"
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestAddPivotTable(t *testing.T) {
	f := NewFile()
	// Create some data in a sheet
	month := []string{"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
	year := []int{2017, 2018, 2019}
	types := []string{"Meat", "Dairy", "Beverages", "Produce"}
	region := []string{"East", "West", "North", "South"}
	assert.NoError(t, f.SetSheetRow("Sheet1", "A1", &[]string{"Month", "Year", "Type", "Sales", "Region"}))
	for i := 0; i < 30; i++ {
		assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i+2), month[rand.Intn(12)]))
		assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("B%d", i+2), year[rand.Intn(3)]))
		assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("C%d", i+2), types[rand.Intn(4)]))
		assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("D%d", i+2), rand.Intn(5000)))
		assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("E%d", i+2), region[rand.Intn(4)]))
	}
	assert.NoError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$1:$E$31",
		PivotTableRange: "Sheet1!$G$2:$M$34",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Filter:          []PivotTableField{{Data: "Region"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales", Subtotal: "Sum", Name: "Summarize by Sum"}},
		RowGrandTotals:  true,
		ColGrandTotals:  true,
		ShowDrill:       true,
		ShowRowHeaders:  true,
		ShowColHeaders:  true,
		ShowLastColumn:  true,
		ShowError:       true,
	}))
	// Use different order of coordinate tests
	assert.NoError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$1:$E$31",
		PivotTableRange: "Sheet1!$U$34:$O$2",
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

	assert.NoError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$1:$E$31",
		PivotTableRange: "Sheet1!$W$2:$AC$34",
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
	assert.NoError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$1:$E$31",
		PivotTableRange: "Sheet1!$G$37:$W$50",
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
	assert.NoError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$1:$E$31",
		PivotTableRange: "Sheet1!$AE$2:$AG$33",
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
	assert.NoError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:           "Sheet1!$A$1:$E$31",
		PivotTableRange:     "Sheet1!$AJ$2:$AP1$35",
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
	f.NewSheet("Sheet2")
	assert.NoError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$1:$E$31",
		PivotTableRange: "Sheet2!$A$1:$AR$15",
		Rows:            []PivotTableField{{Data: "Month"}},
		Columns:         []PivotTableField{{Data: "Region", DefaultSubtotal: true}, {Data: "Type", DefaultSubtotal: true}, {Data: "Year"}},
		Data:            []PivotTableField{{Data: "Sales", Subtotal: "Min", Name: "Summarize by Min"}},
		RowGrandTotals:  true,
		ColGrandTotals:  true,
		ShowDrill:       true,
		ShowRowHeaders:  true,
		ShowColHeaders:  true,
		ShowLastColumn:  true,
	}))
	assert.NoError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$1:$E$31",
		PivotTableRange: "Sheet2!$A$18:$AR$54",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Type"}},
		Columns:         []PivotTableField{{Data: "Region", DefaultSubtotal: true}, {Data: "Year"}},
		Data:            []PivotTableField{{Data: "Sales", Subtotal: "Product", Name: "Summarize by Product"}},
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
		RefersTo: "Sheet1!$A$1:$E$31",
		Comment:  "Pivot Table Data Range",
		Scope:    "Sheet2",
	}))
	assert.NoError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "dataRange",
		PivotTableRange: "Sheet2!$A$57:$AJ$91",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Region", DefaultSubtotal: true}, {Data: "Type"}},
		Data:            []PivotTableField{{Data: "Sales", Subtotal: "Sum", Name: "Sum of Sales"}, {Data: "Sales", Subtotal: "Average", Name: "Average of Sales"}},
		RowGrandTotals:  true,
		ColGrandTotals:  true,
		ShowDrill:       true,
		ShowRowHeaders:  true,
		ShowColHeaders:  true,
		ShowLastColumn:  true,
	}))

	// Test empty pivot table options
	assert.EqualError(t, f.AddPivotTable(nil), ErrParameterRequired.Error())
	// Test invalid data range
	assert.EqualError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$1:$A$1",
		PivotTableRange: "Sheet1!$U$34:$O$2",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales"}},
	}), `parameter 'DataRange' parsing error: parameter is invalid`)
	// Test the data range of the worksheet that is not declared
	assert.EqualError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "$A$1:$E$31",
		PivotTableRange: "Sheet1!$U$34:$O$2",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales"}},
	}), `parameter 'DataRange' parsing error: parameter is invalid`)
	// Test the worksheet declared in the data range does not exist
	assert.EqualError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "SheetN!$A$1:$E$31",
		PivotTableRange: "Sheet1!$U$34:$O$2",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales"}},
	}), "sheet SheetN is not exist")
	// Test the pivot table range of the worksheet that is not declared
	assert.EqualError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$1:$E$31",
		PivotTableRange: "$U$34:$O$2",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales"}},
	}), `parameter 'PivotTableRange' parsing error: parameter is invalid`)
	// Test the worksheet declared in the pivot table range does not exist
	assert.EqualError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$1:$E$31",
		PivotTableRange: "SheetN!$U$34:$O$2",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales"}},
	}), "sheet SheetN is not exist")
	// Test not exists worksheet in data range
	assert.EqualError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "SheetN!$A$1:$E$31",
		PivotTableRange: "Sheet1!$U$34:$O$2",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales"}},
	}), "sheet SheetN is not exist")
	// Test invalid row number in data range
	assert.EqualError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$0:$E$31",
		PivotTableRange: "Sheet1!$U$34:$O$2",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales"}},
	}), `parameter 'DataRange' parsing error: cannot convert cell "A0" to coordinates: invalid cell name "A0"`)
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddPivotTable1.xlsx")))
	// Test with field names that exceed the length limit and invalid subtotal
	assert.NoError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$1:$E$31",
		PivotTableRange: "Sheet1!$G$2:$M$34",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales", Subtotal: "-", Name: strings.Repeat("s", MaxFieldLength+1)}},
	}))

	// Test adjust range with invalid range
	_, _, err := f.adjustRange("")
	assert.EqualError(t, err, ErrParameterRequired.Error())
	// Test adjust range with incorrect range
	_, _, err = f.adjustRange("sheet1!")
	assert.EqualError(t, err, "parameter is invalid")
	// Test get pivot fields order with empty data range
	_, err = f.getPivotFieldsOrder(&PivotTableOption{})
	assert.EqualError(t, err, `parameter 'DataRange' parsing error: parameter is required`)
	// Test add pivot cache with empty data range
	assert.EqualError(t, f.addPivotCache(0, "", &PivotTableOption{}, nil), "parameter 'DataRange' parsing error: parameter is required")
	// Test add pivot cache with invalid data range
	assert.EqualError(t, f.addPivotCache(0, "", &PivotTableOption{
		DataRange:       "$A$1:$E$31",
		PivotTableRange: "Sheet1!$U$34:$O$2",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales"}},
	}, nil), "parameter 'DataRange' parsing error: parameter is invalid")
	// Test add pivot table with empty options
	assert.EqualError(t, f.addPivotTable(0, 0, "", &PivotTableOption{}), "parameter 'PivotTableRange' parsing error: parameter is required")
	// Test add pivot table with invalid data range
	assert.EqualError(t, f.addPivotTable(0, 0, "", &PivotTableOption{}), "parameter 'PivotTableRange' parsing error: parameter is required")
	// Test add pivot fields with empty data range
	assert.EqualError(t, f.addPivotFields(nil, &PivotTableOption{
		DataRange:       "$A$1:$E$31",
		PivotTableRange: "Sheet1!$U$34:$O$2",
		Rows:            []PivotTableField{{Data: "Month", DefaultSubtotal: true}, {Data: "Year"}},
		Columns:         []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
		Data:            []PivotTableField{{Data: "Sales"}},
	}), `parameter 'DataRange' parsing error: parameter is invalid`)
	// Test get pivot fields index with empty data range
	_, err = f.getPivotFieldsIndex([]PivotTableField{}, &PivotTableOption{})
	assert.EqualError(t, err, `parameter 'DataRange' parsing error: parameter is required`)
}

func TestAddPivotRowFields(t *testing.T) {
	f := NewFile()
	// Test invalid data range
	assert.EqualError(t, f.addPivotRowFields(&xlsxPivotTableDefinition{}, &PivotTableOption{
		DataRange: "Sheet1!$A$1:$A$1",
	}), `parameter 'DataRange' parsing error: parameter is invalid`)
}

func TestAddPivotPageFields(t *testing.T) {
	f := NewFile()
	// Test invalid data range
	assert.EqualError(t, f.addPivotPageFields(&xlsxPivotTableDefinition{}, &PivotTableOption{
		DataRange: "Sheet1!$A$1:$A$1",
	}), `parameter 'DataRange' parsing error: parameter is invalid`)
}

func TestAddPivotDataFields(t *testing.T) {
	f := NewFile()
	// Test invalid data range
	assert.EqualError(t, f.addPivotDataFields(&xlsxPivotTableDefinition{}, &PivotTableOption{
		DataRange: "Sheet1!$A$1:$A$1",
	}), `parameter 'DataRange' parsing error: parameter is invalid`)
}

func TestAddPivotColFields(t *testing.T) {
	f := NewFile()
	// Test invalid data range
	assert.EqualError(t, f.addPivotColFields(&xlsxPivotTableDefinition{}, &PivotTableOption{
		DataRange: "Sheet1!$A$1:$A$1",
		Columns:   []PivotTableField{{Data: "Type", DefaultSubtotal: true}},
	}), `parameter 'DataRange' parsing error: parameter is invalid`)
}

func TestGetPivotFieldsOrder(t *testing.T) {
	f := NewFile()
	// Test get pivot fields order with not exist worksheet
	_, err := f.getPivotFieldsOrder(&PivotTableOption{DataRange: "SheetN!$A$1:$E$31"})
	assert.EqualError(t, err, "sheet SheetN is not exist")
}

func TestGetPivotTableFieldName(t *testing.T) {
	f := NewFile()
	f.getPivotTableFieldName("-", []PivotTableField{})
}
