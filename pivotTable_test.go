package excelize

import (
	"fmt"
	"math/rand"
	"path/filepath"
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
	f.SetSheetRow("Sheet1", "A1", &[]string{"Month", "Year", "Type", "Sales", "Region"})
	for i := 0; i < 30; i++ {
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i+2), month[rand.Intn(12)])
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", i+2), year[rand.Intn(3)])
		f.SetCellValue("Sheet1", fmt.Sprintf("C%d", i+2), types[rand.Intn(4)])
		f.SetCellValue("Sheet1", fmt.Sprintf("D%d", i+2), rand.Intn(5000))
		f.SetCellValue("Sheet1", fmt.Sprintf("E%d", i+2), region[rand.Intn(4)])
	}
	assert.NoError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$1:$E$31",
		PivotTableRange: "Sheet1!$G$2:$M$34",
		Rows:            []string{"Month", "Year"},
		Columns:         []string{"Type"},
		Data:            []string{"Sales"},
	}))
	// Use different order of coordinate tests
	assert.NoError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$1:$E$31",
		PivotTableRange: "Sheet1!$U$34:$O$2",
		Rows:            []string{"Month", "Year"},
		Columns:         []string{"Type"},
		Data:            []string{"Sales"},
	}))

	assert.NoError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$1:$E$31",
		PivotTableRange: "Sheet1!$W$2:$AC$34",
		Rows:            []string{"Month", "Year"},
		Columns:         []string{"Region"},
		Data:            []string{"Sales"},
	}))
	assert.NoError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$1:$E$31",
		PivotTableRange: "Sheet1!$G$37:$W$50",
		Rows:            []string{"Month"},
		Columns:         []string{"Region", "Year"},
		Data:            []string{"Sales"},
	}))
	f.NewSheet("Sheet2")
	assert.NoError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$1:$E$31",
		PivotTableRange: "Sheet2!$A$1:$AR$15",
		Rows:            []string{"Month"},
		Columns:         []string{"Region", "Type", "Year"},
		Data:            []string{"Sales"},
	}))
	assert.NoError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$1:$E$31",
		PivotTableRange: "Sheet2!$A$18:$AR$54",
		Rows:            []string{"Month", "Type"},
		Columns:         []string{"Region", "Year"},
		Data:            []string{"Sales"},
	}))

	// Test empty pivot table options
	assert.EqualError(t, f.AddPivotTable(nil), "parameter is required")
	// Test invalid data range
	assert.EqualError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$1:$A$1",
		PivotTableRange: "Sheet1!$U$34:$O$2",
		Rows:            []string{"Month", "Year"},
		Columns:         []string{"Type"},
		Data:            []string{"Sales"},
	}), `parameter 'DataRange' parsing error: parameter is invalid`)
	// Test the data range of the worksheet that is not declared
	assert.EqualError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "$A$1:$E$31",
		PivotTableRange: "Sheet1!$U$34:$O$2",
		Rows:            []string{"Month", "Year"},
		Columns:         []string{"Type"},
		Data:            []string{"Sales"},
	}), `parameter 'DataRange' parsing error: parameter is invalid`)
	// Test the worksheet declared in the data range does not exist
	assert.EqualError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "SheetN!$A$1:$E$31",
		PivotTableRange: "Sheet1!$U$34:$O$2",
		Rows:            []string{"Month", "Year"},
		Columns:         []string{"Type"},
		Data:            []string{"Sales"},
	}), "sheet SheetN is not exist")
	// Test the pivot table range of the worksheet that is not declared
	assert.EqualError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$1:$E$31",
		PivotTableRange: "$U$34:$O$2",
		Rows:            []string{"Month", "Year"},
		Columns:         []string{"Type"},
		Data:            []string{"Sales"},
	}), `parameter 'PivotTableRange' parsing error: parameter is invalid`)
	// Test the worksheet declared in the pivot table range does not exist
	assert.EqualError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$1:$E$31",
		PivotTableRange: "SheetN!$U$34:$O$2",
		Rows:            []string{"Month", "Year"},
		Columns:         []string{"Type"},
		Data:            []string{"Sales"},
	}), "sheet SheetN is not exist")
	// Test not exists worksheet in data range
	assert.EqualError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "SheetN!$A$1:$E$31",
		PivotTableRange: "Sheet1!$U$34:$O$2",
		Rows:            []string{"Month", "Year"},
		Columns:         []string{"Type"},
		Data:            []string{"Sales"},
	}), "sheet SheetN is not exist")
	// Test invalid row number in data range
	assert.EqualError(t, f.AddPivotTable(&PivotTableOption{
		DataRange:       "Sheet1!$A$0:$E$31",
		PivotTableRange: "Sheet1!$U$34:$O$2",
		Rows:            []string{"Month", "Year"},
		Columns:         []string{"Type"},
		Data:            []string{"Sales"},
	}), `parameter 'DataRange' parsing error: cannot convert cell "A0" to coordinates: invalid cell name "A0"`)
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddPivotTable1.xlsx")))

	// Test adjust range with invalid range
	_, _, err := f.adjustRange("")
	assert.EqualError(t, err, "parameter is required")
	// Test get pivot fields order with empty data range
	_, err = f.getPivotFieldsOrder("")
	assert.EqualError(t, err, `parameter 'DataRange' parsing error: parameter is required`)
	// Test add pivot cache with empty data range
	assert.EqualError(t, f.addPivotCache(0, "", &PivotTableOption{}, nil), "parameter 'DataRange' parsing error: parameter is required")
	// Test add pivot cache with invalid data range
	assert.EqualError(t, f.addPivotCache(0, "", &PivotTableOption{
		DataRange:       "$A$1:$E$31",
		PivotTableRange: "Sheet1!$U$34:$O$2",
		Rows:            []string{"Month", "Year"},
		Columns:         []string{"Type"},
		Data:            []string{"Sales"},
	}, nil), "parameter 'DataRange' parsing error: parameter is invalid")
	// Test add pivot table with empty options
	assert.EqualError(t, f.addPivotTable(0, 0, "", &PivotTableOption{}), "parameter 'PivotTableRange' parsing error: parameter is required")
	// Test add pivot table with invalid data range
	assert.EqualError(t, f.addPivotTable(0, 0, "", &PivotTableOption{}), "parameter 'PivotTableRange' parsing error: parameter is required")
	// Test add pivot fields with empty data range
	assert.EqualError(t, f.addPivotFields(nil, &PivotTableOption{
		DataRange:       "$A$1:$E$31",
		PivotTableRange: "Sheet1!$U$34:$O$2",
		Rows:            []string{"Month", "Year"},
		Columns:         []string{"Type"},
		Data:            []string{"Sales"},
	}), `parameter 'DataRange' parsing error: parameter is invalid`)
	// Test get pivot fields index with empty data range
	_, err = f.getPivotFieldsIndex([]string{}, &PivotTableOption{})
	assert.EqualError(t, err, `parameter 'DataRange' parsing error: parameter is required`)
}
