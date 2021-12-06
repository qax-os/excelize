package excelize

import (
	"fmt"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestAddTable(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	err = f.AddTable("Sheet1", "B26", "A21", `{}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	err = f.AddTable("Sheet2", "A2", "B5", `{"table_name":"table","table_style":"TableStyleMedium2", "show_first_column":true,"show_last_column":true,"show_row_stripes":false,"show_column_stripes":true}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	err = f.AddTable("Sheet2", "F1", "F1", `{"table_style":"TableStyleMedium8"}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Test add table in not exist worksheet.
	assert.EqualError(t, f.AddTable("SheetN", "B26", "A21", `{}`), "sheet SheetN is not exist")
	// Test add table with illegal formatset.
	assert.EqualError(t, f.AddTable("Sheet1", "B26", "A21", `{x}`), "invalid character 'x' looking for beginning of object key string")
	// Test add table with illegal cell coordinates.
	assert.EqualError(t, f.AddTable("Sheet1", "A", "B1", `{}`), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	assert.EqualError(t, f.AddTable("Sheet1", "A1", "B", `{}`), newCellNameToCoordinatesError("B", newInvalidCellNameError("B")).Error())

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddTable.xlsx")))

	// Test addTable with illegal cell coordinates.
	f = NewFile()
	assert.EqualError(t, f.addTable("sheet1", "", 0, 0, 0, 0, 0, nil), "invalid cell coordinates [0, 0]")
	assert.EqualError(t, f.addTable("sheet1", "", 1, 1, 0, 0, 0, nil), "invalid cell coordinates [0, 0]")
}

func TestAutoFilter(t *testing.T) {
	outFile := filepath.Join("test", "TestAutoFilter%d.xlsx")

	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	formats := []string{
		``,
		`{"column":"B","expression":"x != blanks"}`,
		`{"column":"B","expression":"x == blanks"}`,
		`{"column":"B","expression":"x != nonblanks"}`,
		`{"column":"B","expression":"x == nonblanks"}`,
		`{"column":"B","expression":"x <= 1 and x >= 2"}`,
		`{"column":"B","expression":"x == 1 or x == 2"}`,
		`{"column":"B","expression":"x == 1 or x == 2*"}`,
	}

	for i, format := range formats {
		t.Run(fmt.Sprintf("Expression%d", i+1), func(t *testing.T) {
			err = f.AutoFilter("Sheet1", "D4", "B1", format)
			assert.NoError(t, err)
			assert.NoError(t, f.SaveAs(fmt.Sprintf(outFile, i+1)))
		})
	}

	// testing AutoFilter with illegal cell coordinates.
	assert.EqualError(t, f.AutoFilter("Sheet1", "A", "B1", ""), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	assert.EqualError(t, f.AutoFilter("Sheet1", "A1", "B", ""), newCellNameToCoordinatesError("B", newInvalidCellNameError("B")).Error())
}

func TestAutoFilterError(t *testing.T) {
	outFile := filepath.Join("test", "TestAutoFilterError%d.xlsx")

	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	formats := []string{
		`{"column":"B","expression":"x <= 1 and x >= blanks"}`,
		`{"column":"B","expression":"x -- y or x == *2*"}`,
		`{"column":"B","expression":"x != y or x ? *2"}`,
		`{"column":"B","expression":"x -- y o r x == *2"}`,
		`{"column":"B","expression":"x -- y"}`,
		`{"column":"A","expression":"x -- y"}`,
	}
	for i, format := range formats {
		t.Run(fmt.Sprintf("Expression%d", i+1), func(t *testing.T) {
			err = f.AutoFilter("Sheet2", "D4", "B1", format)
			if assert.Error(t, err) {
				assert.NoError(t, f.SaveAs(fmt.Sprintf(outFile, i+1)))
			}
		})
	}

	assert.EqualError(t, f.autoFilter("SheetN", "A1", 1, 1, &formatAutoFilter{
		Column:     "A",
		Expression: "",
	}), "sheet SheetN is not exist")
	assert.EqualError(t, f.autoFilter("Sheet1", "A1", 1, 1, &formatAutoFilter{
		Column:     "-",
		Expression: "-",
	}), newInvalidColumnNameError("-").Error())
	assert.EqualError(t, f.autoFilter("Sheet1", "A1", 1, 100, &formatAutoFilter{
		Column:     "A",
		Expression: "-",
	}), `incorrect index of column 'A'`)
	assert.EqualError(t, f.autoFilter("Sheet1", "A1", 1, 1, &formatAutoFilter{
		Column:     "A",
		Expression: "-",
	}), `incorrect number of tokens in criteria '-'`)
}

func TestParseFilterTokens(t *testing.T) {
	f := NewFile()
	// Test with unknown operator.
	_, _, err := f.parseFilterTokens("", []string{"", "!"})
	assert.EqualError(t, err, "unknown operator: !")
	// Test invalid operator in context.
	_, _, err = f.parseFilterTokens("", []string{"", "<", "x != blanks"})
	assert.EqualError(t, err, "the operator '<' in expression '' is not valid in relation to Blanks/NonBlanks'")
}
