package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestAddButton(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.AddButton("Sheet1", "A30", `{"macro": "say_hello","caption": "Press Me","width": 80,"height": 30}`))

	// Test add button on not exists worksheet.
	assert.EqualError(t, f.AddButton("SheetN", "B7", `{"macro": "say_hello","caption": "Press Me","width": 80,"height": 30}`), "sheet SheetN is not exist")
	// Test add button on with illegal cell coordinates
	assert.EqualError(t, f.AddButton("Sheet1", "A", `{"macro": "say_hello","caption": "Press Me","width": 80,"height": 30}`), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
}
