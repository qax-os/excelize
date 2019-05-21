package excelize

import (
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestGetMergeCells(t *testing.T) {
	wants := []struct {
		axis  string
		value string
		start string
		end   string
		x     []string
		y     []string
	}{{
		axis:  "A1:B1",
		value: "A1",
		start: "A1",
		end:   "B1",
		x:     []string{"A1", "B1"},
		y:     []string{"A1"},
	}, {
		axis:  "A2:A3",
		value: "A2",
		start: "A2",
		end:   "A3",
		x:     []string{"A2"},
		y:     []string{"A2", "A3"},
	}, {
		axis:  "A4:B5",
		value: "A4",
		start: "A4",
		end:   "B5",
		x:     []string{"A4", "B4"},
		y:     []string{"A4", "A5"},
	}, {
		axis:  "A7:C10",
		value: "A7",
		start: "A7",
		end:   "C10",
		x:     []string{"A7", "B7", "C7"},
		y:     []string{"A7", "A8", "A9", "A10"},
	}}

	f, err := OpenFile(filepath.Join("test", "MergeCell.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	sheet1 := f.GetSheetName(1)

	mergeCells, err := f.GetMergeCells(sheet1)
	if !assert.Len(t, mergeCells, len(wants)) {
		t.FailNow()
	}
	assert.NoError(t, err)

	for i, m := range mergeCells {
		assert.Equal(t, wants[i].axis, m.GetCellAxis())
		assert.Equal(t, wants[i].value, m.GetCellValue())
		assert.Equal(t, wants[i].start, m.GetStartAxis())
		assert.Equal(t, wants[i].end, m.GetEndAxis())
	}

	for i := range wants {
		cellsX, cellsY, err := f.GetRangeCells(wants[i].axis)
		assert.NoError(t, err)
		assert.ElementsMatch(t, wants[i].x, cellsX)
		assert.ElementsMatch(t, wants[i].y, cellsY)
		assert.Equal(t, wants[i].axis, f.searchMergedCell(mergeCells, wants[i].start))
	}

	// Test get merged cells on not exists worksheet.
	_, err = f.GetMergeCells("SheetN")
	assert.EqualError(t, err, "sheet SheetN is not exist")
}
