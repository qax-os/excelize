package excelize

import (
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestGetMergeCells(t *testing.T) {
	wants := []struct {
		value string
		start string
		end   string
	}{{
		value: "A1",
		start: "A1",
		end:   "B1",
	}, {
		value: "A2",
		start: "A2",
		end:   "A3",
	}, {
		value: "A4",
		start: "A4",
		end:   "B5",
	}, {
		value: "A7",
		start: "A7",
		end:   "C10",
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
		assert.Equal(t, wants[i].value, m.GetCellValue())
		assert.Equal(t, wants[i].start, m.GetStartAxis())
		assert.Equal(t, wants[i].end, m.GetEndAxis())
	}

	// Test get merged cells on not exists worksheet.
	_, err = f.GetMergeCells("SheetN")
	assert.EqualError(t, err, "sheet SheetN is not exist")
}
