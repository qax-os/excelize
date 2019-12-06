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

func TestUnmergeCell(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "MergeCell.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	sheet1 := f.GetSheetName(1)

	xlsx, err := f.workSheetReader(sheet1)
	assert.NoError(t, err)

	mergeCellNum := len(xlsx.MergeCells.Cells)

	assert.EqualError(t, f.UnmergeCell("Sheet1", "A", "A"), `cannot convert cell "A" to coordinates: invalid cell name "A"`)

	// unmerge the mergecell that contains A1
	err = f.UnmergeCell(sheet1, "A1", "A1")
	assert.NoError(t, err)

	if len(xlsx.MergeCells.Cells) != mergeCellNum-1 {
		t.FailNow()
	}

	// unmerge area A7:D3(A3:D7)
	// this will unmerge all since this area overlaps with all others
	err = f.UnmergeCell(sheet1, "D7", "A3")
	assert.NoError(t, err)

	if len(xlsx.MergeCells.Cells) != 0 {
		t.FailNow()
	}

	// Test unmerged area on not exists worksheet.
	err = f.UnmergeCell("SheetN", "A1", "A1")
	assert.EqualError(t, err, "sheet SheetN is not exist")
}
