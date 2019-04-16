package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestAdjustMergeCells(t *testing.T) {
	f := NewFile()
	// testing adjustAutoFilter with illegal cell coordinates.
	assert.EqualError(t, f.adjustMergeCells(&xlsxWorksheet{
		MergeCells: &xlsxMergeCells{
			Cells: []*xlsxMergeCell{
				{
					Ref: "A:B1",
				},
			},
		},
	}, rows, 0, 0), `cannot convert cell "A" to coordinates: invalid cell name "A"`)
	assert.EqualError(t, f.adjustMergeCells(&xlsxWorksheet{
		MergeCells: &xlsxMergeCells{
			Cells: []*xlsxMergeCell{
				{
					Ref: "A1:B",
				},
			},
		},
	}, rows, 0, 0), `cannot convert cell "B" to coordinates: invalid cell name "B"`)
}

func TestAdjustAutoFilter(t *testing.T) {
	f := NewFile()
	// testing adjustAutoFilter with illegal cell coordinates.
	assert.EqualError(t, f.adjustAutoFilter(&xlsxWorksheet{
		AutoFilter: &xlsxAutoFilter{
			Ref: "A:B1",
		},
	}, rows, 0, 0), `cannot convert cell "A" to coordinates: invalid cell name "A"`)
	assert.EqualError(t, f.adjustAutoFilter(&xlsxWorksheet{
		AutoFilter: &xlsxAutoFilter{
			Ref: "A1:B",
		},
	}, rows, 0, 0), `cannot convert cell "B" to coordinates: invalid cell name "B"`)
}

func TestAdjustHelper(t *testing.T) {
	f := NewFile()
	f.NewSheet("Sheet2")
	f.Sheet["xl/worksheets/sheet1.xml"] = &xlsxWorksheet{
		MergeCells: &xlsxMergeCells{
			Cells: []*xlsxMergeCell{
				{
					Ref: "A:B1",
				},
			},
		},
	}
	f.Sheet["xl/worksheets/sheet2.xml"] = &xlsxWorksheet{
		AutoFilter: &xlsxAutoFilter{
			Ref: "A1:B",
		},
	}
	// testing adjustHelper with illegal cell coordinates.
	assert.EqualError(t, f.adjustHelper("Sheet1", rows, 0, 0), `cannot convert cell "A" to coordinates: invalid cell name "A"`)
	assert.EqualError(t, f.adjustHelper("Sheet2", rows, 0, 0), `cannot convert cell "B" to coordinates: invalid cell name "B"`)
	// testing adjustHelper on not exists worksheet.
	assert.EqualError(t, f.adjustHelper("SheetN", rows, 0, 0), "sheet SheetN is not exist")
}
