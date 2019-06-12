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
	assert.NoError(t, f.adjustMergeCells(&xlsxWorksheet{
		MergeCells: &xlsxMergeCells{
			Cells: []*xlsxMergeCell{
				{
					Ref: "A1:B1",
				},
			},
		},
	}, rows, 1, -1))
	assert.NoError(t, f.adjustMergeCells(&xlsxWorksheet{
		MergeCells: &xlsxMergeCells{
			Cells: []*xlsxMergeCell{
				{
					Ref: "A1:A2",
				},
			},
		},
	}, columns, 1, -1))
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

func TestAdjustCalcChain(t *testing.T) {
	f := NewFile()
	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{
			{R: "B2"},
		},
	}
	assert.NoError(t, f.InsertCol("Sheet1", "A"))
	assert.NoError(t, f.InsertRow("Sheet1", 1))

	f.CalcChain.C[0].R = "invalid coordinates"
	assert.EqualError(t, f.InsertCol("Sheet1", "A"), `cannot convert cell "invalid coordinates" to coordinates: invalid cell name "invalid coordinates"`)
	f.CalcChain = nil
	assert.NoError(t, f.InsertCol("Sheet1", "A"))
}

func TestCoordinatesToAreaRef(t *testing.T) {
	f := NewFile()
	ref, err := f.coordinatesToAreaRef([]int{})
	assert.EqualError(t, err, "coordinates length must be 4")
	ref, err = f.coordinatesToAreaRef([]int{1, -1, 1, 1})
	assert.EqualError(t, err, "invalid cell coordinates [1, -1]")
	ref, err = f.coordinatesToAreaRef([]int{1, 1, 1, -1})
	assert.EqualError(t, err, "invalid cell coordinates [1, -1]")
	ref, err = f.coordinatesToAreaRef([]int{1, 1, 1, 1})
	assert.NoError(t, err)
	assert.EqualValues(t, ref, "A1:A1")
}
