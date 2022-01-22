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
	}, rows, 0, 0), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	assert.EqualError(t, f.adjustMergeCells(&xlsxWorksheet{
		MergeCells: &xlsxMergeCells{
			Cells: []*xlsxMergeCell{
				{
					Ref: "A1:B",
				},
			},
		},
	}, rows, 0, 0), newCellNameToCoordinatesError("B", newInvalidCellNameError("B")).Error())
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

	// testing adjustMergeCells
	var cases []struct {
		label      string
		ws         *xlsxWorksheet
		dir        adjustDirection
		num        int
		offset     int
		expect     string
		expectRect []int
	}

	// testing insert
	cases = []struct {
		label      string
		ws         *xlsxWorksheet
		dir        adjustDirection
		num        int
		offset     int
		expect     string
		expectRect []int
	}{
		{
			label: "insert row on ref",
			ws: &xlsxWorksheet{
				MergeCells: &xlsxMergeCells{
					Cells: []*xlsxMergeCell{
						{
							Ref:  "A2:B3",
							rect: []int{1, 2, 2, 3},
						},
					},
				},
			},
			dir:        rows,
			num:        2,
			offset:     1,
			expect:     "A3:B4",
			expectRect: []int{1, 3, 2, 4},
		},
		{
			label: "insert row on bottom of ref",
			ws: &xlsxWorksheet{
				MergeCells: &xlsxMergeCells{
					Cells: []*xlsxMergeCell{
						{
							Ref:  "A2:B3",
							rect: []int{1, 2, 2, 3},
						},
					},
				},
			},
			dir:        rows,
			num:        3,
			offset:     1,
			expect:     "A2:B4",
			expectRect: []int{1, 2, 2, 4},
		},
		{
			label: "insert column on the left",
			ws: &xlsxWorksheet{
				MergeCells: &xlsxMergeCells{
					Cells: []*xlsxMergeCell{
						{
							Ref:  "A2:B3",
							rect: []int{1, 2, 2, 3},
						},
					},
				},
			},
			dir:        columns,
			num:        1,
			offset:     1,
			expect:     "B2:C3",
			expectRect: []int{2, 2, 3, 3},
		},
	}
	for _, c := range cases {
		assert.NoError(t, f.adjustMergeCells(c.ws, c.dir, c.num, 1))
		assert.Equal(t, c.expect, c.ws.MergeCells.Cells[0].Ref, c.label)
		assert.Equal(t, c.expectRect, c.ws.MergeCells.Cells[0].rect, c.label)
	}

	// testing delete
	cases = []struct {
		label      string
		ws         *xlsxWorksheet
		dir        adjustDirection
		num        int
		offset     int
		expect     string
		expectRect []int
	}{
		{
			label: "delete row on top of ref",
			ws: &xlsxWorksheet{
				MergeCells: &xlsxMergeCells{
					Cells: []*xlsxMergeCell{
						{
							Ref:  "A2:B3",
							rect: []int{1, 2, 2, 3},
						},
					},
				},
			},
			dir:        rows,
			num:        2,
			offset:     -1,
			expect:     "A2:B2",
			expectRect: []int{1, 2, 2, 2},
		},
		{
			label: "delete row on bottom of ref",
			ws: &xlsxWorksheet{
				MergeCells: &xlsxMergeCells{
					Cells: []*xlsxMergeCell{
						{
							Ref:  "A2:B3",
							rect: []int{1, 2, 2, 3},
						},
					},
				},
			},
			dir:        rows,
			num:        3,
			offset:     -1,
			expect:     "A2:B2",
			expectRect: []int{1, 2, 2, 2},
		},
		{
			label: "delete column on the ref left",
			ws: &xlsxWorksheet{
				MergeCells: &xlsxMergeCells{
					Cells: []*xlsxMergeCell{
						{
							Ref:  "A2:B3",
							rect: []int{1, 2, 2, 3},
						},
					},
				},
			},
			dir:        columns,
			num:        1,
			offset:     -1,
			expect:     "A2:A3",
			expectRect: []int{1, 2, 1, 3},
		},
		{
			label: "delete column on the ref right",
			ws: &xlsxWorksheet{
				MergeCells: &xlsxMergeCells{
					Cells: []*xlsxMergeCell{
						{
							Ref:  "A2:B3",
							rect: []int{1, 2, 2, 3},
						},
					},
				},
			},
			dir:        columns,
			num:        2,
			offset:     -1,
			expect:     "A2:A3",
			expectRect: []int{1, 2, 1, 3},
		},
	}
	for _, c := range cases {
		assert.NoError(t, f.adjustMergeCells(c.ws, c.dir, c.num, -1))
		assert.Equal(t, c.expect, c.ws.MergeCells.Cells[0].Ref, c.label)
	}

	// testing delete one row/column
	cases = []struct {
		label      string
		ws         *xlsxWorksheet
		dir        adjustDirection
		num        int
		offset     int
		expect     string
		expectRect []int
	}{
		{
			label: "delete one row ref",
			ws: &xlsxWorksheet{
				MergeCells: &xlsxMergeCells{
					Cells: []*xlsxMergeCell{
						{
							Ref:  "A1:B1",
							rect: []int{1, 1, 2, 1},
						},
					},
				},
			},
			dir:    rows,
			num:    1,
			offset: -1,
		},
		{
			label: "delete one column ref",
			ws: &xlsxWorksheet{
				MergeCells: &xlsxMergeCells{
					Cells: []*xlsxMergeCell{
						{
							Ref:  "A1:A2",
							rect: []int{1, 1, 1, 2},
						},
					},
				},
			},
			dir:    columns,
			num:    1,
			offset: -1,
		},
	}
	for _, c := range cases {
		assert.NoError(t, f.adjustMergeCells(c.ws, c.dir, c.num, -1))
		assert.Equal(t, 0, len(c.ws.MergeCells.Cells), c.label)
	}

	f = NewFile()
	p1, p2 := f.adjustMergeCellsHelper(2, 1, 0, 0)
	assert.Equal(t, 1, p1)
	assert.Equal(t, 2, p2)
	f.deleteMergeCell(nil, -1)
}

func TestAdjustAutoFilter(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.adjustAutoFilter(&xlsxWorksheet{
		SheetData: xlsxSheetData{
			Row: []xlsxRow{{Hidden: true, R: 2}},
		},
		AutoFilter: &xlsxAutoFilter{
			Ref: "A1:A3",
		},
	}, rows, 1, -1))
	// testing adjustAutoFilter with illegal cell coordinates.
	assert.EqualError(t, f.adjustAutoFilter(&xlsxWorksheet{
		AutoFilter: &xlsxAutoFilter{
			Ref: "A:B1",
		},
	}, rows, 0, 0), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	assert.EqualError(t, f.adjustAutoFilter(&xlsxWorksheet{
		AutoFilter: &xlsxAutoFilter{
			Ref: "A1:B",
		},
	}, rows, 0, 0), newCellNameToCoordinatesError("B", newInvalidCellNameError("B")).Error())
}

func TestAdjustHelper(t *testing.T) {
	f := NewFile()
	f.NewSheet("Sheet2")
	f.Sheet.Store("xl/worksheets/sheet1.xml", &xlsxWorksheet{
		MergeCells: &xlsxMergeCells{Cells: []*xlsxMergeCell{{Ref: "A:B1"}}},
	})
	f.Sheet.Store("xl/worksheets/sheet2.xml", &xlsxWorksheet{
		AutoFilter: &xlsxAutoFilter{Ref: "A1:B"},
	})
	// testing adjustHelper with illegal cell coordinates.
	assert.EqualError(t, f.adjustHelper("Sheet1", rows, 0, 0), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	assert.EqualError(t, f.adjustHelper("Sheet2", rows, 0, 0), newCellNameToCoordinatesError("B", newInvalidCellNameError("B")).Error())
	// testing adjustHelper on not exists worksheet.
	assert.EqualError(t, f.adjustHelper("SheetN", rows, 0, 0), "sheet SheetN is not exist")
}

func TestAdjustCalcChain(t *testing.T) {
	f := NewFile()
	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{
			{R: "B2", I: 2}, {R: "B2", I: 1},
		},
	}
	assert.NoError(t, f.InsertCol("Sheet1", "A"))
	assert.NoError(t, f.InsertRow("Sheet1", 1))

	f.CalcChain.C[1].R = "invalid coordinates"
	assert.EqualError(t, f.InsertCol("Sheet1", "A"), newCellNameToCoordinatesError("invalid coordinates", newInvalidCellNameError("invalid coordinates")).Error())
	f.CalcChain = nil
	assert.NoError(t, f.InsertCol("Sheet1", "A"))
}
