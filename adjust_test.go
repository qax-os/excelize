package excelize

import (
	"fmt"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestAdjustMergeCells(t *testing.T) {
	f := NewFile()
	// Test adjustAutoFilter with illegal cell reference.
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
	assert.NoError(t, f.adjustMergeCells(&xlsxWorksheet{
		MergeCells: &xlsxMergeCells{
			Cells: []*xlsxMergeCell{
				{
					Ref: "A2",
				},
			},
		},
	}, columns, 1, -1))

	// Test adjustMergeCells.
	var cases []struct {
		label      string
		ws         *xlsxWorksheet
		dir        adjustDirection
		num        int
		offset     int
		expect     string
		expectRect []int
	}

	// Test insert.
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

	// Test delete,
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

	// Test delete one row or column
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
	// Test adjustAutoFilter with illegal cell reference.
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

func TestAdjustTable(t *testing.T) {
	f, sheetName := NewFile(), "Sheet1"
	for idx, tableRange := range [][]string{{"B2", "C3"}, {"E3", "F5"}, {"H5", "H8"}, {"J5", "K9"}} {
		assert.NoError(t, f.AddTable(sheetName, tableRange[0], tableRange[1], fmt.Sprintf(`{
	      "table_name": "table%d",
	      "table_style": "TableStyleMedium2",
	      "show_first_column": true,
	      "show_last_column": true,
	      "show_row_stripes": false,
	      "show_column_stripes": true
	  }`, idx)))
	}
	assert.NoError(t, f.RemoveRow(sheetName, 2))
	assert.NoError(t, f.RemoveRow(sheetName, 3))
	assert.NoError(t, f.RemoveCol(sheetName, "H"))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAdjustTable.xlsx")))

	f = NewFile()
	assert.NoError(t, f.AddTable(sheetName, "A1", "D5", ""))
	// Test adjust table with non-table part.
	f.Pkg.Delete("xl/tables/table1.xml")
	assert.NoError(t, f.RemoveRow(sheetName, 1))
	// Test adjust table with unsupported charset.
	f.Pkg.Store("xl/tables/table1.xml", MacintoshCyrillicCharset)
	assert.NoError(t, f.RemoveRow(sheetName, 1))
	// Test adjust table with invalid table range reference.
	f.Pkg.Store("xl/tables/table1.xml", []byte(`<table ref="-" />`))
	assert.NoError(t, f.RemoveRow(sheetName, 1))
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
	// Test adjustHelper with illegal cell reference.
	assert.EqualError(t, f.adjustHelper("Sheet1", rows, 0, 0), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	assert.EqualError(t, f.adjustHelper("Sheet2", rows, 0, 0), newCellNameToCoordinatesError("B", newInvalidCellNameError("B")).Error())
	// Test adjustHelper on not exists worksheet.
	assert.EqualError(t, f.adjustHelper("SheetN", rows, 0, 0), "sheet SheetN does not exist")
}

func TestAdjustCalcChain(t *testing.T) {
	f := NewFile()
	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{
			{R: "B2", I: 2}, {R: "B2", I: 1},
		},
	}
	assert.NoError(t, f.InsertCols("Sheet1", "A", 1))
	assert.NoError(t, f.InsertRows("Sheet1", 1, 1))

	f.CalcChain.C[1].R = "invalid coordinates"
	assert.EqualError(t, f.InsertCols("Sheet1", "A", 1), newCellNameToCoordinatesError("invalid coordinates", newInvalidCellNameError("invalid coordinates")).Error())
	f.CalcChain = nil
	assert.NoError(t, f.InsertCols("Sheet1", "A", 1))
}

func TestAdjustCols(t *testing.T) {
	sheetName := "Sheet1"
	preset := func() (*File, error) {
		f := NewFile()
		if err := f.SetColWidth(sheetName, "J", "T", 5); err != nil {
			return f, err
		}
		if err := f.SetSheetRow(sheetName, "J1", &[]string{"J1", "K1", "L1", "M1", "N1", "O1", "P1", "Q1", "R1", "S1", "T1"}); err != nil {
			return f, err
		}
		return f, nil
	}
	baseTbl := []string{"B", "J", "O", "O", "O", "U", "V"}
	insertTbl := []int{2, 2, 2, 5, 6, 2, 2}
	expectedTbl := []map[string]float64{
		{"J": defaultColWidth, "K": defaultColWidth, "U": 5, "V": 5, "W": defaultColWidth},
		{"J": defaultColWidth, "K": defaultColWidth, "U": 5, "V": 5, "W": defaultColWidth},
		{"O": 5, "P": 5, "U": 5, "V": 5, "W": defaultColWidth},
		{"O": 5, "S": 5, "X": 5, "Y": 5, "Z": defaultColWidth},
		{"O": 5, "S": 5, "Y": 5, "X": 5, "AA": defaultColWidth},
		{"U": 5, "V": 5, "W": defaultColWidth},
		{"U": defaultColWidth, "V": defaultColWidth, "W": defaultColWidth},
	}
	for idx, columnName := range baseTbl {
		f, err := preset()
		assert.NoError(t, err)
		assert.NoError(t, f.InsertCols(sheetName, columnName, insertTbl[idx]))
		for column, expected := range expectedTbl[idx] {
			width, err := f.GetColWidth(sheetName, column)
			assert.NoError(t, err)
			assert.Equal(t, expected, width, column)
		}
		assert.NoError(t, f.Close())
	}

	baseTbl = []string{"B", "J", "O", "T"}
	expectedTbl = []map[string]float64{
		{"H": defaultColWidth, "I": 5, "S": 5, "T": defaultColWidth},
		{"I": defaultColWidth, "J": 5, "S": 5, "T": defaultColWidth},
		{"I": defaultColWidth, "O": 5, "S": 5, "T": defaultColWidth},
		{"R": 5, "S": 5, "T": defaultColWidth, "U": defaultColWidth},
	}
	for idx, columnName := range baseTbl {
		f, err := preset()
		assert.NoError(t, err)
		assert.NoError(t, f.RemoveCol(sheetName, columnName))
		for column, expected := range expectedTbl[idx] {
			width, err := f.GetColWidth(sheetName, column)
			assert.NoError(t, err)
			assert.Equal(t, expected, width, column)
		}
		assert.NoError(t, f.Close())
	}

	f, err := preset()
	assert.NoError(t, err)
	assert.NoError(t, f.SetColWidth(sheetName, "I", "I", 8))
	for i := 0; i <= 12; i++ {
		assert.NoError(t, f.RemoveCol(sheetName, "I"))
	}
	for c := 9; c <= 21; c++ {
		columnName, err := ColumnNumberToName(c)
		assert.NoError(t, err)
		width, err := f.GetColWidth(sheetName, columnName)
		assert.NoError(t, err)
		assert.Equal(t, defaultColWidth, width, columnName)
	}

	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).Cols = nil
	assert.NoError(t, f.RemoveCol(sheetName, "A"))

	assert.NoError(t, f.Close())
}
