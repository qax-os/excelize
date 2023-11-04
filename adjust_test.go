package excelize

import (
	"fmt"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestAdjustMergeCells(t *testing.T) {
	f := NewFile()
	// Test adjustAutoFilter with illegal cell reference
	assert.Equal(t, f.adjustMergeCells(&xlsxWorksheet{
		MergeCells: &xlsxMergeCells{
			Cells: []*xlsxMergeCell{
				{
					Ref: "A:B1",
				},
			},
		},
	}, rows, 0, 0), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")))
	assert.Equal(t, f.adjustMergeCells(&xlsxWorksheet{
		MergeCells: &xlsxMergeCells{
			Cells: []*xlsxMergeCell{
				{
					Ref: "A1:B",
				},
			},
		},
	}, rows, 0, 0), newCellNameToCoordinatesError("B", newInvalidCellNameError("B")))
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

	// Test adjust merge cells
	var cases []struct {
		label      string
		ws         *xlsxWorksheet
		dir        adjustDirection
		num        int
		offset     int
		expect     string
		expectRect []int
	}

	// Test adjust merged cell when insert rows and columns
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

	// Test adjust merged cells when delete rows and columns
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
		assert.Len(t, c.ws.MergeCells.Cells, 0, c.label)
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
	// Test adjustAutoFilter with illegal cell reference
	assert.Equal(t, f.adjustAutoFilter(&xlsxWorksheet{
		AutoFilter: &xlsxAutoFilter{
			Ref: "A:B1",
		},
	}, rows, 0, 0), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")))
	assert.Equal(t, f.adjustAutoFilter(&xlsxWorksheet{
		AutoFilter: &xlsxAutoFilter{
			Ref: "A1:B",
		},
	}, rows, 0, 0), newCellNameToCoordinatesError("B", newInvalidCellNameError("B")))
}

func TestAdjustTable(t *testing.T) {
	f, sheetName := NewFile(), "Sheet1"
	for idx, reference := range []string{"B2:C3", "E3:F5", "H5:H8", "J5:K9"} {
		assert.NoError(t, f.AddTable(sheetName, &Table{
			Range:             reference,
			Name:              fmt.Sprintf("table%d", idx),
			StyleName:         "TableStyleMedium2",
			ShowFirstColumn:   true,
			ShowLastColumn:    true,
			ShowRowStripes:    boolPtr(false),
			ShowColumnStripes: true,
		}))
	}
	assert.NoError(t, f.RemoveRow(sheetName, 2))
	assert.NoError(t, f.RemoveRow(sheetName, 3))
	assert.NoError(t, f.RemoveRow(sheetName, 3))
	assert.NoError(t, f.RemoveCol(sheetName, "H"))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAdjustTable.xlsx")))

	f = NewFile()
	assert.NoError(t, f.AddTable(sheetName, &Table{Range: "A1:D5"}))
	// Test adjust table with non-table part
	f.Pkg.Delete("xl/tables/table1.xml")
	assert.NoError(t, f.RemoveRow(sheetName, 1))
	// Test adjust table with unsupported charset
	f.Pkg.Store("xl/tables/table1.xml", MacintoshCyrillicCharset)
	assert.NoError(t, f.RemoveRow(sheetName, 1))
	// Test adjust table with invalid table range reference
	f.Pkg.Store("xl/tables/table1.xml", []byte(`<table ref="-" />`))
	assert.NoError(t, f.RemoveRow(sheetName, 1))
}

func TestAdjustHelper(t *testing.T) {
	f := NewFile()
	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)
	f.Sheet.Store("xl/worksheets/sheet1.xml", &xlsxWorksheet{
		MergeCells: &xlsxMergeCells{Cells: []*xlsxMergeCell{{Ref: "A:B1"}}},
	})
	f.Sheet.Store("xl/worksheets/sheet2.xml", &xlsxWorksheet{
		AutoFilter: &xlsxAutoFilter{Ref: "A1:B"},
	})
	// Test adjustHelper with illegal cell reference
	assert.Equal(t, f.adjustHelper("Sheet1", rows, 0, 0), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")))
	assert.Equal(t, f.adjustHelper("Sheet2", rows, 0, 0), newCellNameToCoordinatesError("B", newInvalidCellNameError("B")))
	// Test adjustHelper on not exists worksheet
	assert.EqualError(t, f.adjustHelper("SheetN", rows, 0, 0), "sheet SheetN does not exist")
}

func TestAdjustCalcChain(t *testing.T) {
	f := NewFile()
	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{{R: "B2", I: 2}, {R: "B2", I: 1}, {R: "A1", I: 1}},
	}
	assert.NoError(t, f.InsertCols("Sheet1", "A", 1))
	assert.NoError(t, f.InsertRows("Sheet1", 1, 1))

	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{{R: "B2", I: 1}, {R: "B3"}, {R: "A1"}},
	}
	assert.NoError(t, f.RemoveRow("Sheet1", 3))
	assert.NoError(t, f.RemoveCol("Sheet1", "B"))

	f.CalcChain = &xlsxCalcChain{C: []xlsxCalcChainC{{R: "B2", I: 2}, {R: "B2", I: 1}}}
	f.CalcChain.C[1].R = "invalid coordinates"
	assert.Equal(t, f.InsertCols("Sheet1", "A", 1), newCellNameToCoordinatesError("invalid coordinates", newInvalidCellNameError("invalid coordinates")))
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

func TestAdjustColDimensions(t *testing.T) {
	f := NewFile()
	ws, err := f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellFormula("Sheet1", "C3", "A1+B1"))
	assert.Equal(t, ErrColumnNumber, f.adjustColDimensions("Sheet1", ws, 1, MaxColumns))

	_, err = f.NewSheet("Sheet2")
	assert.NoError(t, err)
	f.Sheet.Delete("xl/worksheets/sheet2.xml")
	f.Pkg.Store("xl/worksheets/sheet2.xml", MacintoshCyrillicCharset)
	assert.EqualError(t, f.adjustColDimensions("Sheet2", ws, 2, 1), "XML syntax error on line 1: invalid UTF-8")
}

func TestAdjustRowDimensions(t *testing.T) {
	f := NewFile()
	ws, err := f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellFormula("Sheet1", "C3", "A1+B1"))
	assert.Equal(t, ErrMaxRows, f.adjustRowDimensions("Sheet1", ws, 1, TotalRows))

	_, err = f.NewSheet("Sheet2")
	assert.NoError(t, err)
	f.Sheet.Delete("xl/worksheets/sheet2.xml")
	f.Pkg.Store("xl/worksheets/sheet2.xml", MacintoshCyrillicCharset)
	assert.EqualError(t, f.adjustRowDimensions("Sheet1", ws, 2, 1), "XML syntax error on line 1: invalid UTF-8")

	f = NewFile()
	_, err = f.NewSheet("Sheet2")
	assert.NoError(t, err)
	ws, err = f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellFormula("Sheet1", "B2", fmt.Sprintf("Sheet2!A%d", TotalRows)))
	assert.Equal(t, ErrMaxRows, f.adjustRowDimensions("Sheet2", ws, 1, TotalRows))
}

func TestAdjustHyperlinks(t *testing.T) {
	f := NewFile()
	ws, err := f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellFormula("Sheet1", "C3", "A1+B1"))
	f.adjustHyperlinks(ws, "Sheet1", rows, 3, -1)

	// Test adjust hyperlinks location with positive offset
	assert.NoError(t, f.SetCellHyperLink("Sheet1", "F5", "Sheet1!A1", "Location"))
	assert.NoError(t, f.InsertRows("Sheet1", 1, 1))
	link, target, err := f.GetCellHyperLink("Sheet1", "F6")
	assert.NoError(t, err)
	assert.True(t, link)
	assert.Equal(t, target, "Sheet1!A1")

	// Test adjust hyperlinks location with negative offset
	assert.NoError(t, f.RemoveRow("Sheet1", 1))
	link, target, err = f.GetCellHyperLink("Sheet1", "F5")
	assert.NoError(t, err)
	assert.True(t, link)
	assert.Equal(t, target, "Sheet1!A1")

	// Test adjust hyperlinks location on remove row
	assert.NoError(t, f.RemoveRow("Sheet1", 5))
	link, target, err = f.GetCellHyperLink("Sheet1", "F5")
	assert.NoError(t, err)
	assert.False(t, link)
	assert.Empty(t, target)

	// Test adjust hyperlinks location on remove column
	assert.NoError(t, f.SetCellHyperLink("Sheet1", "F5", "Sheet1!A1", "Location"))
	assert.NoError(t, f.RemoveCol("Sheet1", "F"))
	link, target, err = f.GetCellHyperLink("Sheet1", "F5")
	assert.NoError(t, err)
	assert.False(t, link)
	assert.Empty(t, target)

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAdjustHyperlinks.xlsx")))
	assert.NoError(t, f.Close())
}

func TestAdjustFormula(t *testing.T) {
	f := NewFile()
	formulaType, ref := STCellFormulaTypeShared, "C1:C5"
	assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "A1+B1", FormulaOpts{Ref: &ref, Type: &formulaType}))
	assert.NoError(t, f.DuplicateRowTo("Sheet1", 1, 10))
	assert.NoError(t, f.InsertCols("Sheet1", "B", 1))
	assert.NoError(t, f.InsertRows("Sheet1", 1, 1))
	for cell, expected := range map[string]string{"D2": "A2+C2", "D3": "A3+C3", "D11": "A11+C11"} {
		formula, err := f.GetCellFormula("Sheet1", cell)
		assert.NoError(t, err)
		assert.Equal(t, expected, formula)
	}
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAdjustFormula.xlsx")))
	assert.NoError(t, f.Close())

	assert.NoError(t, f.adjustFormula("Sheet1", "Sheet1", nil, rows, 0, 0, false))
	assert.Equal(t, newCellNameToCoordinatesError("-", newInvalidCellNameError("-")), f.adjustFormula("Sheet1", "Sheet1", &xlsxF{Ref: "-"}, rows, 0, 0, false))
	assert.Equal(t, ErrColumnNumber, f.adjustFormula("Sheet1", "Sheet1", &xlsxF{Ref: "XFD1:XFD1"}, columns, 0, 1, false))

	_, err := f.adjustFormulaRef("Sheet1", "Sheet1", "XFE1", columns, 0, 1)
	assert.Equal(t, ErrColumnNumber, err)
	_, err = f.adjustFormulaRef("Sheet1", "Sheet1", "XFD1", columns, 0, 1)
	assert.Equal(t, ErrColumnNumber, err)

	f = NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "XFD1"))
	assert.Equal(t, ErrColumnNumber, f.InsertCols("Sheet1", "A", 1))

	assert.NoError(t, f.SetCellFormula("Sheet1", "B2", fmt.Sprintf("A%d", TotalRows)))
	assert.Equal(t, ErrMaxRows, f.InsertRows("Sheet1", 1, 1))

	f = NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "B3", "SUM(1048576:1:2)"))
	assert.Equal(t, ErrMaxRows, f.InsertRows("Sheet1", 1, 1))

	f = NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "B3", "SUM(XFD:A:B)"))
	assert.Equal(t, ErrColumnNumber, f.InsertCols("Sheet1", "A", 1))

	f = NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "B3", "SUM(A:B:XFD)"))
	assert.Equal(t, ErrColumnNumber, f.InsertCols("Sheet1", "A", 1))

	// Test adjust formula with defined name in formula text
	f = NewFile()
	assert.NoError(t, f.SetDefinedName(&DefinedName{
		Name:     "Amount",
		RefersTo: "Sheet1!$B$2",
	}))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B2", "Amount+B3"))
	assert.NoError(t, f.RemoveRow("Sheet1", 1))
	formula, err := f.GetCellFormula("Sheet1", "B1")
	assert.NoError(t, err)
	assert.Equal(t, "Amount+B2", formula)

	// Test adjust formula with array formula
	f = NewFile()
	formulaType, reference := STCellFormulaTypeArray, "A3:A3"
	assert.NoError(t, f.SetCellFormula("Sheet1", "A3", "A1:A2", FormulaOpts{Ref: &reference, Type: &formulaType}))
	assert.NoError(t, f.InsertRows("Sheet1", 1, 1))
	formula, err = f.GetCellFormula("Sheet1", "A4")
	assert.NoError(t, err)
	assert.Equal(t, "A2:A3", formula)

	// Test adjust formula on duplicate row with array formula
	f = NewFile()
	formulaType, reference = STCellFormulaTypeArray, "A3"
	assert.NoError(t, f.SetCellFormula("Sheet1", "A3", "A1:A2", FormulaOpts{Ref: &reference, Type: &formulaType}))
	assert.NoError(t, f.InsertRows("Sheet1", 1, 1))
	formula, err = f.GetCellFormula("Sheet1", "A4")
	assert.NoError(t, err)
	assert.Equal(t, "A2:A3", formula)

	// Test adjust formula on duplicate row with relative and absolute cell references
	f = NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "B10", "A$10+$A11&\" \""))
	assert.NoError(t, f.DuplicateRowTo("Sheet1", 10, 2))
	formula, err = f.GetCellFormula("Sheet1", "B2")
	assert.NoError(t, err)
	assert.Equal(t, "A$2+$A3&\" \"", formula)

	t.Run("for_cells_affected_directly", func(t *testing.T) {
		// Test insert row in middle of range with relative and absolute cell references
		f := NewFile()
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "$A1+A$2"))
		assert.NoError(t, f.InsertRows("Sheet1", 2, 1))
		formula, err := f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "$A1+A$3", formula)
		assert.NoError(t, f.RemoveRow("Sheet1", 2))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "$A1+A$2", formula)

		// Test insert column in middle of range
		f = NewFile()
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "B1+C1"))
		assert.NoError(t, f.InsertCols("Sheet1", "C", 1))
		formula, err = f.GetCellFormula("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Equal(t, "B1+D1", formula)
		assert.NoError(t, f.RemoveCol("Sheet1", "C"))
		formula, err = f.GetCellFormula("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Equal(t, "B1+C1", formula)

		// Test insert row and column in a rectangular range
		f = NewFile()
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "D4+D5+E4+E5"))
		assert.NoError(t, f.InsertCols("Sheet1", "E", 1))
		assert.NoError(t, f.InsertRows("Sheet1", 5, 1))
		formula, err = f.GetCellFormula("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Equal(t, "D4+D6+F4+F6", formula)

		// Test insert row in middle of range
		f = NewFile()
		formulaType, reference := STCellFormulaTypeArray, "B1:B1"
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "A1:A2", FormulaOpts{Ref: &reference, Type: &formulaType}))
		assert.NoError(t, f.InsertRows("Sheet1", 2, 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "A1:A3", formula)
		assert.NoError(t, f.RemoveRow("Sheet1", 2))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "A1:A2", formula)

		// Test insert column in middle of range
		f = NewFile()
		formulaType, reference = STCellFormulaTypeArray, "A1:A1"
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "B1:C1", FormulaOpts{Ref: &reference, Type: &formulaType}))
		assert.NoError(t, f.InsertCols("Sheet1", "C", 1))
		formula, err = f.GetCellFormula("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Equal(t, "B1:D1", formula)
		assert.NoError(t, f.RemoveCol("Sheet1", "C"))
		formula, err = f.GetCellFormula("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Equal(t, "B1:C1", formula)

		// Test insert row and column in a rectangular range
		f = NewFile()
		formulaType, reference = STCellFormulaTypeArray, "A1:A1"
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "D4:E5", FormulaOpts{Ref: &reference, Type: &formulaType}))
		assert.NoError(t, f.InsertCols("Sheet1", "E", 1))
		assert.NoError(t, f.InsertRows("Sheet1", 5, 1))
		formula, err = f.GetCellFormula("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Equal(t, "D4:F6", formula)
	})
	t.Run("for_cells_affected_indirectly", func(t *testing.T) {
		f := NewFile()
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "A3+A4"))
		assert.NoError(t, f.InsertRows("Sheet1", 2, 1))
		formula, err := f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "A4+A5", formula)
		assert.NoError(t, f.RemoveRow("Sheet1", 2))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "A3+A4", formula)

		f = NewFile()
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "D3+D4"))
		assert.NoError(t, f.InsertCols("Sheet1", "C", 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "E3+E4", formula)
		assert.NoError(t, f.RemoveCol("Sheet1", "C"))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "D3+D4", formula)
	})
	t.Run("for_entire_cols_rows_reference", func(t *testing.T) {
		f := NewFile()
		// Test adjust formula on insert row in the middle of the range
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "SUM(A2:A3:A4,,Table1[])"))
		assert.NoError(t, f.InsertRows("Sheet1", 3, 1))
		formula, err := f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM(A2:A4:A5,,Table1[])", formula)

		// Test adjust formula on insert at the top of the range
		assert.NoError(t, f.InsertRows("Sheet1", 2, 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM(A3:A5:A6,,Table1[])", formula)

		f = NewFile()
		// Test adjust formula on insert row in the middle of the range
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "SUM('Sheet 1'!A2,A3)"))
		assert.NoError(t, f.InsertRows("Sheet1", 3, 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM('Sheet 1'!A2,A4)", formula)

		// Test adjust formula on insert row at the top of the range
		assert.NoError(t, f.InsertRows("Sheet1", 2, 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM('Sheet 1'!A3,A5)", formula)

		f = NewFile()
		// Test adjust formula on insert col in the middle of the range
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "SUM(C3:D3)"))
		assert.NoError(t, f.InsertCols("Sheet1", "D", 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM(C3:E3)", formula)

		// Test adjust formula on insert at the top of the range
		assert.NoError(t, f.InsertCols("Sheet1", "C", 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM(D3:F3)", formula)

		f = NewFile()
		// Test adjust formula on insert column in the middle of the range
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "SUM(C3,D3)"))
		assert.NoError(t, f.InsertCols("Sheet1", "D", 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM(C3,E3)", formula)

		// Test adjust formula on insert column at the top of the range
		assert.NoError(t, f.InsertCols("Sheet1", "C", 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM(D3,F3)", formula)

		f = NewFile()
		// Test adjust formula on insert row in the middle of the range (range of whole row)
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "SUM(2:3)"))
		assert.NoError(t, f.InsertRows("Sheet1", 3, 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM(2:4)", formula)

		// Test adjust formula on insert row at the top of the range (range of whole row)
		assert.NoError(t, f.InsertRows("Sheet1", 2, 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM(3:5)", formula)

		f = NewFile()
		// Test adjust formula on insert row in the middle of the range (range of whole column)
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "SUM(C:D)"))
		assert.NoError(t, f.InsertCols("Sheet1", "D", 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM(C:E)", formula)

		// Test adjust formula on insert row at the top of the range (range of whole column)
		assert.NoError(t, f.InsertCols("Sheet1", "C", 1))
		formula, err = f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "SUM(D:F)", formula)
	})
	t.Run("for_all_worksheet_cells_with_rows_insert", func(t *testing.T) {
		f := NewFile()
		_, err := f.NewSheet("Sheet2")
		assert.NoError(t, err)
		// Tests formulas referencing Sheet2 should update but those referencing the original sheet should not
		tbl := [][]string{
			{"B1", "Sheet2!A1+Sheet2!A2", "Sheet2!A1+Sheet2!A3", "Sheet2!A2+Sheet2!A4"},
			{"C1", "A1+A2", "A1+A2", "A1+A2"},
			{"D1", "Sheet2!B1:B2", "Sheet2!B1:B3", "Sheet2!B2:B4"},
			{"E1", "B1:B2", "B1:B2", "B1:B2"},
			{"F1", "SUM(Sheet2!C1:C2)", "SUM(Sheet2!C1:C3)", "SUM(Sheet2!C2:C4)"},
			{"G1", "SUM(C1:C2)", "SUM(C1:C2)", "SUM(C1:C2)"},
			{"H1", "SUM(Sheet2!D1,Sheet2!D2)", "SUM(Sheet2!D1,Sheet2!D3)", "SUM(Sheet2!D2,Sheet2!D4)"},
			{"I1", "SUM(D1,D2)", "SUM(D1,D2)", "SUM(D1,D2)"},
		}
		for _, preset := range tbl {
			assert.NoError(t, f.SetCellFormula("Sheet1", preset[0], preset[1]))
		}
		// Test adjust formula on insert row in the middle of the range
		assert.NoError(t, f.InsertRows("Sheet2", 2, 1))
		for _, preset := range tbl {
			formula, err := f.GetCellFormula("Sheet1", preset[0])
			assert.NoError(t, err)
			assert.Equal(t, preset[2], formula)
		}

		// Test adjust formula on insert row in the top of the range
		assert.NoError(t, f.InsertRows("Sheet2", 1, 1))
		for _, preset := range tbl {
			formula, err := f.GetCellFormula("Sheet1", preset[0])
			assert.NoError(t, err)
			assert.Equal(t, preset[3], formula)
		}
	})
	t.Run("for_all_worksheet_cells_with_cols_insert", func(t *testing.T) {
		f := NewFile()
		_, err := f.NewSheet("Sheet2")
		assert.NoError(t, err)
		tbl := [][]string{
			{"A1", "Sheet2!A1+Sheet2!B1", "Sheet2!A1+Sheet2!C1", "Sheet2!B1+Sheet2!D1"},
			{"A2", "A1+B1", "A1+B1", "A1+B1"},
			{"A3", "Sheet2!A2:B2", "Sheet2!A2:C2", "Sheet2!B2:D2"},
			{"A4", "A2:B2", "A2:B2", "A2:B2"},
			{"A5", "SUM(Sheet2!A3:B3)", "SUM(Sheet2!A3:C3)", "SUM(Sheet2!B3:D3)"},
			{"A6", "SUM(A3:B3)", "SUM(A3:B3)", "SUM(A3:B3)"},
			{"A7", "SUM(Sheet2!A4,Sheet2!B4)", "SUM(Sheet2!A4,Sheet2!C4)", "SUM(Sheet2!B4,Sheet2!D4)"},
			{"A8", "SUM(A4,B4)", "SUM(A4,B4)", "SUM(A4,B4)"},
		}
		for _, preset := range tbl {
			assert.NoError(t, f.SetCellFormula("Sheet1", preset[0], preset[1]))
		}
		// Test adjust formula on insert column in the middle of the range
		assert.NoError(t, f.InsertCols("Sheet2", "B", 1))
		for _, preset := range tbl {
			formula, err := f.GetCellFormula("Sheet1", preset[0])
			assert.NoError(t, err)
			assert.Equal(t, preset[2], formula)
		}
		// Test adjust formula on insert column in the top of the range
		assert.NoError(t, f.InsertCols("Sheet2", "A", 1))
		for _, preset := range tbl {
			formula, err := f.GetCellFormula("Sheet1", preset[0])
			assert.NoError(t, err)
			assert.Equal(t, preset[3], formula)
		}
	})
	t.Run("for_cross_sheet_ref_with_rows_insert)", func(t *testing.T) {
		f := NewFile()
		_, err := f.NewSheet("Sheet2")
		assert.NoError(t, err)
		_, err = f.NewSheet("Sheet3")
		assert.NoError(t, err)
		// Tests formulas referencing Sheet2 should update but those referencing
		// the original sheet or Sheet 3 should not update
		tbl := [][]string{
			{"B1", "Sheet2!A1+Sheet2!A2+Sheet1!A3+Sheet1!A4", "Sheet2!A1+Sheet2!A3+Sheet1!A3+Sheet1!A4", "Sheet2!A2+Sheet2!A4+Sheet1!A3+Sheet1!A4"},
			{"C1", "Sheet2!B1+Sheet2!B2+B3+B4", "Sheet2!B1+Sheet2!B3+B3+B4", "Sheet2!B2+Sheet2!B4+B3+B4"},
			{"D1", "Sheet2!C1+Sheet2!C2+Sheet3!A3+Sheet3!A4", "Sheet2!C1+Sheet2!C3+Sheet3!A3+Sheet3!A4", "Sheet2!C2+Sheet2!C4+Sheet3!A3+Sheet3!A4"},
			{"E1", "SUM(Sheet2!D1:D2,Sheet1!A3:A4)", "SUM(Sheet2!D1:D3,Sheet1!A3:A4)", "SUM(Sheet2!D2:D4,Sheet1!A3:A4)"},
			{"F1", "SUM(Sheet2!E1:E2,A3:A4)", "SUM(Sheet2!E1:E3,A3:A4)", "SUM(Sheet2!E2:E4,A3:A4)"},
			{"G1", "SUM(Sheet2!F1:F2,Sheet3!A3:A4)", "SUM(Sheet2!F1:F3,Sheet3!A3:A4)", "SUM(Sheet2!F2:F4,Sheet3!A3:A4)"},
		}
		for _, preset := range tbl {
			assert.NoError(t, f.SetCellFormula("Sheet1", preset[0], preset[1]))
		}
		// Test adjust formula on insert row in the middle of the range
		assert.NoError(t, f.InsertRows("Sheet2", 2, 1))
		for _, preset := range tbl {
			formula, err := f.GetCellFormula("Sheet1", preset[0])
			assert.NoError(t, err)
			assert.Equal(t, preset[2], formula)
		}
		// Test adjust formula on insert row in the top of the range
		assert.NoError(t, f.InsertRows("Sheet2", 1, 1))
		for _, preset := range tbl {
			formula, err := f.GetCellFormula("Sheet1", preset[0])
			assert.NoError(t, err)
			assert.Equal(t, preset[3], formula)
		}
	})
	t.Run("for_cross_sheet_ref_with_cols_insert)", func(t *testing.T) {
		f := NewFile()
		_, err := f.NewSheet("Sheet2")
		assert.NoError(t, err)
		_, err = f.NewSheet("Sheet3")
		assert.NoError(t, err)
		// Tests formulas referencing Sheet2 should update but those referencing
		// the original sheet or Sheet 3 should not update
		tbl := [][]string{
			{"A1", "Sheet2!A1+Sheet2!B1+Sheet1!C1+Sheet1!D1", "Sheet2!A1+Sheet2!C1+Sheet1!C1+Sheet1!D1", "Sheet2!B1+Sheet2!D1+Sheet1!C1+Sheet1!D1"},
			{"A2", "Sheet2!A2+Sheet2!B2+C2+D2", "Sheet2!A2+Sheet2!C2+C2+D2", "Sheet2!B2+Sheet2!D2+C2+D2"},
			{"A3", "Sheet2!A3+Sheet2!B3+Sheet3!C3+Sheet3!D3", "Sheet2!A3+Sheet2!C3+Sheet3!C3+Sheet3!D3", "Sheet2!B3+Sheet2!D3+Sheet3!C3+Sheet3!D3"},
			{"A4", "SUM(Sheet2!A4:B4,Sheet1!C4:D4)", "SUM(Sheet2!A4:C4,Sheet1!C4:D4)", "SUM(Sheet2!B4:D4,Sheet1!C4:D4)"},
			{"A5", "SUM(Sheet2!A5:B5,C5:D5)", "SUM(Sheet2!A5:C5,C5:D5)", "SUM(Sheet2!B5:D5,C5:D5)"},
			{"A6", "SUM(Sheet2!A6:B6,Sheet3!C6:D6)", "SUM(Sheet2!A6:C6,Sheet3!C6:D6)", "SUM(Sheet2!B6:D6,Sheet3!C6:D6)"},
		}
		for _, preset := range tbl {
			assert.NoError(t, f.SetCellFormula("Sheet1", preset[0], preset[1]))
		}
		// Test adjust formula on insert row in the middle of the range
		assert.NoError(t, f.InsertCols("Sheet2", "B", 1))
		for _, preset := range tbl {
			formula, err := f.GetCellFormula("Sheet1", preset[0])
			assert.NoError(t, err)
			assert.Equal(t, preset[2], formula)
		}
		// Test adjust formula on insert row in the top of the range
		assert.NoError(t, f.InsertCols("Sheet2", "A", 1))
		for _, preset := range tbl {
			formula, err := f.GetCellFormula("Sheet1", preset[0])
			assert.NoError(t, err)
			assert.Equal(t, preset[3], formula)
		}
	})
	t.Run("for_cross_sheet_ref_with_chart_sheet)", func(t *testing.T) {
		assert.NoError(t, f.AddChartSheet("Chart1", &Chart{Type: Line}))
		assert.NoError(t, f.InsertRows("Sheet1", 2, 1))
		assert.NoError(t, f.InsertCols("Sheet1", "A", 1))
	})
}

func TestAdjustVolatileDeps(t *testing.T) {
	f := NewFile()
	f.Pkg.Store(defaultXMLPathVolatileDeps, []byte(fmt.Sprintf(`<volTypes xmlns="%s"><volType><main><tp><tr r="C2" s="2"/><tr r="C2" s="1"/><tr r="D3" s="1"/></tp></main></volType></volTypes>`, NameSpaceSpreadSheet.Value)))
	assert.NoError(t, f.InsertCols("Sheet1", "A", 1))
	assert.NoError(t, f.InsertRows("Sheet1", 2, 1))
	assert.Equal(t, "D3", f.VolatileDeps.VolType[0].Main[0].Tp[0].Tr[1].R)
	assert.NoError(t, f.RemoveCol("Sheet1", "D"))
	assert.NoError(t, f.RemoveRow("Sheet1", 4))
	assert.Len(t, f.VolatileDeps.VolType[0].Main[0].Tp[0].Tr, 1)

	f = NewFile()
	f.Pkg.Store(defaultXMLPathVolatileDeps, MacintoshCyrillicCharset)
	assert.EqualError(t, f.InsertRows("Sheet1", 2, 1), "XML syntax error on line 1: invalid UTF-8")

	f = NewFile()
	f.Pkg.Store(defaultXMLPathVolatileDeps, []byte(fmt.Sprintf(`<volTypes xmlns="%s"><volType><main><tp><tr r="A" s="1"/></tp></main></volType></volTypes>`, NameSpaceSpreadSheet.Value)))
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), f.InsertCols("Sheet1", "A", 1))
	f.volatileDepsWriter()
}
