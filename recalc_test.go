// Copyright 2016 - 2026 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"strings"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
)

// cellXML fetches the raw xlsxC for a cell so the test can inspect
// <f>, <v>, and <t> directly without going through GetCellValue's
// formatting layer.
func cellXML(t *testing.T, f *File, sheet, addr string) *xlsxC {
	t.Helper()
	ws, err := f.workSheetReader(sheet)
	assert.NoError(t, err)
	col, row, err := CellNameToCoordinates(addr)
	assert.NoError(t, err)
	assert.LessOrEqual(t, row, len(ws.SheetData.Row), addr)
	rowData := ws.SheetData.Row[row-1]
	assert.LessOrEqual(t, col, len(rowData.C), addr)
	return &rowData.C[col-1]
}

func TestRecalcCellTypes(t *testing.T) {
	cases := []struct {
		name    string
		formula string
		inputs  map[string]int
		wantT   string
		wantV   string
	}{
		{"numeric", "SUM(A1:A2)", map[string]int{"A1": 10, "A2": 32}, "", "42"},
		{"boolean", "A1>0", map[string]int{"A1": 5}, "b", "1"},
		{"chained", "A1*3", map[string]int{"A1": 7}, "", "21"},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			f := NewFile()
			for k, v := range tc.inputs {
				assert.NoError(t, f.SetCellInt("Sheet1", k, int64(v)))
			}
			assert.NoError(t, f.SetCellFormula("Sheet1", "B1", tc.formula))
			assert.NoError(t, f.RecalcCell("Sheet1", "B1"))
			c := cellXML(t, f, "Sheet1", "B1")
			assert.Equal(t, tc.formula, c.F.Content, "formula preserved")
			assert.Equal(t, tc.wantT, c.T)
			assert.Equal(t, tc.wantV, c.V)
		})
	}
}

func TestRecalcCellNoFormula(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellInt("Sheet1", "A1", 5))
	assert.ErrorIs(t, f.RecalcCell("Sheet1", "A1"), ErrCellNoFormula)
}

func TestRecalcCellNoTypeSRegression(t *testing.T) {
	// Previously a recalc path round-tripped numeric results through
	// string, storing cells as t="s" (shared string). Downstream
	// aggregates over that blob silently summed to 0. RecalcCell must
	// persist numeric results with an implicit t attribute.
	f := NewFile()
	assert.NoError(t, f.SetCellInt("Sheet1", "A1", 3))
	assert.NoError(t, f.SetCellInt("Sheet1", "A2", 4))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A3", "SUM(A1:A2)"))
	assert.NoError(t, f.RecalcCell("Sheet1", "A3"))
	c := cellXML(t, f, "Sheet1", "A3")
	assert.NotEqual(t, "s", c.T)
	assert.Equal(t, "", c.T)
	assert.Equal(t, "7", c.V)
}

func TestRecalcCellSharedFormulaBand(t *testing.T) {
	// Row 40 shared-formula group with master G40; RecalcCell on each
	// member must leave the shared metadata intact.
	f := NewFile()
	for col, v := range map[string]int64{"G": 2, "H": 3, "I": 4, "J": 5} {
		assert.NoError(t, f.SetCellInt("Sheet1", col+"1", v))
	}
	ref := "G40:J40"
	sharedType := STCellFormulaTypeShared
	assert.NoError(t, f.SetCellFormula("Sheet1", "G40", "SUM(G1:G39)", FormulaOpts{
		Type: &sharedType,
		Ref:  &ref,
	}))
	master := cellXML(t, f, "Sheet1", "G40")
	si := *master.F.Si

	for _, addr := range []string{"G40", "H40", "I40", "J40"} {
		assert.NoError(t, f.RecalcCell("Sheet1", addr))
	}

	master = cellXML(t, f, "Sheet1", "G40")
	assert.Equal(t, "SUM(G1:G39)", master.F.Content)
	assert.Equal(t, ref, master.F.Ref)
	assert.Equal(t, si, *master.F.Si)
	assert.Equal(t, "2", master.V)
	for _, child := range []string{"H40", "I40", "J40"} {
		c := cellXML(t, f, "Sheet1", child)
		assert.Equal(t, sharedType, c.F.T, child)
		assert.Equal(t, "", c.F.Ref, child)
		assert.Equal(t, si, *c.F.Si, child)
	}
}

func TestRecalcCellCircularReferenceBounded(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "B1+1"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "A1+1"))
	done := make(chan struct{})
	go func() {
		_ = f.RecalcCell("Sheet1", "A1")
		close(done)
	}()
	select {
	case <-done:
	case <-time.After(2 * time.Second):
		t.Fatal("RecalcCell did not terminate on circular reference")
	}
}

func TestRecalcCellRoundTrip(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellInt("Sheet1", "A1", 2))
	assert.NoError(t, f.SetCellInt("Sheet1", "A2", 3))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A3", "A1*A2"))
	assert.NoError(t, f.RecalcCell("Sheet1", "A3"))

	buf, err := f.WriteToBuffer()
	assert.NoError(t, err)
	f2, err := OpenReader(strings.NewReader(buf.String()))
	assert.NoError(t, err)
	defer func() { _ = f2.Close() }()

	formula, err := f2.GetCellFormula("Sheet1", "A3")
	assert.NoError(t, err)
	assert.Equal(t, "A1*A2", formula)
	value, err := f2.GetCellValue("Sheet1", "A3", Options{RawCellValue: true})
	assert.NoError(t, err)
	assert.Equal(t, "6", value)
}
