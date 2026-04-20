// Copyright 2016 - 2026 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"strconv"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

// cellXML fetches the raw xlsxC for a cell so the test can inspect
// `<f>`, `<v>`, and `<t>` directly.
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

func TestSetCellCachedValueNumeric(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "SUM(1,2)"))
	assert.NoError(t, f.setCellCachedValue("Sheet1", "A1", newNumberFormulaArg(42)))

	c := cellXML(t, f, "Sheet1", "A1")
	assert.NotNil(t, c.F, "formula element must be preserved")
	assert.Equal(t, "SUM(1,2)", c.F.Content, "formula body unchanged")
	assert.Equal(t, "", c.T, "numeric cell uses implicit t attribute")
	assert.Equal(t, "42", c.V)
}

func TestSetCellCachedValueBool(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "TRUE()"))
	boolArg := newBoolFormulaArg(true)
	assert.NoError(t, f.setCellCachedValue("Sheet1", "A1", boolArg))

	c := cellXML(t, f, "Sheet1", "A1")
	assert.NotNil(t, c.F)
	assert.Equal(t, "TRUE()", c.F.Content)
	assert.Equal(t, "b", c.T)
	assert.Equal(t, "1", c.V)

	assert.NoError(t, f.setCellCachedValue("Sheet1", "A1", newBoolFormulaArg(false)))
	c = cellXML(t, f, "Sheet1", "A1")
	assert.Equal(t, "b", c.T)
	assert.Equal(t, "0", c.V)
}

func TestSetCellCachedValueString(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "A1", `CONCATENATE("hello"," ","world")`))
	assert.NoError(t, f.setCellCachedValue("Sheet1", "A1", newStringFormulaArg("hello world")))

	c := cellXML(t, f, "Sheet1", "A1")
	assert.NotNil(t, c.F)
	assert.Equal(t, "str", c.T, "string result of formula must be inline, not shared")
	assert.Equal(t, "hello world", c.V)
	assert.Nil(t, c.IS, "inline-string remnant from prior cache must be cleared")
}

func TestSetCellCachedValueError(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "1/0"))
	errArg := newErrorFormulaArg(formulaErrorDIV, formulaErrorDIV)
	assert.NoError(t, f.setCellCachedValue("Sheet1", "A1", errArg))

	c := cellXML(t, f, "Sheet1", "A1")
	assert.NotNil(t, c.F)
	assert.Equal(t, "e", c.T)
	assert.Equal(t, "#DIV/0!", c.V)
}

func TestSetCellCachedValueMatrixCollapses(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "SUM(B1:B3)"))
	matrix := newMatrixFormulaArg([][]formulaArg{{newNumberFormulaArg(7)}})
	assert.NoError(t, f.setCellCachedValue("Sheet1", "A1", matrix))

	c := cellXML(t, f, "Sheet1", "A1")
	assert.Equal(t, "", c.T)
	assert.Equal(t, "7", c.V)
}

func TestSetCellCachedValueNoFormula(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellInt("Sheet1", "A1", 5))
	err := f.setCellCachedValue("Sheet1", "A1", newNumberFormulaArg(99))
	assert.ErrorIs(t, err, ErrCellNoFormula)
}

func TestSetCellCachedValueSharedFormulaChild(t *testing.T) {
	// Construct a shared-formula group G40:K40 with master G40 holding
	// SUM(G9:G39). One SetCellFormula call with Type=shared + Ref
	// registers the master and seeds the children; setSharedFormula
	// fills each child's `<f t="shared" si="N"/>` entry automatically.
	f := NewFile()
	ref, master := "G40:K40", "SUM(G9:G39)"
	sharedType := STCellFormulaTypeShared
	assert.NoError(t, f.SetCellFormula("Sheet1", "G40", master, FormulaOpts{
		Type: &sharedType,
		Ref:  &ref,
	}))

	master40 := cellXML(t, f, "Sheet1", "G40")
	assert.NotNil(t, master40.F.Si, "shared master carries si")
	sharedIdx := *master40.F.Si

	// Seed each child with a distinct cached value so we can detect if
	// one write bleeds into another.
	children := []string{"H40", "I40", "J40", "K40"}
	for i, child := range children {
		assert.NoError(t, f.setCellCachedValue("Sheet1", child, newNumberFormulaArg(float64(i+10))))
	}
	assert.NoError(t, f.setCellCachedValue("Sheet1", "G40", newNumberFormulaArg(1)))

	master40 = cellXML(t, f, "Sheet1", "G40")
	assert.NotNil(t, master40.F, "master formula must remain")
	assert.Equal(t, master, master40.F.Content, "master formula body unchanged")
	assert.Equal(t, ref, master40.F.Ref, "master ref preserved")
	assert.NotNil(t, master40.F.Si, "master si preserved")
	assert.Equal(t, sharedIdx, *master40.F.Si)
	assert.Equal(t, "1", master40.V)

	// Each child keeps its shared-ref formula element (empty body,
	// shared type, matching si) and its own cache. No child's cache
	// was disturbed by writes to another child or the master.
	for i, child := range children {
		c := cellXML(t, f, "Sheet1", child)
		assert.NotNil(t, c.F, child+": shared-formula element must remain")
		assert.Equal(t, "", c.F.Content, child+": shared child has no inline body")
		assert.Equal(t, sharedType, c.F.T, child+": shared type preserved")
		assert.Equal(t, "", c.F.Ref, child+": only the master carries a ref")
		assert.NotNil(t, c.F.Si, child+": shared index pointer preserved")
		assert.Equal(t, sharedIdx, *c.F.Si, child+": shared index value preserved")
		assert.Equal(t, strconv.Itoa(i+10), c.V, child+": cache matches prior write")
	}
}

// TestSetCellCachedValueRoundTrip closes the loop: save the workbook
// with populated caches, re-open, and confirm the formula + cached
// value both survive serialisation.
func TestSetCellCachedValueRoundTrip(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "SUM(1,2,3,4,5)"))
	assert.NoError(t, f.setCellCachedValue("Sheet1", "A1", newNumberFormulaArg(15)))

	buf, err := f.WriteToBuffer()
	assert.NoError(t, err)
	f2, err := OpenReader(strings.NewReader(buf.String()))
	assert.NoError(t, err)
	defer func() { _ = f2.Close() }()

	formula, err := f2.GetCellFormula("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "SUM(1,2,3,4,5)", formula)

	value, err := f2.GetCellValue("Sheet1", "A1", Options{RawCellValue: true})
	assert.NoError(t, err)
	assert.Equal(t, "15", value)
}
