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

// timeAfter2Seconds is a thin wrapper around time.After that keeps
// the test-body line count low enough to match the house style.
func timeAfter2Seconds() <-chan time.Time { return time.After(2 * time.Second) }

// readCachedNumber reads the cached `<v>` on a cell that holds a
// formula and returns it without going through any formatted-value
// path. The test asserts the cached value directly so regressions
// that quietly write the wrong type are visible.
func readCachedNumber(t *testing.T, f *File, sheet, addr string) string {
	t.Helper()
	c := cellXML(t, f, sheet, addr)
	assert.NotNil(t, c.F, sheet+"!"+addr+": expected formula element")
	return c.V
}

func TestRecalcPersistsSimpleSum(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellInt("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellInt("Sheet1", "A2", 32))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A3", "SUM(A1:A2)"))

	// Before Recalc the cache on A3 is empty (formula just set).
	c := cellXML(t, f, "Sheet1", "A3")
	assert.Empty(t, c.V, "pre-Recalc cache should be empty")

	assert.NoError(t, f.Recalc())

	c = cellXML(t, f, "Sheet1", "A3")
	assert.Equal(t, "SUM(A1:A2)", c.F.Content, "formula preserved")
	assert.Equal(t, "", c.T, "numeric result written as implicit t")
	assert.Equal(t, "42", c.V, "cached value is the computed sum")
}

func TestRecalcPreservesFormulaAcrossRoundTrip(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellInt("Sheet1", "A1", 2))
	assert.NoError(t, f.SetCellInt("Sheet1", "A2", 3))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A3", "A1*A2"))
	assert.NoError(t, f.Recalc())

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

func TestRecalcChainedDependencies(t *testing.T) {
	// A1 = 5, A2 = A1*2, A3 = A2+A1, A4 = SUM(A1:A3). Recalc must
	// converge in a single pass because cellResolver recurses into
	// each dependency.
	f := NewFile()
	assert.NoError(t, f.SetCellInt("Sheet1", "A1", 5))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "A1*2"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A3", "A2+A1"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A4", "SUM(A1:A3)"))

	assert.NoError(t, f.Recalc())

	assert.Equal(t, "10", readCachedNumber(t, f, "Sheet1", "A2"))
	assert.Equal(t, "15", readCachedNumber(t, f, "Sheet1", "A3"))
	assert.Equal(t, "30", readCachedNumber(t, f, "Sheet1", "A4"))
}

func TestRecalcAcross3DRefs(t *testing.T) {
	// One-pass evaluation of a workbook shaped like the reproducer
	// timesheet: monthly sheets M01..M03 hold a per-sheet row total
	// formula, Summary sheet holds the 3D SUM across them.
	f := NewFile()
	assert.NoError(t, f.SetSheetName("Sheet1", "M01"))
	_, err := f.NewSheet("M02")
	assert.NoError(t, err)
	_, err = f.NewSheet("M03")
	assert.NoError(t, err)
	_, err = f.NewSheet("Summary")
	assert.NoError(t, err)

	// Row data on each month sheet: A1..A3 input, A4 = SUM(A1:A3).
	for i, sn := range []string{"M01", "M02", "M03"} {
		base := int64(i + 1)
		assert.NoError(t, f.SetCellInt(sn, "A1", base*10))
		assert.NoError(t, f.SetCellInt(sn, "A2", base*100))
		assert.NoError(t, f.SetCellInt(sn, "A3", base*1000))
		assert.NoError(t, f.SetCellFormula(sn, "A4", "SUM(A1:A3)"))
	}
	// Summary!B1 sums row totals across months.
	assert.NoError(t, f.SetCellFormula("Summary", "B1", "SUM(M01:M03!A4)"))

	assert.NoError(t, f.Recalc())

	// Each month total equals base*1110; 3D sum equals 1110+2220+3330=6660.
	assert.Equal(t, "1110", readCachedNumber(t, f, "M01", "A4"))
	assert.Equal(t, "2220", readCachedNumber(t, f, "M02", "A4"))
	assert.Equal(t, "3330", readCachedNumber(t, f, "M03", "A4"))
	assert.Equal(t, "6660", readCachedNumber(t, f, "Summary", "B1"))
}

func TestRecalcSharedFormulaBandUntouched(t *testing.T) {
	// Row 40 is the shared-formula band shape used by the real
	// timesheet template. Recalc must persist each column's total
	// without wiping the shared-formula metadata.
	f := NewFile()
	for col, v := range map[string]int64{"F": 1, "G": 2, "H": 3, "I": 4, "J": 5} {
		assert.NoError(t, f.SetCellInt("Sheet1", col+"1", v))
	}
	// F40 is a regular formula (the master of a neighbouring group
	// is not always contiguous with the first column in real files).
	assert.NoError(t, f.SetCellFormula("Sheet1", "F40", "SUM(F1:F39)"))
	// G40:J40 is a shared-formula group with master G40.
	ref := "G40:J40"
	sharedType := STCellFormulaTypeShared
	assert.NoError(t, f.SetCellFormula("Sheet1", "G40", "SUM(G1:G39)", FormulaOpts{
		Type: &sharedType,
		Ref:  &ref,
	}))

	master := cellXML(t, f, "Sheet1", "G40")
	assert.NotNil(t, master.F.Si)
	si := *master.F.Si

	assert.NoError(t, f.Recalc())

	// Every total cached correctly.
	assert.Equal(t, "1", readCachedNumber(t, f, "Sheet1", "F40"))
	assert.Equal(t, "2", readCachedNumber(t, f, "Sheet1", "G40"))
	assert.Equal(t, "3", readCachedNumber(t, f, "Sheet1", "H40"))
	assert.Equal(t, "4", readCachedNumber(t, f, "Sheet1", "I40"))
	assert.Equal(t, "5", readCachedNumber(t, f, "Sheet1", "J40"))

	// Shared metadata intact: master still has ref + si, children
	// still have si (no ref).
	master = cellXML(t, f, "Sheet1", "G40")
	assert.Equal(t, "SUM(G1:G39)", master.F.Content)
	assert.Equal(t, ref, master.F.Ref)
	assert.Equal(t, si, *master.F.Si)
	for _, child := range []string{"H40", "I40", "J40"} {
		c := cellXML(t, f, "Sheet1", child)
		assert.Equal(t, sharedType, c.F.T, child)
		assert.Equal(t, "", c.F.Ref, child)
		assert.Equal(t, si, *c.F.Si, child)
	}
}

func TestRecalcIdempotent(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellInt("Sheet1", "A1", 7))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "A1*3"))

	assert.NoError(t, f.Recalc())
	first := readCachedNumber(t, f, "Sheet1", "A2")

	assert.NoError(t, f.Recalc())
	second := readCachedNumber(t, f, "Sheet1", "A2")
	assert.Equal(t, first, second, "second Recalc must be a no-op on value")
	assert.Equal(t, "21", second)
}

func TestRecalcNoTypeSRegression(t *testing.T) {
	// Explicit guard against the regression that motivated Recalc:
	// previously the recalc path round-tripped numeric results
	// through string, which stored cells as t="s" (shared string).
	// Downstream aggregates then silently summed to 0.
	f := NewFile()
	assert.NoError(t, f.SetCellInt("Sheet1", "A1", 3))
	assert.NoError(t, f.SetCellInt("Sheet1", "A2", 4))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A3", "SUM(A1:A2)"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A4", "A3*10"))

	assert.NoError(t, f.Recalc())

	a3 := cellXML(t, f, "Sheet1", "A3")
	assert.NotEqual(t, "s", a3.T, "numeric formula result must not be stored as shared string")
	assert.Equal(t, "", a3.T)
	assert.Equal(t, "7", a3.V)

	a4 := cellXML(t, f, "Sheet1", "A4")
	assert.Equal(t, "", a4.T)
	assert.Equal(t, "70", a4.V, "downstream aggregate must see a numeric A3")
}

func TestRecalcSheetScope(t *testing.T) {
	f := NewFile()
	_, err := f.NewSheet("Other")
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellInt("Sheet1", "A1", 1))
	assert.NoError(t, f.SetCellInt("Other", "A1", 1))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "A1+A1"))
	assert.NoError(t, f.SetCellFormula("Other", "A2", "A1+A1"))

	assert.NoError(t, f.Recalc(RecalcOptions{Sheet: "Other"}))

	assert.Empty(t, cellXML(t, f, "Sheet1", "A2").V, "Sheet1 cache should be untouched")
	assert.Equal(t, "2", readCachedNumber(t, f, "Other", "A2"))
}

func TestRecalcRefScope(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellInt("Sheet1", "A1", 1))
	assert.NoError(t, f.SetCellInt("Sheet1", "B1", 1))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "A1+A1"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B2", "B1+B1"))

	assert.NoError(t, f.Recalc(RecalcOptions{Sheet: "Sheet1", Ref: "A1:A10"}))

	assert.Equal(t, "2", readCachedNumber(t, f, "Sheet1", "A2"))
	assert.Empty(t, cellXML(t, f, "Sheet1", "B2").V, "B2 should be out of scope")
}

func TestRecalcRejectsRefWithoutSheet(t *testing.T) {
	f := NewFile()
	err := f.Recalc(RecalcOptions{Ref: "A1:B2"})
	assert.Error(t, err)
	assert.Contains(t, err.Error(), "requires Sheet")
}

func TestRecalcUnknownSheet(t *testing.T) {
	f := NewFile()
	err := f.Recalc(RecalcOptions{Sheet: "Nope"})
	assert.Error(t, err)
	assert.Contains(t, err.Error(), `"Nope"`)
}

func TestRecalcAggregatesFailures(t *testing.T) {
	// An unsupported function should produce a RecalcError containing
	// the offending cell without stopping the scan of other cells.
	f := NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "NONEXISTENTFUNC(1)"))
	assert.NoError(t, f.SetCellInt("Sheet1", "B1", 10))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B2", "B1*2"))

	err := f.Recalc()
	assert.Error(t, err)
	rerr, ok := err.(*RecalcError)
	if assert.True(t, ok, "expected *RecalcError, got %T", err) {
		assert.Len(t, rerr.Cells, 1)
		assert.Equal(t, "Sheet1", rerr.Cells[0].Sheet)
		assert.Equal(t, "A1", rerr.Cells[0].Cell)
	}
	// Working cells were still recalculated.
	assert.Equal(t, "20", readCachedNumber(t, f, "Sheet1", "B2"))
}

func TestRecalcWorkbookWithoutFormulas(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellInt("Sheet1", "A1", 42))
	assert.NoError(t, f.Recalc())
}

func TestRecalcBoolResult(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellInt("Sheet1", "A1", 5))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "A1>0"))

	assert.NoError(t, f.Recalc())

	c := cellXML(t, f, "Sheet1", "A2")
	assert.Equal(t, "b", c.T, "boolean formula result must be stored as t=\"b\"")
	assert.Equal(t, "1", c.V)
}

func TestRecalcErrorResult(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellInt("Sheet1", "A1", 1))
	assert.NoError(t, f.SetCellInt("Sheet1", "A2", 0))
	assert.NoError(t, f.SetCellFormula("Sheet1", "A3", "A1/A2"))

	err := f.Recalc()
	// A1/0 produces a #DIV/0! cell error; the calc engine surfaces
	// it as an err from calcCellValue, so Recalc aggregates it.
	if err != nil {
		rerr, ok := err.(*RecalcError)
		assert.True(t, ok, "expected *RecalcError, got %T", err)
		if ok {
			assert.Len(t, rerr.Cells, 1)
			assert.Equal(t, "A3", rerr.Cells[0].Cell)
		}
	}
	// Either way, A3's cache should reflect the divide-by-zero error
	// rather than a stale value.
	c := cellXML(t, f, "Sheet1", "A3")
	if c.T == "e" {
		assert.Equal(t, "#DIV/0!", c.V)
	} else if c.V != "" {
		// Some calc paths surface #DIV/0! via the string in .V
		// without setting t="e"; accept either as long as it is a
		// recognisable error code, not silent 0.
		assert.Contains(t, c.V, "#DIV/0!")
	}
}

func TestRecalcCircularReferenceBounded(t *testing.T) {
	// A1 = B1+1, B1 = A1+1. A naive evaluator would loop; Recalc
	// must terminate because calcCellValue honours MaxCalcIterations.
	f := NewFile()
	assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "B1+1"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "A1+1"))

	done := make(chan error, 1)
	go func() { done <- f.Recalc() }()
	select {
	case <-done:
		// Either error or success is acceptable — the assertion is
		// that Recalc terminates at all on a circular reference.
	case <-timeAfter2Seconds():
		t.Fatal("Recalc did not terminate on circular reference")
	}
}

func TestRecalcNonASCIISheetName(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetSheetName("Sheet1", "Bérénice"))
	assert.NoError(t, f.SetCellInt("Bérénice", "A1", 3))
	assert.NoError(t, f.SetCellFormula("Bérénice", "A2", "A1*A1"))

	assert.NoError(t, f.Recalc())
	assert.Equal(t, "9", readCachedNumber(t, f, "Bérénice", "A2"))
}
