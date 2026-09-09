package excelize

import (
	"testing"

	"runtime/debug"
)

// Attacker model: an untrusted workbook declares two single-cell dynamic
// array formulas that reference each other through ANCHORARRAY (a legal
// Excel construct):
//
//	A1 (array, ref A1:A1): _xlfn.ANCHORARRAY($B$1)
//	B1 (array, ref B1:B1): _xlfn.ANCHORARRAY($A$1)
//
// The circular-reference control in cellResolver (iterations map +
// MaxCalcIterations) lives on ONE calcContext. ANCHORARRAY evaluates spill
// range cells via the EXPORTED CalcCellValue (calc.go:15149), which builds a
// FRESH calcContext with a fresh entry marker and a fresh iterations budget
// (calc.go:896-900). During a pure cycle no evaluation ever completes, so
// nothing is ever written to formulaArgCache / calcRawCache either, and each
// nested call re-enters with a full budget -> unbounded recursion -> fatal
// "stack overflow" (process crash, unrecoverable in Go).
//
// Before the fix the process dies with a runtime fatal error; after the fix
// the calculation terminates and this test prints
// S58_ANCHORARRAY_CYCLE_TERMINATED.
func TestS58AnchorArrayMutualCycleUnboundedRecursion(t *testing.T) {
	// Shrink the stack budget so the (unbounded) recursion reaches the fatal
	// error quickly. Does not change the semantics of the bug.
	debug.SetMaxStack(64 << 20)

	f := NewFile()
	ftA, refA := STCellFormulaTypeArray, "A1:A1"
	ftB, refB := STCellFormulaTypeArray, "B1:B1"
	if err := f.SetCellFormula("Sheet1", "A1", "_xlfn.ANCHORARRAY($B$1)",
		FormulaOpts{Ref: &refA, Type: &ftA}); err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", "_xlfn.ANCHORARRAY($A$1)",
		FormulaOpts{Ref: &refB, Type: &ftB}); err != nil {
		t.Fatal(err)
	}

	t.Log("S58_ANCHORARRAY_RECURSION_UNBOUNDED: entering mutually recursive ANCHORARRAY formulas")
	_, _ = f.CalcCellValue("Sheet1", "A1")
	t.Log("S58_ANCHORARRAY_CYCLE_TERMINATED")
}
