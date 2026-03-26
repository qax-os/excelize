package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestNewInvalidColNameError(t *testing.T) {
	assert.EqualError(t, newInvalidColumnNameError("A"), "invalid column name \"A\"")
	assert.EqualError(t, newInvalidColumnNameError(""), "invalid column name \"\"")
}

func TestNewInvalidRowNumberError(t *testing.T) {
	assert.EqualError(t, newInvalidRowNumberError(0), "invalid row number 0")
}

func TestNewInvalidCellNameError(t *testing.T) {
	assert.EqualError(t, newInvalidCellNameError("A"), "invalid cell name \"A\"")
	assert.EqualError(t, newInvalidCellNameError(""), "invalid cell name \"\"")
}

func TestNewInvalidExcelDateError(t *testing.T) {
	assert.EqualError(t, newInvalidExcelDateError(-1), "invalid date value -1.000000, negative values are not supported")
}
