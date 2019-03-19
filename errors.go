package excelize

import (
	"fmt"
)

func newInvalidColumnNameError(col string) error {
	return fmt.Errorf("invalid column name %q", col)
}

func newInvalidRowNumberError(row int) error {
	return fmt.Errorf("invalid row number %d", row)
}

func newInvalidCellNameError(cell string) error {
	return fmt.Errorf("invalid cell name %q", cell)
}
