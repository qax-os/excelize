package excelize

import "testing"

func TestCheckCellInArea(t *testing.T) {
	expectedTrueCellInAreaList := [][2]string{
		{"c2", "A1:AAZ32"},
		{"AA0", "Z0:AB1"},
		{"B9", "A1:B9"},
		{"C2", "C2:C2"},
	}

	for _, expectedTrueCellInArea := range expectedTrueCellInAreaList {
		cell := expectedTrueCellInArea[0]
		area := expectedTrueCellInArea[1]

		cellInArea := checkCellInArea(cell, area)

		if !cellInArea {
			t.Fatalf("Expected cell %v to be in area %v, got false\n", cell, area)
		}
	}

	expectedFalseCellInAreaList := [][2]string{
		{"c2", "A4:AAZ32"},
		{"C4", "D6:A1"}, // weird case, but you never know
		{"AEF42", "BZ40:AEF41"},
	}

	for _, expectedFalseCellInArea := range expectedFalseCellInAreaList {
		cell := expectedFalseCellInArea[0]
		area := expectedFalseCellInArea[1]

		cellInArea := checkCellInArea(cell, area)

		if cellInArea {
			t.Fatalf("Expected cell %v not to be inside of area %v, but got true\n", cell, area)
		}
	}
}
