package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestAxisLowerOrEqualThanIsTrue(t *testing.T) {
	trueExpectedInputList := [][2]string{
		{"A", "B"},
		{"A", "AA"},
		{"B", "AA"},
		{"BC", "ABCD"},
		{"1", "2"},
		{"2", "11"},
	}

	for i, trueExpectedInput := range trueExpectedInputList {
		t.Run(fmt.Sprintf("TestData%d", i), func(t *testing.T) {
			assert.True(t, axisLowerOrEqualThan(trueExpectedInput[0], trueExpectedInput[1]))
		})
	}
}

func TestAxisLowerOrEqualThanIsFalse(t *testing.T) {
	falseExpectedInputList := [][2]string{
		{"B", "A"},
		{"AA", "A"},
		{"AA", "B"},
		{"ABCD", "AB"},
		{"2", "1"},
		{"11", "2"},
	}

	for i, falseExpectedInput := range falseExpectedInputList {
		t.Run(fmt.Sprintf("TestData%d", i), func(t *testing.T) {
			assert.False(t, axisLowerOrEqualThan(falseExpectedInput[0], falseExpectedInput[1]))
		})
	}
}

func TestGetCellColRow(t *testing.T) {
	cellExpectedColRowList := [][3]string{
		{"C220", "C", "220"},
		{"aaef42", "aaef", "42"},
		{"bonjour", "bonjour", ""},
		{"59", "", "59"},
		{"", "", ""},
	}

	for i, test := range cellExpectedColRowList {
		t.Run(fmt.Sprintf("TestData%d", i), func(t *testing.T) {
			col, row := getCellColRow(test[0])
			assert.Equal(t, test[1], col, "Unexpected col")
			assert.Equal(t, test[2], row, "Unexpected row")
		})
	}
}
