package excelize

import (
	"fmt"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
)

type dateTest struct {
	ExcelValue float64
	GoValue    time.Time
}

func TestTimeToExcelTime(t *testing.T) {
	trueExpectedInputList := []dateTest{
		{0.0, time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)},
		{25569.0, time.Unix(0, 0)},
		{43269.0, time.Date(2018, 6, 18, 0, 0, 0, 0, time.UTC)},
		{401769.0, time.Date(3000, 1, 1, 0, 0, 0, 0, time.UTC)},
	}

	for i, test := range trueExpectedInputList {
		t.Run(fmt.Sprintf("TestData%d", i+1), func(t *testing.T) {
			assert.Equal(t, test.ExcelValue, timeToExcelTime(test.GoValue))
		})
	}
}

func TestTimeFromExcelTime(t *testing.T) {
	trueExpectedInputList := []dateTest{
		{0.0, time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)},
		{60.0, time.Date(1900, 2, 28, 0, 0, 0, 0, time.UTC)},
		{61.0, time.Date(1900, 3, 1, 0, 0, 0, 0, time.UTC)},
		{41275.0, time.Date(2013, 1, 1, 0, 0, 0, 0, time.UTC)},
		{401769.0, time.Date(3000, 1, 1, 0, 0, 0, 0, time.UTC)},
	}

	for i, test := range trueExpectedInputList {
		t.Run(fmt.Sprintf("TestData%d", i+1), func(t *testing.T) {
			assert.Equal(t, test.GoValue, timeFromExcelTime(test.ExcelValue, false))
		})
	}
}
