package excelize

import (
    "testing"
    "time"
)

type dateTest struct {
    ExcelValue float64
    GoValue    time.Time
}

func TestTimeToExcelTime(t *testing.T) {
    trueExpectedInputList := []dateTest {
        {0.0, time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)},
        {25569.0, time.Unix(0, 0)},
        {43269.0, time.Date(2018, 6, 18, 0, 0, 0, 0, time.UTC)},
        {401769.0, time.Date(3000, 1, 1, 0, 0, 0, 0, time.UTC)},
    }

    for _, test := range trueExpectedInputList {
        if test.ExcelValue != timeToExcelTime(test.GoValue) {
            t.Fatalf("Expected %v from %v = true, got %v\n", test.ExcelValue, test.GoValue, timeToExcelTime(test.GoValue))
        }
    }
}

func TestTimeFromExcelTime(t *testing.T) {
    trueExpectedInputList := []dateTest {
        {0.0, time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)},
        {60.0, time.Date(1900, 2, 28, 0, 0, 0, 0, time.UTC)},
        {61.0, time.Date(1900, 3, 1, 0, 0, 0, 0, time.UTC)},
        {41275.0, time.Date(2013, 1, 1, 0, 0, 0, 0, time.UTC)},
        {401769.0, time.Date(3000, 1, 1, 0, 0, 0, 0, time.UTC)},
    }

    for _, test := range trueExpectedInputList {
        if test.GoValue != timeFromExcelTime(test.ExcelValue, false) {
            t.Fatalf("Expected %v from %v = true, got %v\n", test.GoValue, test.ExcelValue, timeFromExcelTime(test.ExcelValue, false))
        }
    }
}
