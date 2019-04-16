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

var trueExpectedDateList = []dateTest{
	{0.0000000000000000, time.Date(1899, time.December, 30, 0, 0, 0, 0, time.UTC)},
	{25569.000000000000, time.Unix(0, 0).UTC()},

	// Expected values extracted from real xlsx file
	{1.0000000000000000, time.Date(1900, time.January, 1, 0, 0, 0, 0, time.UTC)},
	{1.0000115740740740, time.Date(1900, time.January, 1, 0, 0, 1, 0, time.UTC)},
	{1.0006944444444446, time.Date(1900, time.January, 1, 0, 1, 0, 0, time.UTC)},
	{1.0416666666666667, time.Date(1900, time.January, 1, 1, 0, 0, 0, time.UTC)},
	{2.0000000000000000, time.Date(1900, time.January, 2, 0, 0, 0, 0, time.UTC)},
	{43269.000000000000, time.Date(2018, time.June, 18, 0, 0, 0, 0, time.UTC)},
	{43542.611111111109, time.Date(2019, time.March, 18, 14, 40, 0, 0, time.UTC)},
	{401769.00000000000, time.Date(3000, time.January, 1, 0, 0, 0, 0, time.UTC)},
}

func TestTimeToExcelTime(t *testing.T) {
	for i, test := range trueExpectedDateList {
		t.Run(fmt.Sprintf("TestData%d", i+1), func(t *testing.T) {
			excelTime, err := timeToExcelTime(test.GoValue)
			assert.NoError(t, err)
			assert.Equalf(t, test.ExcelValue, excelTime,
				"Time: %s", test.GoValue.String())
		})
	}
}

func TestTimeToExcelTime_Timezone(t *testing.T) {
	location, err := time.LoadLocation("America/Los_Angeles")
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	for i, test := range trueExpectedDateList {
		t.Run(fmt.Sprintf("TestData%d", i+1), func(t *testing.T) {
			_, err := timeToExcelTime(test.GoValue.In(location))
			assert.EqualError(t, err, "only UTC time expected")
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

func TestTimeFromExcelTime_1904(t *testing.T) {
	_, _ = shiftJulianToNoon(1, -0.6)
	timeFromExcelTime(61, true)
	timeFromExcelTime(62, true)
}
