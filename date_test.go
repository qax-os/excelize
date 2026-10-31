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
	{0, time.Date(1899, time.December, 30, 0, 0, 0, 0, time.UTC)},
	{25569, time.Unix(0, 0).UTC()},
	// Expected values extracted from real spreadsheet
	{1, time.Date(1900, time.January, 1, 0, 0, 0, 0, time.UTC)},
	{1.0000115740740740, time.Date(1900, time.January, 1, 0, 0, 1, 0, time.UTC)},
	{1.0006944444444446, time.Date(1900, time.January, 1, 0, 1, 0, 0, time.UTC)},
	{1.0416666666666667, time.Date(1900, time.January, 1, 1, 0, 0, 0, time.UTC)},
	{2, time.Date(1900, time.January, 2, 0, 0, 0, 0, time.UTC)},
	{43269, time.Date(2018, time.June, 18, 0, 0, 0, 0, time.UTC)},
	{43542.611111111109, time.Date(2019, time.March, 18, 14, 40, 0, 0, time.UTC)},
	{401769, time.Date(3000, time.January, 1, 0, 0, 0, 0, time.UTC)},
}

var excelTimeInputList = []dateTest{
	{0, time.Date(1899, time.December, 30, 0, 0, 0, 0, time.UTC)},
	{0.5, time.Date(1899, time.December, 30, 12, 0, 0, 0, time.UTC)},
	{1, time.Date(1900, time.January, 1, 0, 0, 0, 0, time.UTC)},
	{1.25, time.Date(1900, time.January, 1, 6, 0, 0, 0, time.UTC)},
	{31, time.Date(1900, time.January, 31, 0, 0, 0, 0, time.UTC)},
	{32, time.Date(1900, time.February, 1, 0, 0, 0, 0, time.UTC)},
	{401769, time.Date(3000, time.January, 1, 0, 0, 0, 0, time.UTC)},
	{41275, time.Date(2013, time.January, 1, 0, 0, 0, 0, time.UTC)},
	{44450.3333333333, time.Date(2021, time.September, 11, 8, 0, 0, 0, time.UTC)},
	{59, time.Date(1900, time.February, 28, 0, 0, 0, 0, time.UTC)},
	{59.75, time.Date(1900, time.February, 28, 18, 0, 0, 0, time.UTC)},
	{60, time.Date(1900, time.February, 28, 0, 0, 0, 0, time.UTC)},
	{60.5, time.Date(1900, time.February, 28, 12, 0, 0, 0, time.UTC)},
	{61, time.Date(1900, time.March, 1, 0, 0, 0, 0, time.UTC)},
	{61.5, time.Date(1900, time.March, 1, 12, 0, 0, 0, time.UTC)},
	{62, time.Date(1900, time.March, 2, 0, 0, 0, 0, time.UTC)},
}

func TestTimeToExcelTime(t *testing.T) {
	for i, test := range trueExpectedDateList {
		t.Run(fmt.Sprintf("TestData%d", i+1), func(t *testing.T) {
			excelTime, err := timeToExcelTime(test.GoValue, false)
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
			_, err := timeToExcelTime(test.GoValue.In(location), false)
			assert.NoError(t, err)
		})
	}
}

func TestTimeFromExcelTime(t *testing.T) {
	for hour := 0; hour < 24; hour++ {
		for minVal := 0; minVal < 60; minVal++ {
			for sec := 0; sec < 60; sec++ {
				date := time.Date(2021, time.December, 30, hour, minVal, sec, 0, time.UTC)
				// Test use 1900 date system
				excel1900Time, err := timeToExcelTime(date, false)
				assert.NoError(t, err)
				date1900Out := timeFromExcelTime(excel1900Time, false)
				assert.EqualValues(t, hour, date1900Out.Hour())
				assert.EqualValues(t, minVal, date1900Out.Minute())
				assert.EqualValues(t, sec, date1900Out.Second())
				// Test use 1904 date system
				excel1904Time, err := timeToExcelTime(date, true)
				assert.NoError(t, err)
				date1904Out := timeFromExcelTime(excel1904Time, true)
				assert.EqualValues(t, hour, date1904Out.Hour())
				assert.EqualValues(t, minVal, date1904Out.Minute())
				assert.EqualValues(t, sec, date1904Out.Second())
			}
		}
	}
}

func TestTimeFromExcelTime_1904(t *testing.T) {
	julianDays, julianFraction := shiftJulianToNoon(1, -0.6)
	assert.Equal(t, julianDays, 0.0)
	assert.Equal(t, julianFraction, 0.9)
	julianDays, julianFraction = shiftJulianToNoon(1, 0.1)
	assert.Equal(t, julianDays, 1.0)
	assert.Equal(t, julianFraction, 0.6)
	assert.Equal(t, timeFromExcelTime(61, true), time.Date(1904, time.March, 2, 0, 0, 0, 0, time.UTC))
	assert.Equal(t, timeFromExcelTime(62, true), time.Date(1904, time.March, 3, 0, 0, 0, 0, time.UTC))
}

func TestExcelDateToTime(t *testing.T) {
	// Check normal case
	for _, test := range excelTimeInputList {
		timeValue, err := ExcelDateToTime(test.ExcelValue, false)
		assert.NoError(t, err)
		assert.Equal(t, test.GoValue, timeValue, "excelDate: %v", test.ExcelValue)
	}
	// Check error case
	_, err := ExcelDateToTime(-1, false)
	assert.EqualError(t, err, newInvalidExcelDateError(-1).Error())
}
