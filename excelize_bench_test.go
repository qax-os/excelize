package excelize

import (
	"testing"
	"time"
	"math/rand"
	"strconv"
)

const cols = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

func init() {
	rand.Seed(time.Now().UTC().UnixNano())
}

func BenchmarkOldSetValue(b *testing.B) {
	xlsx, err := OpenFile("./test/Workbook1.xlsx")
	rand.Seed( time.Now().UnixNano())
	if err != nil {
		b.Log(err)
	}

	maxCol := len(cols)
	maxRow := maxCol

	for n := 0; n < b.N; n++ {
		col := cols[rand.Intn(maxCol)]
		row := strconv.Itoa(1 + rand.Intn(maxRow))
		axis := string(col)+ row
		value := rand.Intn(100)
		xlsx.SetCellValue("Sheet2", axis, value)
	}
}

func BenchmarkOldSetRangeStyle(b *testing.B) {
	xlsx, err := OpenFile("./test/Workbook1.xlsx")
	rand.Seed( time.Now().UnixNano())
	if err != nil {
		b.Log(err)
	}

	var style int
	style, err = xlsx.NewStyle(`{"alignment":{"horizontal":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":45,"vertical":"top","wrap_text":true}}`)
	if err != nil {
		b.Log(err)
	}

	maxCol := len(cols)
	maxRow := maxCol

	for n := 0; n < b.N; n++ {
		col1 := cols[rand.Intn(maxCol)]
		row1 := strconv.Itoa(1 + rand.Intn(maxRow))
		axis1 := string(col1)+ row1

		col2 := cols[rand.Intn(maxCol)]
		row2 := strconv.Itoa(1 + rand.Intn(maxRow))
		axis2 := string(col2)+ row2

		xlsx.SetCellStyle("Sheet2", axis1, axis2, style)
	}
}

func BenchmarkOldSetCellStyle(b *testing.B) {
	xlsx, err := OpenFile("./test/Workbook1.xlsx")
	rand.Seed( time.Now().UnixNano())
	if err != nil {
		b.Log(err)
	}

	var style int
	style, err = xlsx.NewStyle(`{"alignment":{"horizontal":"center","ident":1,"justify_last_line":true,"reading_order":0,"relative_indent":1,"shrink_to_fit":true,"text_rotation":45,"vertical":"top","wrap_text":true}}`)
	if err != nil {
		b.Log(err)
	}

	maxCol := len(cols)
	maxRow := maxCol

	for n := 0; n < b.N; n++ {
		col := cols[rand.Intn(maxCol)]
		row := strconv.Itoa(1 + rand.Intn(maxRow))
		axis := string(col)+ row

		xlsx.SetCellStyle("Sheet2", axis, axis, style)
	}
}
