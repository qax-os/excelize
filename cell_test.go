package excelize

import (
	"fmt"
	_ "image/jpeg"
	"math"
	"os"
	"path/filepath"
	"reflect"
	"strconv"
	"strings"
	"sync"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
)

func TestConcurrency(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)
	wg := new(sync.WaitGroup)
	for i := 1; i <= 5; i++ {
		wg.Add(1)
		go func(val int, t *testing.T) {
			// Concurrency set cell value
			assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("A%d", val), val))
			assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("B%d", val), strconv.Itoa(val)))
			// Concurrency get cell value
			_, err := f.GetCellValue("Sheet1", fmt.Sprintf("A%d", val))
			assert.NoError(t, err)
			// Concurrency set rows
			assert.NoError(t, f.SetSheetRow("Sheet1", "B6", &[]interface{}{
				" Hello",
				[]byte("World"), 42, int8(1<<8/2 - 1), int16(1<<16/2 - 1), int32(1<<32/2 - 1),
				int64(1<<32/2 - 1), float32(42.65418), -42.65418, float32(42), float64(42),
				uint(1<<32 - 1), uint8(1<<8 - 1), uint16(1<<16 - 1), uint32(1<<32 - 1),
				uint64(1<<32 - 1), true, complex64(5 + 10i),
			}))
			// Concurrency create style
			style, err := f.NewStyle(&Style{Font: &Font{Color: "1265BE", Underline: "single"}})
			assert.NoError(t, err)
			// Concurrency set cell style
			assert.NoError(t, f.SetCellStyle("Sheet1", "A3", "A3", style))
			// Concurrency get cell style
			_, err = f.GetCellStyle("Sheet1", "A3")
			assert.NoError(t, err)
			// Concurrency add picture
			assert.NoError(t, f.AddPicture("Sheet1", "F21", filepath.Join("test", "images", "excel.jpg"),
				&GraphicOptions{
					OffsetX:       10,
					OffsetY:       10,
					Hyperlink:     "https://github.com/xuri/excelize",
					HyperlinkType: "External",
					Positioning:   "oneCell",
				},
			))
			// Concurrency get cell picture
			pics, err := f.GetPictures("Sheet1", "A1")
			assert.Len(t, pics, 0)
			assert.NoError(t, err)
			// Concurrency iterate rows
			rows, err := f.Rows("Sheet1")
			assert.NoError(t, err)
			for rows.Next() {
				_, err := rows.Columns()
				assert.NoError(t, err)
			}
			// Concurrency iterate columns
			cols, err := f.Cols("Sheet1")
			assert.NoError(t, err)
			for cols.Next() {
				_, err := cols.Rows()
				assert.NoError(t, err)
			}
			// Concurrency set columns style
			assert.NoError(t, f.SetColStyle("Sheet1", "C:E", style))
			// Concurrency get columns style
			styleID, err := f.GetColStyle("Sheet1", "D")
			assert.NoError(t, err)
			assert.Equal(t, style, styleID)
			// Concurrency set columns width
			assert.NoError(t, f.SetColWidth("Sheet1", "A", "B", 10))
			// Concurrency get columns width
			width, err := f.GetColWidth("Sheet1", "A")
			assert.NoError(t, err)
			assert.Equal(t, 10.0, width)
			// Concurrency set columns visible
			assert.NoError(t, f.SetColVisible("Sheet1", "A:B", true))
			// Concurrency get columns visible
			visible, err := f.GetColVisible("Sheet1", "A")
			assert.NoError(t, err)
			assert.Equal(t, true, visible)
			// Concurrency add data validation
			dv := NewDataValidation(true)
			dv.Sqref = fmt.Sprintf("A%d:B%d", val, val)
			assert.NoError(t, dv.SetRange(10, 20, DataValidationTypeWhole, DataValidationOperatorGreaterThan))
			dv.SetInput(fmt.Sprintf("title:%d", val), strconv.Itoa(val))
			assert.NoError(t, f.AddDataValidation("Sheet1", dv))
			// Concurrency delete data validation with reference sequence
			assert.NoError(t, f.DeleteDataValidation("Sheet1", dv.Sqref))
			wg.Done()
		}(i, t)
	}
	wg.Wait()
	val, err := f.GetCellValue("Sheet1", "A1")
	if err != nil {
		t.Error(err)
	}
	assert.Equal(t, "1", val)
	// Test the length of data validation
	dataValidations, err := f.GetDataValidations("Sheet1")
	assert.NoError(t, err)
	assert.Len(t, dataValidations, 0)
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestConcurrency.xlsx")))
	assert.NoError(t, f.Close())
}

func TestCheckCellInRangeRef(t *testing.T) {
	f := NewFile()
	expectedTrueCellInRangeRefList := [][2]string{
		{"c2", "A1:AAZ32"},
		{"B9", "A1:B9"},
		{"C2", "C2:C2"},
	}

	for _, expectedTrueCellInRangeRef := range expectedTrueCellInRangeRefList {
		cell := expectedTrueCellInRangeRef[0]
		reference := expectedTrueCellInRangeRef[1]
		ok, err := f.checkCellInRangeRef(cell, reference)
		assert.NoError(t, err)
		assert.Truef(t, ok,
			"Expected cell %v to be in range reference %v, got false\n", cell, reference)
	}

	expectedFalseCellInRangeRefList := [][2]string{
		{"c2", "A4:AAZ32"},
		{"C4", "D6:A1"}, // weird case, but you never know
		{"AEF42", "BZ40:AEF41"},
	}

	for _, expectedFalseCellInRangeRef := range expectedFalseCellInRangeRefList {
		cell := expectedFalseCellInRangeRef[0]
		reference := expectedFalseCellInRangeRef[1]
		ok, err := f.checkCellInRangeRef(cell, reference)
		assert.NoError(t, err)
		assert.Falsef(t, ok,
			"Expected cell %v not to be inside of range reference %v, but got true\n", cell, reference)
	}

	ok, err := f.checkCellInRangeRef("A1", "A:B")
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), err)
	assert.False(t, ok)

	ok, err = f.checkCellInRangeRef("AA0", "Z0:AB1")
	assert.Equal(t, newCellNameToCoordinatesError("AA0", newInvalidCellNameError("AA0")), err)
	assert.False(t, ok)
}

func TestSetCellFloat(t *testing.T) {
	sheet := "Sheet1"
	t.Run("with no decimal", func(t *testing.T) {
		f := NewFile()
		assert.NoError(t, f.SetCellFloat(sheet, "A1", 123.0, -1, 64))
		assert.NoError(t, f.SetCellFloat(sheet, "A2", 123.0, 1, 64))
		val, err := f.GetCellValue(sheet, "A1")
		assert.NoError(t, err)
		assert.Equal(t, "123", val, "A1 should be 123")
		val, err = f.GetCellValue(sheet, "A2")
		assert.NoError(t, err)
		assert.Equal(t, "123", val, "A2 should be 123")
	})

	t.Run("with a decimal and precision limit", func(t *testing.T) {
		f := NewFile()
		assert.NoError(t, f.SetCellFloat(sheet, "A1", 123.42, 1, 64))
		val, err := f.GetCellValue(sheet, "A1")
		assert.NoError(t, err)
		assert.Equal(t, "123.4", val, "A1 should be 123.4")
	})

	t.Run("with a decimal and no limit", func(t *testing.T) {
		f := NewFile()
		assert.NoError(t, f.SetCellFloat(sheet, "A1", 123.42, -1, 64))
		val, err := f.GetCellValue(sheet, "A1")
		assert.NoError(t, err)
		assert.Equal(t, "123.42", val, "A1 should be 123.42")
	})
	f := NewFile()
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), f.SetCellFloat(sheet, "A", 123.42, -1, 64))
	// Test set cell float data type value with invalid sheet name
	assert.Equal(t, ErrSheetNameInvalid, f.SetCellFloat("Sheet:1", "A1", 123.42, -1, 64))
}

func TestSetCellUint(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", uint8(math.MaxUint8)))
	result, err := f.GetCellValue("Sheet1", "A1")
	assert.Equal(t, "255", result)
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", uint(math.MaxUint16)))
	result, err = f.GetCellValue("Sheet1", "A1")
	assert.Equal(t, "65535", result)
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", uint(math.MaxUint32)))
	result, err = f.GetCellValue("Sheet1", "A1")
	assert.Equal(t, "4294967295", result)
	assert.NoError(t, err)
	// Test uint cell value not exists worksheet
	assert.EqualError(t, f.SetCellUint("SheetN", "A1", 1), "sheet SheetN does not exist")
	// Test uint cell value with illegal cell reference
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), f.SetCellUint("Sheet1", "A", 1))
}

func TestSetCellValuesMultiByte(t *testing.T) {
	f := NewFile()
	row := []interface{}{
		// Test set cell value with multi byte characters value
		strings.Repeat("\u4E00", TotalCellChars+1),
		// Test set cell value with XML escape characters
		strings.Repeat("<>", TotalCellChars/2),
		strings.Repeat(">", TotalCellChars-1),
		strings.Repeat(">", TotalCellChars+1),
	}
	assert.NoError(t, f.SetSheetRow("Sheet1", "A1", &row))
	// Test set cell value with XML escape characters in stream writer
	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)
	streamWriter, err := f.NewStreamWriter("Sheet2")
	assert.NoError(t, err)
	assert.NoError(t, streamWriter.SetRow("A1", row))
	assert.NoError(t, streamWriter.Flush())
	for _, sheetName := range []string{"Sheet1", "Sheet2"} {
		for cell, expected := range map[string]int{
			"A1": TotalCellChars,
			"B1": TotalCellChars - 1,
			"C1": TotalCellChars - 1,
			"D1": TotalCellChars,
		} {
			result, err := f.GetCellValue(sheetName, cell)
			assert.NoError(t, err)
			assert.Len(t, []rune(result), expected)
		}
	}
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellValuesMultiByte.xlsx")))
}

func TestSetCellValue(t *testing.T) {
	f := NewFile()
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), f.SetCellValue("Sheet1", "A", time.Now().UTC()))
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), f.SetCellValue("Sheet1", "A", time.Duration(1e13)))
	// Test set cell value with column and row style inherit
	style1, err := f.NewStyle(&Style{NumFmt: 2})
	assert.NoError(t, err)
	style2, err := f.NewStyle(&Style{NumFmt: 9})
	assert.NoError(t, err)
	assert.NoError(t, f.SetColStyle("Sheet1", "B", style1))
	assert.NoError(t, f.SetRowStyle("Sheet1", 1, 1, style2))
	assert.NoError(t, f.SetCellValue("Sheet1", "B1", 0.5))
	assert.NoError(t, f.SetCellValue("Sheet1", "B2", 0.5))
	B1, err := f.GetCellValue("Sheet1", "B1")
	assert.NoError(t, err)
	assert.Equal(t, "50%", B1)
	B2, err := f.GetCellValue("Sheet1", "B2")
	assert.NoError(t, err)
	assert.Equal(t, "0.50", B2)

	// Test set cell value with invalid sheet name
	assert.Equal(t, ErrSheetNameInvalid, f.SetCellValue("Sheet:1", "A1", "A1"))
	// Test set cell value with unsupported charset shared strings table
	f.SharedStrings = nil
	f.Pkg.Store(defaultXMLPathSharedStrings, MacintoshCyrillicCharset)
	assert.EqualError(t, f.SetCellValue("Sheet1", "A1", "A1"), "XML syntax error on line 1: invalid UTF-8")
	// Test set cell value with unsupported charset workbook
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	assert.EqualError(t, f.SetCellValue("Sheet1", "A1", time.Now().UTC()), "XML syntax error on line 1: invalid UTF-8")
	// Test set cell value with the shared string table's count not equal with unique count
	f = NewFile()
	f.SharedStrings = nil
	f.Pkg.Store(defaultXMLPathSharedStrings, []byte(fmt.Sprintf(`<sst xmlns="%s" count="2" uniqueCount="1"><si><t>a</t></si><si><t>a</t></si></sst>`, NameSpaceSpreadSheet.Value)))
	f.Sheet.Store("xl/worksheets/sheet1.xml", &xlsxWorksheet{
		SheetData: xlsxSheetData{Row: []xlsxRow{
			{R: 1, C: []xlsxC{{R: "A1", T: "str", V: "1"}}},
		}},
	})
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "b"))
	val, err := f.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "b", val)
	assert.NoError(t, f.SetCellValue("Sheet1", "B1", "b"))
	val, err = f.GetCellValue("Sheet1", "B1")
	assert.NoError(t, err)
	assert.Equal(t, "b", val)

	f = NewFile()
	// Test set cell value with an IEEE 754 "not-a-number" value or infinity
	for num, expected := range map[float64]string{
		math.NaN():   "NaN",
		math.Inf(0):  "+Inf",
		math.Inf(-1): "-Inf",
	} {
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", num))
		val, err := f.GetCellValue("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Equal(t, expected, val)
	}
	// Test set cell value with time duration
	for val, expected := range map[time.Duration]string{
		time.Hour*21 + time.Minute*51 + time.Second*44: "21:51:44",
		time.Hour*21 + time.Minute*50:                  "21:50",
		time.Hour*24 + time.Minute*51 + time.Second*44: "24:51:44",
		time.Hour*24 + time.Minute*50:                  "24:50:00",
	} {
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", val))
		val, err := f.GetCellValue("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Equal(t, expected, val)
	}
	// Test set cell value with time
	for val, expected := range map[time.Time]string{
		time.Date(2024, time.October, 1, 0, 0, 0, 0, time.UTC):   "Oct-24",
		time.Date(2024, time.October, 10, 0, 0, 0, 0, time.UTC):  "10-10-24",
		time.Date(2024, time.October, 10, 12, 0, 0, 0, time.UTC): "10/10/24 12:00",
	} {
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", val))
		val, err := f.GetCellValue("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Equal(t, expected, val)
	}
}

func TestSetCellValues(t *testing.T) {
	f := NewFile()
	err := f.SetCellValue("Sheet1", "A1", time.Date(2010, time.December, 31, 0, 0, 0, 0, time.UTC))
	assert.NoError(t, err)

	v, err := f.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, v, "12-31-10")

	// Test date value lower than min date supported by Excel
	err = f.SetCellValue("Sheet1", "A1", time.Date(1600, time.December, 31, 0, 0, 0, 0, time.UTC))
	assert.NoError(t, err)

	v, err = f.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, v, "1600-12-31T00:00:00Z")
}

func TestSetCellBool(t *testing.T) {
	f := NewFile()
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), f.SetCellBool("Sheet1", "A", true))
	// Test set cell boolean data type value with invalid sheet name
	assert.Equal(t, ErrSheetNameInvalid, f.SetCellBool("Sheet:1", "A1", true))
}

func TestSetCellTime(t *testing.T) {
	date, err := time.Parse(time.RFC3339Nano, "2009-11-10T23:00:00Z")
	assert.NoError(t, err)
	for location, expected := range map[string]string{
		"America/New_York": "40127.75",
		"Asia/Shanghai":    "40128.291666666664",
		"Europe/London":    "40127.958333333336",
		"UTC":              "40127.958333333336",
	} {
		timezone, err := time.LoadLocation(location)
		assert.NoError(t, err)
		c := &xlsxC{}
		isNum, err := c.setCellTime(date.In(timezone), false)
		assert.NoError(t, err)
		assert.Equal(t, true, isNum)
		assert.Equal(t, expected, c.V)
	}
}

func TestGetCellValue(t *testing.T) {
	// Test get cell value without r attribute of the row
	f := NewFile()
	sheetData := `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>%s</sheetData></worksheet>`

	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(fmt.Sprintf(sheetData, `<row r="3"><c t="inlineStr"><is><t>A3</t></is></c></row><row><c t="inlineStr"><is><t>A4</t></is></c><c t="inlineStr"><is><t>B4</t></is></c></row><row r="7"><c t="inlineStr"><is><t>A7</t></is></c><c t="inlineStr"><is><t>B7</t></is></c></row><row><c t="inlineStr"><is><t>A8</t></is></c><c t="inlineStr"><is><t>B8</t></is></c></row>`)))
	f.checked = sync.Map{}
	cells := []string{"A3", "A4", "B4", "A7", "B7"}
	rows, err := f.GetRows("Sheet1")
	assert.Equal(t, [][]string{nil, nil, {"A3"}, {"A4", "B4"}, nil, nil, {"A7", "B7"}, {"A8", "B8"}}, rows)
	assert.NoError(t, err)
	for _, cell := range cells {
		value, err := f.GetCellValue("Sheet1", cell)
		assert.Equal(t, cell, value)
		assert.NoError(t, err)
	}
	cols, err := f.GetCols("Sheet1")
	assert.Equal(t, [][]string{{"", "", "A3", "A4", "", "", "A7", "A8"}, {"", "", "", "B4", "", "", "B7", "B8"}}, cols)
	assert.NoError(t, err)

	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(fmt.Sprintf(sheetData, `<row r="2"><c r="A2" t="inlineStr"><is><t>A2</t></is></c></row><row r="2"><c r="B2" t="inlineStr"><is><t>B2</t></is></c></row>`)))
	f.checked = sync.Map{}
	cell, err := f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "A2", cell)
	assert.NoError(t, err)

	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(fmt.Sprintf(sheetData, `<row r="2"><c r="A2" t="inlineStr"><is><t>A2</t></is></c></row><row r="2"><c r="B2" t="inlineStr"><is><t>B2</t></is></c></row>`)))
	f.checked = sync.Map{}
	rows, err = f.GetRows("Sheet1")
	assert.Equal(t, [][]string{nil, {"A2", "B2"}}, rows)
	assert.NoError(t, err)

	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(fmt.Sprintf(sheetData, `<row r="1"><c r="A1" t="inlineStr"><is><t>A1</t></is></c></row><row r="1"><c r="B1" t="inlineStr"><is><t>B1</t></is></c></row>`)))
	f.checked = sync.Map{}
	rows, err = f.GetRows("Sheet1")
	assert.Equal(t, [][]string{{"A1", "B1"}}, rows)
	assert.NoError(t, err)

	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(fmt.Sprintf(sheetData, `<row><c t="inlineStr"><is><t>A3</t></is></c></row><row><c t="inlineStr"><is><t>A4</t></is></c><c t="inlineStr"><is><t>B4</t></is></c></row><row r="7"><c t="inlineStr"><is><t>A7</t></is></c><c t="inlineStr"><is><t>B7</t></is></c></row><row><c t="inlineStr"><is><t>A8</t></is></c><c t="inlineStr"><is><t>B8</t></is></c></row>`)))
	f.checked = sync.Map{}
	rows, err = f.GetRows("Sheet1")
	assert.Equal(t, [][]string{{"A3"}, {"A4", "B4"}, nil, nil, nil, nil, {"A7", "B7"}, {"A8", "B8"}}, rows)
	assert.NoError(t, err)

	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(fmt.Sprintf(sheetData, `<row r="0"><c r="H6" t="inlineStr"><is><t>H6</t></is></c><c r="A1" t="inlineStr"><is><t>r0A6</t></is></c><c r="F4" t="inlineStr"><is><t>F4</t></is></c></row><row><c r="A1" t="inlineStr"><is><t>A6</t></is></c><c r="B1" t="inlineStr"><is><t>B6</t></is></c><c r="C1" t="inlineStr"><is><t>C6</t></is></c></row><row r="3"><c r="A3"><v>100</v></c><c r="B3" t="inlineStr"><is><t>B3</t></is></c></row>`)))
	f.checked = sync.Map{}
	cell, err = f.GetCellValue("Sheet1", "H6")
	assert.Equal(t, "H6", cell)
	assert.NoError(t, err)
	rows, err = f.GetRows("Sheet1")
	assert.Equal(t, [][]string{
		{"A6", "B6", "C6"},
		nil,
		{"100", "B3"},
		{"", "", "", "", "", "F4"},
		nil,
		{"", "", "", "", "", "", "", "H6"},
	}, rows)
	assert.NoError(t, err)

	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(fmt.Sprintf(sheetData, `<row><c r="A1" t="inlineStr"><is><t>A1</t></is></c></row><row></row><row><c r="A3" t="inlineStr"><is><t>A3</t></is></c></row>`)))
	f.checked = sync.Map{}
	rows, err = f.GetRows("Sheet1")
	assert.Equal(t, [][]string{{"A1"}, nil, {"A3"}}, rows)
	assert.NoError(t, err)
	cell, err = f.GetCellValue("Sheet1", "A3")
	assert.Equal(t, "A3", cell)
	assert.NoError(t, err)

	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(fmt.Sprintf(sheetData, `
	<row r="1"><c r="A1"><v>2422.3000000000002</v></c></row>
	<row r="2"><c r="A2"><v>2422.3000000000002</v></c></row>
	<row r="3"><c r="A3"><v>12.4</v></c></row>
	<row r="4"><c r="A4"><v>964</v></c></row>
	<row r="5"><c r="A5"><v>1101.5999999999999</v></c></row>
	<row r="6"><c r="A6"><v>275.39999999999998</v></c></row>
	<row r="7"><c r="A7"><v>68.900000000000006</v></c></row>
	<row r="8"><c r="A8"><v>44385.208333333336</v></c></row>
	<row r="9"><c r="A9"><v>5.0999999999999996</v></c></row>
	<row r="10"><c r="A10"><v>5.1100000000000003</v></c></row>
	<row r="11"><c r="A11"><v>5.0999999999999996</v></c></row>
	<row r="12"><c r="A12"><v>5.1109999999999998</v></c></row>
	<row r="13"><c r="A13"><v>5.1111000000000004</v></c></row>
	<row r="14"><c r="A14"><v>2422.012345678</v></c></row>
	<row r="15"><c r="A15"><v>2422.0123456789</v></c></row>
	<row r="16"><c r="A16"><v>12.012345678901</v></c></row>
	<row r="17"><c r="A17"><v>964</v></c></row>
	<row r="18"><c r="A18"><v>1101.5999999999999</v></c></row>
	<row r="19"><c r="A19"><v>275.39999999999998</v></c></row>
	<row r="20"><c r="A20"><v>68.900000000000006</v></c></row>
	<row r="21"><c r="A21"><v>8.8880000000000001E-2</v></c></row>
	<row r="22"><c r="A22"><v>4.0000000000000003e-5</v></c></row>
	<row r="23"><c r="A23"><v>2422.3000000000002</v></c></row>
	<row r="24"><c r="A24"><v>1101.5999999999999</v></c></row>
	<row r="25"><c r="A25"><v>275.39999999999998</v></c></row>
	<row r="26"><c r="A26"><v>68.900000000000006</v></c></row>
	<row r="27"><c r="A27"><v>1.1000000000000001</v></c></row>
	<row r="28"><c r="A28" t="inlineStr"><is><t>1234567890123_4</t></is></c></row>
	<row r="29"><c r="A29" t="inlineStr"><is><t>123456789_0123_4</t></is></c></row>
	<row r="30"><c r="A30"><v>+0.0000000000000000002399999999999992E-4</v></c></row>
	<row r="31"><c r="A31"><v>7.2399999999999992E-2</v></c></row>
	<row r="32"><c r="A32" t="d"><v>20200208T080910.123</v></c></row>
	<row r="33"><c r="A33" t="d"><v>20200208T080910,123</v></c></row>
	<row r="34"><c r="A34" t="d"><v>20221022T150529Z</v></c></row>
	<row r="35"><c r="A35" t="d"><v>2022-10-22T15:05:29Z</v></c></row>
	<row r="36"><c r="A36" t="d"><v>2020-07-10 15:00:00.000</v></c></row>`)))
	f.checked = sync.Map{}
	rows, err = f.GetCols("Sheet1")
	assert.Equal(t, []string{
		"2422.3",
		"2422.3",
		"12.4",
		"964",
		"1101.6",
		"275.4",
		"68.9",
		"44385.2083333333",
		"5.1",
		"5.11",
		"5.1",
		"5.111",
		"5.1111",
		"2422.012345678",
		"2422.0123456789",
		"12.012345678901",
		"964",
		"1101.6",
		"275.4",
		"68.9",
		"0.08888",
		"0.00004",
		"2422.3",
		"1101.6",
		"275.4",
		"68.9",
		"1.1",
		"1234567890123_4",
		"123456789_0123_4",
		"2.39999999999999E-23",
		"0.0724",
		"43869.3397004977",
		"43869.3397004977",
		"44856.6288078704",
		"44856.6288078704",
		"2020-07-10 15:00:00.000",
	}, rows[0])
	assert.NoError(t, err)

	// Test get cell value with unsupported charset shared strings table
	f.SharedStrings = nil
	f.Pkg.Store(defaultXMLPathSharedStrings, MacintoshCyrillicCharset)
	_, value := f.GetCellValue("Sheet1", "A1")
	assert.EqualError(t, value, "XML syntax error on line 1: invalid UTF-8")
	// Test get cell value with invalid sheet name
	_, err = f.GetCellValue("Sheet:1", "A1")
	assert.Equal(t, ErrSheetNameInvalid, err)
}

func TestGetCellType(t *testing.T) {
	f := NewFile()
	cellType, err := f.GetCellType("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, CellTypeUnset, cellType)
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "A1"))
	cellType, err = f.GetCellType("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, CellTypeSharedString, cellType)
	_, err = f.GetCellType("Sheet1", "A")
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), err)
	// Test get cell type with invalid sheet name
	_, err = f.GetCellType("Sheet:1", "A1")
	assert.Equal(t, ErrSheetNameInvalid, err)
}

func TestGetValueFrom(t *testing.T) {
	f := NewFile()
	c := xlsxC{T: "s"}
	sst, err := f.sharedStringsReader()
	assert.NoError(t, err)
	value, err := c.getValueFrom(f, sst, false)
	assert.NoError(t, err)
	assert.Equal(t, "", value)

	c = xlsxC{T: "s", V: " 1 "}
	value, err = c.getValueFrom(f, &xlsxSST{Count: 1, SI: []xlsxSI{{}, {T: &xlsxT{Val: "s"}}}}, false)
	assert.NoError(t, err)
	assert.Equal(t, "s", value)
}

func TestGetCellFormula(t *testing.T) {
	// Test get cell formula on not exist worksheet
	f := NewFile()
	_, err := f.GetCellFormula("SheetN", "A1")
	assert.EqualError(t, err, "sheet SheetN does not exist")

	// Test get cell formula with invalid sheet name
	_, err = f.GetCellFormula("Sheet:1", "A1")
	assert.Equal(t, ErrSheetNameInvalid, err)

	// Test get cell formula on no formula cell
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", true))
	_, err = f.GetCellFormula("Sheet1", "A1")
	assert.NoError(t, err)

	// Test get cell shared formula
	f = NewFile()
	sheetData := `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>2*A1</f></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" ref="B2:B7" si="0">%s</f></c></row><row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/></c></row><row r="4"><c r="A4"><v>4</v></c><c r="B4"><f t="shared" si="0"/></c></row><row r="5"><c r="A5"><v>5</v></c><c r="B5"><f t="shared" si="0"/></c></row><row r="6"><c r="A6"><v>6</v></c><c r="B6"><f t="shared" si="0"/></c></row><row r="7"><c r="A7"><v>7</v></c><c r="B7"><f t="shared" si="0"/></c></row></sheetData></worksheet>`

	for sharedFormula, expected := range map[string]string{
		`2*A2`:           `2*A3`,
		`2*A1A`:          `2*A2A`,
		`2*$A$2+LEN("")`: `2*$A$2+LEN("")`,
	} {
		f.Sheet.Delete("xl/worksheets/sheet1.xml")
		f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(fmt.Sprintf(sheetData, sharedFormula)))
		formula, err := f.GetCellFormula("Sheet1", "B3")
		assert.NoError(t, err)
		assert.Equal(t, expected, formula)
	}

	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2"><c r="B2"><f t="shared" si="0"></f></c></row></sheetData></worksheet>`))
	formula, err := f.GetCellFormula("Sheet1", "B2")
	assert.NoError(t, err)
	assert.Equal(t, "", formula)

	// Test get array formula with invalid cell range reference
	f = NewFile()
	assert.NoError(t, f.AddChartSheet("Chart1", &Chart{Type: Line}))
	_, err = f.NewSheet("Sheet2")
	assert.NoError(t, err)
	formulaType, ref := STCellFormulaTypeArray, "B1:B2"
	assert.NoError(t, f.SetCellFormula("Sheet2", "B1", "A1:B2", FormulaOpts{Ref: &ref, Type: &formulaType}))
	ws, ok := f.Sheet.Load("xl/worksheets/sheet3.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).SheetData.Row[0].C[1].F.Ref = ":"
	_, err = f.getCellFormula("Sheet2", "A1", true)
	assert.Equal(t, newCellNameToCoordinatesError("", newInvalidCellNameError("")), err)

	// Test set formula for the cells in array formula range with unsupported charset
	f = NewFile()
	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", MacintoshCyrillicCharset)
	assert.EqualError(t, f.setArrayFormulaCells(), "XML syntax error on line 1: invalid UTF-8")
}

func ExampleFile_SetCellFloat() {
	f := NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	x := 3.14159265
	if err := f.SetCellFloat("Sheet1", "A1", x, 2, 64); err != nil {
		fmt.Println(err)
	}
	val, err := f.GetCellValue("Sheet1", "A1")
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println(val)
	// Output: 3.14
}

func BenchmarkSetCellValue(b *testing.B) {
	values := []string{"First", "Second", "Third", "Fourth", "Fifth", "Sixth"}
	cols := []string{"A", "B", "C", "D", "E", "F"}
	f := NewFile()
	b.ResetTimer()
	for i := 1; i <= b.N; i++ {
		for j := 0; j < len(values); j++ {
			if err := f.SetCellValue("Sheet1", cols[j]+strconv.Itoa(i), values[j]); err != nil {
				b.Error(err)
			}
		}
	}
}

func TestOverflowNumericCell(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "OverflowNumericCell.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	val, err := f.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	// GOARCH=amd64 - all ok; GOARCH=386 - actual: "-2147483648"
	assert.Equal(t, "8595602512225", val, "A1 should be 8595602512225")
	assert.NoError(t, f.Close())
}

func TestSetCellFormula(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	assert.NoError(t, f.SetCellFormula("Sheet1", "B19", "SUM(Sheet2!D2,Sheet2!D11)"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "C19", "SUM(Sheet2!D2,Sheet2!D9)"))

	// Test set cell formula with invalid sheet name
	assert.Equal(t, ErrSheetNameInvalid, f.SetCellFormula("Sheet:1", "A1", "SUM(1,2)"))

	// Test set cell formula with illegal rows number
	assert.Equal(t, newCellNameToCoordinatesError("C", newInvalidCellNameError("C")), f.SetCellFormula("Sheet1", "C", "SUM(Sheet2!D2,Sheet2!D9)"))

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellFormula1.xlsx")))
	assert.NoError(t, f.Close())

	f, err = OpenFile(filepath.Join("test", "CalcChain.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	// Test remove cell formula
	assert.NoError(t, f.SetCellFormula("Sheet1", "A1", ""))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellFormula2.xlsx")))
	// Test remove all cell formula
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", ""))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellFormula3.xlsx")))
	assert.NoError(t, f.Close())

	// Test set shared formula for the cells
	f = NewFile()
	for r := 1; r <= 5; r++ {
		assert.NoError(t, f.SetSheetRow("Sheet1", fmt.Sprintf("A%d", r), &[]interface{}{r, r + 1}))
	}
	formulaType, ref := STCellFormulaTypeShared, "C1:C5"
	assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "=A1+B1", FormulaOpts{Ref: &ref, Type: &formulaType}))
	sharedFormulaSpreadsheet := filepath.Join("test", "TestSetCellFormula4.xlsx")
	assert.NoError(t, f.SaveAs(sharedFormulaSpreadsheet))

	f, err = OpenFile(sharedFormulaSpreadsheet)
	assert.NoError(t, err)
	ref = "D1:D5"
	assert.NoError(t, f.SetCellFormula("Sheet1", "D1", "=A1+C1", FormulaOpts{Ref: &ref, Type: &formulaType}))
	ref = ""
	assert.Equal(t, ErrParameterInvalid, f.SetCellFormula("Sheet1", "D1", "=A1+C1", FormulaOpts{Ref: &ref, Type: &formulaType}))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellFormula5.xlsx")))

	// Test set table formula for the cells
	f = NewFile()
	for idx, row := range [][]interface{}{{"A", "B", "C"}, {1, 2}} {
		assert.NoError(t, f.SetSheetRow("Sheet1", fmt.Sprintf("A%d", idx+1), &row))
	}
	assert.NoError(t, f.AddTable("Sheet1", &Table{Range: "A1:C2", Name: "Table1", StyleName: "TableStyleMedium2"}))
	formulaType = STCellFormulaTypeDataTable
	assert.NoError(t, f.SetCellFormula("Sheet1", "C2", "=SUM(Table1[[A]:[B]])", FormulaOpts{Type: &formulaType}))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellFormula6.xlsx")))

	// Test set array formula with invalid cell range reference
	formulaType, ref = STCellFormulaTypeArray, ":"
	assert.Equal(t, newCellNameToCoordinatesError("", newInvalidCellNameError("")), f.SetCellFormula("Sheet1", "B1", "A1:A2", FormulaOpts{Ref: &ref, Type: &formulaType}))

	// Test set array formula with invalid cell reference
	formulaType, ref = STCellFormulaTypeArray, "A1:A2"
	assert.Equal(t, ErrColumnNumber, f.SetCellFormula("Sheet1", "A1", "SUM(XFE1:XFE2)", FormulaOpts{Ref: &ref, Type: &formulaType}))
}

func TestGetCellRichText(t *testing.T) {
	f, theme := NewFile(), 1

	runsSource := []RichTextRun{
		{
			Text: "a\n",
		},
		{
			Text: "b",
			Font: &Font{
				Underline:  "single",
				Color:      "ff0000",
				ColorTheme: &theme,
				ColorTint:  0.5,
				Bold:       true,
				Italic:     true,
				Family:     "Times New Roman",
				Size:       100,
				Strike:     true,
			},
		},
	}
	assert.NoError(t, f.SetCellRichText("Sheet1", "A1", runsSource))
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", false))

	runs, err := f.GetCellRichText("Sheet1", "A2")
	assert.NoError(t, err)
	assert.Equal(t, []RichTextRun(nil), runs)

	runs, err = f.GetCellRichText("Sheet1", "A1")
	assert.NoError(t, err)

	assert.Equal(t, runsSource[0].Text, runs[0].Text)
	assert.Nil(t, runs[0].Font)
	assert.NotNil(t, runs[1].Font)

	runsSource[1].Font.Color = strings.ToUpper(runsSource[1].Font.Color)
	assert.True(t, reflect.DeepEqual(runsSource[1].Font, runs[1].Font), "should get the same font")

	// Test get cell rich text with inlineStr
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).SheetData.Row[0].C[0] = xlsxC{
		T: "inlineStr",
		IS: &xlsxSI{
			T: &xlsxT{Val: "A"},
			R: []xlsxR{{T: &xlsxT{Val: "1"}}},
		},
	}
	runs, err = f.GetCellRichText("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, []RichTextRun{{Text: "A"}, {Text: "1"}}, runs)

	// Test get cell rich text when string item index overflow
	ws, ok = f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).SheetData.Row[0].C[0] = xlsxC{V: "2", IS: &xlsxSI{}}
	runs, err = f.GetCellRichText("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, 0, len(runs))
	// Test get cell rich text when string item index is negative
	ws, ok = f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).SheetData.Row[0].C[0] = xlsxC{T: "s", V: "-1", IS: &xlsxSI{}}
	runs, err = f.GetCellRichText("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, 0, len(runs))
	// Test get cell rich text when string item index is invalid
	ws, ok = f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).SheetData.Row[0].C[0] = xlsxC{T: "s", V: "A", IS: &xlsxSI{}}
	runs, err = f.GetCellRichText("Sheet1", "A1")
	assert.EqualError(t, err, "strconv.Atoi: parsing \"A\": invalid syntax")
	assert.Equal(t, 0, len(runs))
	// Test get cell rich text on invalid string item index
	ws, ok = f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).SheetData.Row[0].C[0] = xlsxC{V: "x"}
	runs, err = f.GetCellRichText("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, 0, len(runs))
	// Test set cell rich text on not exists worksheet
	_, err = f.GetCellRichText("SheetN", "A1")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Test set cell rich text with illegal cell reference
	_, err = f.GetCellRichText("Sheet1", "A")
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), err)
	// Test set rich text color theme without tint
	assert.NoError(t, f.SetCellRichText("Sheet1", "A1", []RichTextRun{{Font: &Font{ColorTheme: &theme}}}))
	// Test set rich text color tint without theme
	assert.NoError(t, f.SetCellRichText("Sheet1", "A1", []RichTextRun{{Font: &Font{ColorTint: 0.5}}}))

	// Test set cell rich text with unsupported charset shared strings table
	f.SharedStrings = nil
	f.Pkg.Store(defaultXMLPathSharedStrings, MacintoshCyrillicCharset)
	assert.EqualError(t, f.SetCellRichText("Sheet1", "A1", runsSource), "XML syntax error on line 1: invalid UTF-8")
	// Test get cell rich text with unsupported charset shared strings table
	f.SharedStrings = nil
	f.Pkg.Store(defaultXMLPathSharedStrings, MacintoshCyrillicCharset)
	_, err = f.GetCellRichText("Sheet1", "A1")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	// Test get cell rich text with invalid sheet name
	_, err = f.GetCellRichText("Sheet:1", "A1")
	assert.Equal(t, ErrSheetNameInvalid, err)
}

func TestSetCellRichText(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetRowHeight("Sheet1", 1, 35))
	assert.NoError(t, f.SetColWidth("Sheet1", "A", "A", 44))
	richTextRun := []RichTextRun{
		{
			Text: "bold",
			Font: &Font{
				Bold:         true,
				Color:        "2354E8",
				ColorIndexed: 0,
				Family:       "Times New Roman",
			},
		},
		{
			Text: " and ",
			Font: &Font{
				Family: "Times New Roman",
			},
		},
		{
			Text: "italic ",
			Font: &Font{
				Bold:   true,
				Color:  "E83723",
				Italic: true,
				Family: "Times New Roman",
			},
		},
		{
			Text: "text with color and font-family, ",
			Font: &Font{
				Bold:   true,
				Color:  "2354E8",
				Family: "Times New Roman",
			},
		},
		{
			Text: "\r\nlarge text with ",
			Font: &Font{
				Size:  14,
				Color: "AD23E8",
			},
		},
		{
			Text: "strike",
			Font: &Font{
				Color:  "E89923",
				Strike: true,
			},
		},
		{
			Text: " superscript",
			Font: &Font{
				Color:     "DBC21F",
				VertAlign: "superscript",
			},
		},
		{
			Text: " and ",
			Font: &Font{
				Size:      14,
				Color:     "AD23E8",
				VertAlign: "baseline",
			},
		},
		{
			Text: "underline",
			Font: &Font{
				Color:     "23E833",
				Underline: "single",
			},
		},
		{
			Text: " subscript.",
			Font: &Font{
				Color:     "017505",
				VertAlign: "subscript",
			},
		},
	}
	assert.NoError(t, f.SetCellRichText("Sheet1", "A1", richTextRun))
	assert.NoError(t, f.SetCellRichText("Sheet1", "A2", richTextRun))
	style, err := f.NewStyle(&Style{
		Alignment: &Alignment{
			WrapText: true,
		},
	})
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "A1", "A1", style))

	runs, err := f.GetCellRichText("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, richTextRun, runs)

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellRichText.xlsx")))
	// Test set cell rich text on not exists worksheet
	assert.EqualError(t, f.SetCellRichText("SheetN", "A1", richTextRun), "sheet SheetN does not exist")
	// Test set cell rich text with invalid sheet name
	assert.EqualError(t, f.SetCellRichText("Sheet:1", "A1", richTextRun), ErrSheetNameInvalid.Error())
	// Test set cell rich text with illegal cell reference
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), f.SetCellRichText("Sheet1", "A", richTextRun))
	richTextRun = []RichTextRun{{Text: strings.Repeat("s", TotalCellChars+1)}}
	// Test set cell rich text with characters over the maximum limit
	assert.EqualError(t, f.SetCellRichText("Sheet1", "A1", richTextRun), ErrCellCharsLength.Error())
}

func TestFormattedValue(t *testing.T) {
	f := NewFile()
	result, err := f.formattedValue(&xlsxC{S: 0, V: "43528"}, false, CellTypeNumber)
	assert.NoError(t, err)
	assert.Equal(t, "43528", result)

	// S is too large
	result, err = f.formattedValue(&xlsxC{S: 15, V: "43528"}, false, CellTypeNumber)
	assert.NoError(t, err)
	assert.Equal(t, "43528", result)

	// S is too small
	result, err = f.formattedValue(&xlsxC{S: -15, V: "43528"}, false, CellTypeNumber)
	assert.NoError(t, err)
	assert.Equal(t, "43528", result)

	result, err = f.formattedValue(&xlsxC{S: 1, V: "43528"}, false, CellTypeNumber)
	assert.NoError(t, err)
	assert.Equal(t, "43528", result)
	customNumFmt := "[$-409]MM/DD/YYYY"
	_, err = f.NewStyle(&Style{
		CustomNumFmt: &customNumFmt,
	})
	assert.NoError(t, err)
	result, err = f.formattedValue(&xlsxC{S: 1, V: "43528"}, false, CellTypeNumber)
	assert.NoError(t, err)
	assert.Equal(t, "03/04/2019", result)

	// Test format value with no built-in number format ID
	numFmtID := 5
	f.Styles.CellXfs.Xf = append(f.Styles.CellXfs.Xf, xlsxXf{
		NumFmtID: &numFmtID,
	})
	result, err = f.formattedValue(&xlsxC{S: 2, V: "43528"}, false, CellTypeNumber)
	assert.NoError(t, err)
	assert.Equal(t, "43528", result)

	// Test format value with invalid number format ID
	f.Styles.CellXfs.Xf = append(f.Styles.CellXfs.Xf, xlsxXf{
		NumFmtID: nil,
	})
	result, err = f.formattedValue(&xlsxC{S: 3, V: "43528"}, false, CellTypeNumber)
	assert.NoError(t, err)
	assert.Equal(t, "43528", result)

	// Test format value with empty number format
	f.Styles.NumFmts = nil
	f.Styles.CellXfs.Xf = append(f.Styles.CellXfs.Xf, xlsxXf{
		NumFmtID: &numFmtID,
	})
	result, err = f.formattedValue(&xlsxC{S: 1, V: "43528"}, false, CellTypeNumber)
	assert.NoError(t, err)
	assert.Equal(t, "43528", result)

	// Test format numeric value with shared string data type
	f.Styles.NumFmts, numFmtID = nil, 11
	f.Styles.CellXfs.Xf = append(f.Styles.CellXfs.Xf, xlsxXf{
		NumFmtID: &numFmtID,
	})
	result, err = f.formattedValue(&xlsxC{S: 5, V: "43528"}, false, CellTypeSharedString)
	assert.NoError(t, err)
	assert.Equal(t, "43528", result)

	// Test format decimal value with build-in number format ID
	styleID, err := f.NewStyle(&Style{
		NumFmt: 1,
	})
	assert.NoError(t, err)
	result, err = f.formattedValue(&xlsxC{S: styleID, V: "310.56"}, false, CellTypeNumber)
	assert.NoError(t, err)
	assert.Equal(t, "311", result)

	assert.Equal(t, "0_0", format("0_0", "", false, CellTypeNumber, nil))

	// Test format value with unsupported charset workbook
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	_, err = f.formattedValue(&xlsxC{S: 1, V: "43528"}, false, CellTypeNumber)
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")

	// Test format value with unsupported charset style sheet
	f.Styles = nil
	f.Pkg.Store(defaultXMLPathStyles, MacintoshCyrillicCharset)
	_, err = f.formattedValue(&xlsxC{S: 1, V: "43528"}, false, CellTypeNumber)
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")

	assert.Equal(t, "text", format("text", "0", false, CellTypeNumber, nil))
}

func TestFormattedValueNilXfs(t *testing.T) {
	// Set the CellXfs to nil and verify that the formattedValue function does not crash
	f := NewFile()
	f.Styles.CellXfs = nil
	result, err := f.formattedValue(&xlsxC{S: 3, V: "43528"}, false, CellTypeNumber)
	assert.NoError(t, err)
	assert.Equal(t, "43528", result)
}

func TestFormattedValueNilNumFmts(t *testing.T) {
	// Set the NumFmts value to nil and verify that the formattedValue function does not crash
	f := NewFile()
	f.Styles.NumFmts = nil
	result, err := f.formattedValue(&xlsxC{S: 3, V: "43528"}, false, CellTypeNumber)
	assert.NoError(t, err)
	assert.Equal(t, "43528", result)
}

func TestFormattedValueNilWorkbook(t *testing.T) {
	// Set the Workbook value to nil and verify that the formattedValue function does not crash
	f := NewFile()
	f.WorkBook = nil
	result, err := f.formattedValue(&xlsxC{S: 3, V: "43528"}, false, CellTypeNumber)
	assert.NoError(t, err)
	assert.Equal(t, "43528", result)
}

func TestFormattedValueNilWorkbookPr(t *testing.T) {
	// Set the WorkBook.WorkbookPr value to nil and verify that the formattedValue function does not
	// crash.
	f := NewFile()
	f.WorkBook.WorkbookPr = nil
	result, err := f.formattedValue(&xlsxC{S: 3, V: "43528"}, false, CellTypeNumber)
	assert.NoError(t, err)
	assert.Equal(t, "43528", result)
}

func TestGetCustomNumFmtCode(t *testing.T) {
	expected := "[$-ja-JP-x-gannen,80]ggge\"年\"m\"月\"d\"日\";@"
	styleSheet := &xlsxStyleSheet{NumFmts: &xlsxNumFmts{NumFmt: []*xlsxNumFmt{
		{NumFmtID: 164, FormatCode16: expected},
	}}}
	numFmtCode, ok := styleSheet.getCustomNumFmtCode(164)
	assert.Equal(t, expected, numFmtCode)
	assert.True(t, ok)
}

func TestSharedStringsError(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"), Options{UnzipXMLSizeLimit: 128})
	assert.NoError(t, err)
	tempFile, ok := f.tempFiles.Load(defaultXMLPathSharedStrings)
	assert.True(t, ok)
	f.tempFiles.Store(defaultXMLPathSharedStrings, "")
	assert.Equal(t, "1", f.getFromStringItem(1))
	// Cleanup undelete temporary files
	assert.NoError(t, os.Remove(tempFile.(string)))
	// Test reload the file error on set cell value and rich text. The error message was different between macOS and Windows
	err = f.SetCellValue("Sheet1", "A19", "A19")
	assert.Error(t, err)

	f.tempFiles.Store(defaultXMLPathSharedStrings, "")
	err = f.SetCellRichText("Sheet1", "A19", []RichTextRun{})
	assert.Error(t, err)
	assert.NoError(t, f.Close())

	f, err = OpenFile(filepath.Join("test", "Book1.xlsx"), Options{UnzipXMLSizeLimit: 128})
	assert.NoError(t, err)
	rows, err := f.Rows("Sheet1")
	assert.NoError(t, err)
	const maxUint16 = 1<<16 - 1
	currentRow := 0
	for rows.Next() {
		currentRow++
		if currentRow == 19 {
			_, err := rows.Columns()
			assert.NoError(t, err)
			// Test get cell value from string item with invalid offset
			f.sharedStringItem[1] = []uint{maxUint16 - 1, maxUint16}
			assert.Equal(t, "1", f.getFromStringItem(1))
			break
		}
	}
	assert.NoError(t, rows.Close())
	// Test shared string item temporary files has been closed before close the workbook
	assert.NoError(t, f.sharedStringTemp.Close())
	assert.Error(t, f.Close())
	// Cleanup undelete temporary files
	f.tempFiles.Range(func(k, v interface{}) bool {
		return assert.NoError(t, os.Remove(v.(string)))
	})

	f, err = OpenFile(filepath.Join("test", "Book1.xlsx"), Options{UnzipXMLSizeLimit: 128})
	assert.NoError(t, err)
	rows, err = f.Rows("Sheet1")
	assert.NoError(t, err)
	currentRow = 0
	for rows.Next() {
		currentRow++
		if currentRow == 19 {
			_, err := rows.Columns()
			assert.NoError(t, err)
			break
		}
	}
	assert.NoError(t, rows.Close())
	assert.NoError(t, f.sharedStringTemp.Close())
	// Test shared string item temporary files has been closed before set the cell value
	assert.Error(t, f.SetCellValue("Sheet1", "A1", "A1"))
	assert.Error(t, f.Close())
	// Cleanup undelete temporary files
	f.tempFiles.Range(func(k, v interface{}) bool {
		return assert.NoError(t, os.Remove(v.(string)))
	})
}

func TestSetCellIntFunc(t *testing.T) {
	cases := []struct {
		val    interface{}
		target string
	}{
		{val: 128, target: "128"},
		{val: int8(-128), target: "-128"},
		{val: int16(-32768), target: "-32768"},
		{val: int32(-2147483648), target: "-2147483648"},
		{val: int64(-9223372036854775808), target: "-9223372036854775808"},
		{val: uint(128), target: "128"},
		{val: uint8(255), target: "255"},
		{val: uint16(65535), target: "65535"},
		{val: uint32(4294967295), target: "4294967295"},
		{val: uint64(18446744073709551615), target: "18446744073709551615"},
	}
	for _, c := range cases {
		cell := &xlsxC{}
		setCellIntFunc(cell, c.val)
		assert.Equal(t, c.target, cell.V)
	}
}

func TestSIString(t *testing.T) {
	assert.Empty(t, xlsxSI{}.String())
}
