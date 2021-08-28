package excelize

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

var validColumns = []struct {
	Name string
	Num  int
}{
	{Name: "A", Num: 1},
	{Name: "Z", Num: 26},
	{Name: "AA", Num: 26 + 1},
	{Name: "AK", Num: 26 + 11},
	{Name: "ak", Num: 26 + 11},
	{Name: "Ak", Num: 26 + 11},
	{Name: "aK", Num: 26 + 11},
	{Name: "AZ", Num: 26 + 26},
	{Name: "ZZ", Num: 26 + 26*26},
	{Name: "AAA", Num: 26 + 26*26 + 1},
}

var invalidColumns = []struct {
	Name string
	Num  int
}{
	{Name: "", Num: -1},
	{Name: " ", Num: -1},
	{Name: "_", Num: -1},
	{Name: "__", Num: -1},
	{Name: "-1", Num: -1},
	{Name: "0", Num: -1},
	{Name: " A", Num: -1},
	{Name: "A ", Num: -1},
	{Name: "A1", Num: -1},
	{Name: "1A", Num: -1},
	{Name: " a", Num: -1},
	{Name: "a ", Num: -1},
	{Name: "a1", Num: -1},
	{Name: "1a", Num: -1},
	{Name: " _", Num: -1},
	{Name: "_ ", Num: -1},
	{Name: "_1", Num: -1},
	{Name: "1_", Num: -1},
}

var invalidCells = []string{"", "A", "AA", " A", "A ", "1A", "A1A", "A1 ", " A1", "1A1", "a-1", "A-1"}

var invalidIndexes = []int{-100, -2, -1, 0}

func TestColumnNameToNumber_OK(t *testing.T) {
	const msg = "Column %q"
	for _, col := range validColumns {
		out, err := ColumnNameToNumber(col.Name)
		if assert.NoErrorf(t, err, msg, col.Name) {
			assert.Equalf(t, col.Num, out, msg, col.Name)
		}
	}
}

func TestColumnNameToNumber_Error(t *testing.T) {
	const msg = "Column %q"
	for _, col := range invalidColumns {
		out, err := ColumnNameToNumber(col.Name)
		if assert.Errorf(t, err, msg, col.Name) {
			assert.Equalf(t, col.Num, out, msg, col.Name)
		}
	}
	_, err := ColumnNameToNumber("XFE")
	assert.EqualError(t, err, ErrColumnNumber.Error())
}

func TestColumnNumberToName_OK(t *testing.T) {
	const msg = "Column %q"
	for _, col := range validColumns {
		out, err := ColumnNumberToName(col.Num)
		if assert.NoErrorf(t, err, msg, col.Name) {
			assert.Equalf(t, strings.ToUpper(col.Name), out, msg, col.Name)
		}
	}
}

func TestColumnNumberToName_Error(t *testing.T) {
	out, err := ColumnNumberToName(-1)
	if assert.Error(t, err) {
		assert.Equal(t, "", out)
	}

	out, err = ColumnNumberToName(0)
	if assert.Error(t, err) {
		assert.Equal(t, "", out)
	}

	_, err = ColumnNumberToName(TotalColumns + 1)
	assert.EqualError(t, err, ErrColumnNumber.Error())
}

func TestSplitCellName_OK(t *testing.T) {
	const msg = "Cell \"%s%d\""
	for i, col := range validColumns {
		row := i + 1
		c, r, err := SplitCellName(col.Name + strconv.Itoa(row))
		if assert.NoErrorf(t, err, msg, col.Name, row) {
			assert.Equalf(t, col.Name, c, msg, col.Name, row)
			assert.Equalf(t, row, r, msg, col.Name, row)
		}
	}
}

func TestSplitCellName_Error(t *testing.T) {
	const msg = "Cell %q"
	for _, cell := range invalidCells {
		c, r, err := SplitCellName(cell)
		if assert.Errorf(t, err, msg, cell) {
			assert.Equalf(t, "", c, msg, cell)
			assert.Equalf(t, -1, r, msg, cell)
		}
	}
}

func TestJoinCellName_OK(t *testing.T) {
	const msg = "Cell \"%s%d\""

	for i, col := range validColumns {
		row := i + 1
		cell, err := JoinCellName(col.Name, row)
		if assert.NoErrorf(t, err, msg, col.Name, row) {
			assert.Equalf(t, strings.ToUpper(fmt.Sprintf("%s%d", col.Name, row)), cell, msg, row)
		}
	}
}

func TestJoinCellName_Error(t *testing.T) {
	const msg = "Cell \"%s%d\""

	test := func(col string, row int) {
		cell, err := JoinCellName(col, row)
		if assert.Errorf(t, err, msg, col, row) {
			assert.Equalf(t, "", cell, msg, col, row)
		}
	}

	for _, col := range invalidColumns {
		test(col.Name, 1)
		for _, row := range invalidIndexes {
			test("A", row)
			test(col.Name, row)
		}
	}

}

func TestCellNameToCoordinates_OK(t *testing.T) {
	const msg = "Cell \"%s%d\""
	for i, col := range validColumns {
		row := i + 1
		c, r, err := CellNameToCoordinates(col.Name + strconv.Itoa(row))
		if assert.NoErrorf(t, err, msg, col.Name, row) {
			assert.Equalf(t, col.Num, c, msg, col.Name, row)
			assert.Equalf(t, i+1, r, msg, col.Name, row)
		}
	}
}

func TestCellNameToCoordinates_Error(t *testing.T) {
	const msg = "Cell %q"
	for _, cell := range invalidCells {
		c, r, err := CellNameToCoordinates(cell)
		if assert.Errorf(t, err, msg, cell) {
			assert.Equalf(t, -1, c, msg, cell)
			assert.Equalf(t, -1, r, msg, cell)
		}
	}
	_, _, err := CellNameToCoordinates("A1048577")
	assert.EqualError(t, err, "row number exceeds maximum limit")
}

func TestCoordinatesToCellName_OK(t *testing.T) {
	const msg = "Coordinates [%d, %d]"
	for i, col := range validColumns {
		row := i + 1
		cell, err := CoordinatesToCellName(col.Num, row)
		if assert.NoErrorf(t, err, msg, col.Num, row) {
			assert.Equalf(t, strings.ToUpper(col.Name+strconv.Itoa(row)), cell, msg, col.Num, row)
		}
	}
}

func TestCoordinatesToCellName_Error(t *testing.T) {
	const msg = "Coordinates [%d, %d]"

	test := func(col, row int) {
		cell, err := CoordinatesToCellName(col, row)
		if assert.Errorf(t, err, msg, col, row) {
			assert.Equalf(t, "", cell, msg, col, row)
		}
	}

	for _, col := range invalidIndexes {
		test(col, 1)
		for _, row := range invalidIndexes {
			test(1, row)
			test(col, row)
		}
	}
}

func TestCoordinatesToAreaRef(t *testing.T) {
	f := NewFile()
	_, err := f.coordinatesToAreaRef([]int{})
	assert.EqualError(t, err, ErrCoordinates.Error())
	_, err = f.coordinatesToAreaRef([]int{1, -1, 1, 1})
	assert.EqualError(t, err, "invalid cell coordinates [1, -1]")
	_, err = f.coordinatesToAreaRef([]int{1, 1, 1, -1})
	assert.EqualError(t, err, "invalid cell coordinates [1, -1]")
	ref, err := f.coordinatesToAreaRef([]int{1, 1, 1, 1})
	assert.NoError(t, err)
	assert.EqualValues(t, ref, "A1:A1")
}

func TestSortCoordinates(t *testing.T) {
	assert.EqualError(t, sortCoordinates(make([]int, 3)), ErrCoordinates.Error())
}

func TestInStrSlice(t *testing.T) {
	assert.EqualValues(t, -1, inStrSlice([]string{}, ""))
}

func TestBytesReplace(t *testing.T) {
	s := []byte{0x01}
	assert.EqualValues(t, s, bytesReplace(s, []byte{}, []byte{}, 0))
}

func TestGetRootElement(t *testing.T) {
	assert.Equal(t, 0, len(getRootElement(xml.NewDecoder(strings.NewReader("")))))
}

func TestSetIgnorableNameSpace(t *testing.T) {
	f := NewFile()
	f.xmlAttr["xml_path"] = []xml.Attr{{}}
	f.setIgnorableNameSpace("xml_path", 0, xml.Attr{Name: xml.Name{Local: "c14"}})
	assert.EqualValues(t, "c14", f.xmlAttr["xml_path"][0].Value)
}

func TestStack(t *testing.T) {
	s := NewStack()
	assert.Equal(t, s.Peek(), nil)
	assert.Equal(t, s.Pop(), nil)
}

func TestGenXMLNamespace(t *testing.T) {
	assert.Equal(t, genXMLNamespace([]xml.Attr{
		{Name: xml.Name{Space: NameSpaceXML, Local: "space"}, Value: "preserve"},
	}), `xml:space="preserve">`)
}

func TestBstrUnmarshal(t *testing.T) {
	bstrs := map[string]string{
		"*":                           "*",
		"*_x0000_":                    "*\x00",
		"*_x0008_":                    "*\b",
		"_x0008_*":                    "\b*",
		"*_x0008_*":                   "*\b*",
		"*_x4F60__x597D_":             "*你好",
		"*_xG000_":                    "*_xG000_",
		"*_xG05F_x0001_*":             "*_xG05F*",
		"*_x005F__x0008_*":            "*_\b*",
		"*_x005F_x0001_*":             "*_x0001_*",
		"*_x005f_x005F__x0008_*":      "*_x005F_\b*",
		"*_x005F_x005F_xG05F_x0006_*": "*_x005F_xG05F*",
		"*_x005F_x005F_x005F_x0006_*": "*_x005F_x0006_*",
		"_x005F__x0008_******":        "_\b******",
		"******_x005F__x0008_":        "******_\b",
		"******_x005F__x0008_******":  "******_\b******",
	}
	for bstr, expected := range bstrs {
		assert.Equal(t, expected, bstrUnmarshal(bstr))
	}
}

func TestBstrMarshal(t *testing.T) {
	bstrs := map[string]string{
		"*_xG05F_*":       "*_xG05F_*",
		"*_x0008_*":       "*_x005F_x0008_*",
		"*_x005F_*":       "*_x005F_x005F_*",
		"*_x005F_xG006_*": "*_x005F_x005F_xG006_*",
		"*_x005F_x0006_*": "*_x005F_x005F_x005F_x0006_*",
	}
	for bstr, expected := range bstrs {
		assert.Equal(t, expected, bstrMarshal(bstr))
	}
}
