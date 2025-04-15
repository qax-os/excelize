package excelize

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"
	"sync"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
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
	assert.ErrorIs(t, err, ErrColumnNumber)
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

	_, err = ColumnNumberToName(MaxColumns + 1)
	assert.ErrorIs(t, err, ErrColumnNumber)
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
	assert.EqualError(t, err, ErrMaxRows.Error())
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

func TestCoordinatesToRangeRef(t *testing.T) {
	_, err := coordinatesToRangeRef([]int{})
	assert.EqualError(t, err, ErrCoordinates.Error())
	_, err = coordinatesToRangeRef([]int{1, -1, 1, 1})
	assert.Equal(t, newCoordinatesToCellNameError(1, -1), err)
	_, err = coordinatesToRangeRef([]int{1, 1, 1, -1})
	assert.Equal(t, newCoordinatesToCellNameError(1, -1), err)
	ref, err := coordinatesToRangeRef([]int{1, 1, 1, 1})
	assert.NoError(t, err)
	assert.EqualValues(t, ref, "A1:A1")
}

func TestSortCoordinates(t *testing.T) {
	assert.EqualError(t, sortCoordinates(make([]int, 3)), ErrCoordinates.Error())
}

func TestInStrSlice(t *testing.T) {
	assert.EqualValues(t, -1, inStrSlice([]string{}, "", true))
}

func TestAttrValue(t *testing.T) {
	assert.Empty(t, (&attrValString{}).Value())
	assert.False(t, (&attrValBool{}).Value())
	assert.Zero(t, (&attrValFloat{}).Value())
}

func TestBoolValMarshal(t *testing.T) {
	bold := true
	node := &xlsxFont{B: &attrValBool{Val: &bold}}
	data, err := xml.Marshal(node)
	assert.NoError(t, err)
	assert.Equal(t, `<xlsxFont><b val="1"></b></xlsxFont>`, string(data))

	node = &xlsxFont{}
	err = xml.Unmarshal(data, node)
	assert.NoError(t, err)
	assert.NotEqual(t, nil, node)
	assert.NotEqual(t, nil, node.B)
	assert.NotEqual(t, nil, node.B.Val)
	assert.Equal(t, true, *node.B.Val)
}

func TestBoolValUnmarshalXML(t *testing.T) {
	node := xlsxFont{}
	assert.NoError(t, xml.Unmarshal([]byte("<xlsxFont><b val=\"\"></b></xlsxFont>"), &node))
	assert.Equal(t, true, *node.B.Val)
	for content, err := range map[string]string{
		"<xlsxFont><b val=\"0\"><i></i></b></xlsxFont>": "unexpected child of attrValBool",
		"<xlsxFont><b val=\"x\"></b></xlsxFont>":        "strconv.ParseBool: parsing \"x\": invalid syntax",
	} {
		assert.EqualError(t, xml.Unmarshal([]byte(content), &node), err)
	}
	attr := attrValBool{}
	assert.EqualError(t, attr.UnmarshalXML(xml.NewDecoder(strings.NewReader("")), xml.StartElement{}), io.EOF.Error())
}

func TestExtUnmarshalXML(t *testing.T) {
	f, extLst := NewFile(), decodeExtLst{}
	expected := fmt.Sprintf(`<extLst><ext uri="%s" xmlns:x14="%s"/></extLst>`,
		ExtURISlicerCachesX14, NameSpaceSpreadSheetX14.Value)
	assert.NoError(t, f.xmlNewDecoder(strings.NewReader(expected)).Decode(&extLst))
	assert.Len(t, extLst.Ext, 1)
	assert.Equal(t, extLst.Ext[0].URI, ExtURISlicerCachesX14)
}

func TestBytesReplace(t *testing.T) {
	s := []byte{0x01}
	assert.EqualValues(t, s, bytesReplace(s, []byte{}, []byte{}, 0))
}

func TestGetRootElement(t *testing.T) {
	assert.Len(t, getRootElement(xml.NewDecoder(strings.NewReader(""))), 0)
	// Test get workbook root element which all workbook XML namespace has prefix
	f := NewFile()
	d := f.xmlNewDecoder(bytes.NewReader([]byte(`<x:workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></x:workbook>`)))
	assert.Len(t, getRootElement(d), 3)
}

func TestSetIgnorableNameSpace(t *testing.T) {
	f := NewFile()
	f.xmlAttr.Store("xml_path", []xml.Attr{{}})
	f.setIgnorableNameSpace("xml_path", 0, xml.Attr{Name: xml.Name{Local: "c14"}})
	attrs, ok := f.xmlAttr.Load("xml_path")
	assert.EqualValues(t, "c14", attrs.([]xml.Attr)[0].Value)
	assert.True(t, ok)
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
		"*_xG05F_x0001_*":             "*_xG05F\x01*",
		"*_x005F__x0008_*":            "*_\b*",
		"*_x005F_x0001_*":             "*_x0001_*",
		"*_x005f_x005F__x0008_*":      "*_x005F_\b*",
		"*_x005F_x005F_xG05F_x0006_*": "*_x005F_xG05F\x06*",
		"*_x005F_x005F_x005F_x0006_*": "*_x005F_x0006_*",
		"_x005F__x0008_******":        "_\b******",
		"******_x005F__x0008_":        "******_\b",
		"******_x005F__x0008_******":  "******_\b******",
		"_x000x_x005F_x000x_":         "_x000x_x000x_",
	}
	for bstr, expected := range bstrs {
		assert.Equal(t, expected, bstrUnmarshal(bstr), bstr)
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

func TestReadBytes(t *testing.T) {
	f := &File{tempFiles: sync.Map{}}
	sheet := "xl/worksheets/sheet1.xml"
	f.tempFiles.Store(sheet, "/d/")
	assert.Equal(t, []byte{}, f.readBytes(sheet))
}

func TestUnzipToTemp(t *testing.T) {
	os.Setenv("TMPDIR", "test")
	defer os.Unsetenv("TMPDIR")
	assert.NoError(t, os.Chmod(os.TempDir(), 0o444))
	f := NewFile()
	data := []byte("PK\x03\x040000000PK\x01\x0200000" +
		"0000000000000000000\x00" +
		"\x00\x00\x00\x00\x00000000000000PK\x01" +
		"\x020000000000000000000" +
		"00000\v\x00\x00\x00\x00\x00000000000" +
		"00000000000000PK\x01\x0200" +
		"00000000000000000000" +
		"00\v\x00\x00\x00\x00\x00000000000000" +
		"00000000000PK\x01\x020000<" +
		"0\x00\x0000000000000000\v\x00\v" +
		"\x00\x00\x00\x00\x0000000000\x00\x00\x00\x00000" +
		"00000000PK\x01\x0200000000" +
		"0000000000000000\v\x00\x00\x00" +
		"\x00\x0000PK\x05\x06000000\x05\x00\xfd\x00\x00\x00" +
		"\v\x00\x00\x00\x00\x00")
	z, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	assert.NoError(t, err)

	_, err = f.unzipToTemp(z.File[0])
	require.Error(t, err)
	assert.NoError(t, os.Chmod(os.TempDir(), 0o755))

	_, err = f.unzipToTemp(z.File[0])
	assert.EqualError(t, err, "EOF")
}
