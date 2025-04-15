package excelize

import (
	"archive/zip"
	"bytes"
	"compress/gzip"
	"encoding/xml"
	"fmt"
	"image/color"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"io"
	"math"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"sync"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
)

func TestOpenFile(t *testing.T) {
	// Test update the spreadsheet file
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)

	// Test get all the rows in a not exists worksheet
	_, err = f.GetRows("Sheet4")
	assert.EqualError(t, err, "sheet Sheet4 does not exist")
	// Test get all the rows with invalid sheet name
	_, err = f.GetRows("Sheet:1")
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())
	// Test get all the rows in a worksheet
	rows, err := f.GetRows("Sheet2")
	expected := [][]string{
		{"Monitor", "", "Brand", "", "inlineStr"},
		{"> 23 Inch", "19", "HP", "200"},
		{"20-23 Inch", "24", "DELL", "450"},
		{"17-20 Inch", "56", "Lenove", "200"},
		{"< 17 Inch", "21", "SONY", "510"},
		{"", "", "Acer", "315"},
		{"", "", "IBM", "127"},
		{"", "", "ASUS", "89"},
		{"", "", "Apple", "348"},
		{"", "", "SAMSUNG", "53"},
		{"", "", "Other", "37", "", "", "", "", ""},
	}
	assert.NoError(t, err)
	assert.Equal(t, expected, rows)

	assert.NoError(t, f.UpdateLinkedValue())

	assert.NoError(t, f.SetCellDefault("Sheet2", "A1", strconv.FormatFloat(100.1588, 'f', -1, 32)))
	assert.NoError(t, f.SetCellDefault("Sheet2", "A1", strconv.FormatFloat(-100.1588, 'f', -1, 64)))
	// Test set cell value with invalid sheet name
	assert.EqualError(t, f.SetCellDefault("Sheet:1", "A1", ""), ErrSheetNameInvalid.Error())
	// Test set cell value with illegal row number
	assert.EqualError(t, f.SetCellDefault("Sheet2", "A", strconv.FormatFloat(-100.1588, 'f', -1, 64)),
		newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())

	assert.NoError(t, f.SetCellInt("Sheet2", "A1", 100))

	// Test set cell integer value with illegal row number
	assert.EqualError(t, f.SetCellInt("Sheet2", "A", 100), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	// Test set cell integer value with invalid sheet name
	assert.EqualError(t, f.SetCellInt("Sheet:1", "A1", 100), ErrSheetNameInvalid.Error())

	assert.NoError(t, f.SetCellStr("Sheet2", "C11", "Knowns"))
	// Test max characters in a cell
	assert.NoError(t, f.SetCellStr("Sheet2", "D11", strings.Repeat("c", TotalCellChars+2)))
	_, err = f.NewSheet(":\\/?*[]Maximum 31 characters allowed in sheet title.")
	assert.EqualError(t, err, ErrSheetNameLength.Error())
	// Test set worksheet name with illegal name
	assert.EqualError(t, f.SetSheetName("Maximum 31 characters allowed i", "[Rename]:\\/?* Maximum 31 characters allowed in sheet title."), ErrSheetNameLength.Error())
	assert.EqualError(t, f.SetCellInt("Sheet3", "A23", 10), "sheet Sheet3 does not exist")
	assert.EqualError(t, f.SetCellStr("Sheet3", "b230", "10"), "sheet Sheet3 does not exist")
	assert.EqualError(t, f.SetCellStr("Sheet10", "b230", "10"), "sheet Sheet10 does not exist")
	// Test set cell string data type value with invalid sheet name
	assert.EqualError(t, f.SetCellStr("Sheet:1", "A1", "1"), ErrSheetNameInvalid.Error())
	// Test set cell string value with illegal row number
	assert.EqualError(t, f.SetCellStr("Sheet1", "A", "10"), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())

	f.SetActiveSheet(2)
	// Test get cell formula with given rows number
	_, err = f.GetCellFormula("Sheet1", "B19")
	assert.NoError(t, err)
	// Test get cell formula with illegal worksheet name
	_, err = f.GetCellFormula("Sheet2", "B20")
	assert.NoError(t, err)
	_, err = f.GetCellFormula("Sheet1", "B20")
	assert.NoError(t, err)

	// Test get cell formula with illegal rows number
	_, err = f.GetCellFormula("Sheet1", "B")
	assert.EqualError(t, err, newCellNameToCoordinatesError("B", newInvalidCellNameError("B")).Error())
	// Test get shared cell formula
	_, err = f.GetCellFormula("Sheet2", "H11")
	assert.NoError(t, err)
	_, err = f.GetCellFormula("Sheet2", "I11")
	assert.NoError(t, err)
	getSharedFormula(&xlsxWorksheet{}, 0, "")

	// Test read cell value with given illegal rows number
	_, err = f.GetCellValue("Sheet2", "a-1")
	assert.EqualError(t, err, newCellNameToCoordinatesError("A-1", newInvalidCellNameError("A-1")).Error())
	_, err = f.GetCellValue("Sheet2", "A")
	assert.EqualError(t, err, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())

	// Test read cell value with given lowercase column number
	_, err = f.GetCellValue("Sheet2", "a5")
	assert.NoError(t, err)
	_, err = f.GetCellValue("Sheet2", "C11")
	assert.NoError(t, err)
	_, err = f.GetCellValue("Sheet2", "D11")
	assert.NoError(t, err)
	_, err = f.GetCellValue("Sheet2", "D12")
	assert.NoError(t, err)
	// Test SetCellValue function
	assert.NoError(t, f.SetCellValue("Sheet2", "F1", " Hello"))
	assert.NoError(t, f.SetCellValue("Sheet2", "G1", []byte("World")))
	assert.NoError(t, f.SetCellValue("Sheet2", "F2", 42))
	assert.NoError(t, f.SetCellValue("Sheet2", "F3", int8(1<<8/2-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F4", int16(1<<16/2-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F5", int32(1<<32/2-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F6", int64(1<<32/2-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F7", float32(42.65418)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F8", -42.65418))
	assert.NoError(t, f.SetCellValue("Sheet2", "F9", float32(42)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F10", float64(42)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F11", uint(1<<32-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F12", uint8(1<<8-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F13", uint16(1<<16-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F14", uint32(1<<32-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F15", uint64(1<<32-1)))
	assert.NoError(t, f.SetCellValue("Sheet2", "F16", true))
	assert.NoError(t, f.SetCellValue("Sheet2", "F17", complex64(5+10i)))

	// Test on not exists worksheet
	assert.EqualError(t, f.SetCellDefault("SheetN", "A1", ""), "sheet SheetN does not exist")
	assert.EqualError(t, f.SetCellFloat("SheetN", "A1", 42.65418, 2, 32), "sheet SheetN does not exist")
	assert.EqualError(t, f.SetCellBool("SheetN", "A1", true), "sheet SheetN does not exist")
	assert.EqualError(t, f.SetCellFormula("SheetN", "A1", ""), "sheet SheetN does not exist")
	assert.EqualError(t, f.SetCellHyperLink("SheetN", "A1", "Sheet1!A40", "Location"), "sheet SheetN does not exist")

	// Test boolean write
	boolTest := []struct {
		value    bool
		raw      bool
		expected string
	}{
		{false, true, "0"},
		{true, true, "1"},
		{false, false, "FALSE"},
		{true, false, "TRUE"},
	}
	for _, test := range boolTest {
		assert.NoError(t, f.SetCellValue("Sheet2", "F16", test.value))
		val, err := f.GetCellValue("Sheet2", "F16", Options{RawCellValue: test.raw})
		assert.NoError(t, err)
		assert.Equal(t, test.expected, val)
	}

	assert.NoError(t, f.SetCellValue("Sheet2", "G2", nil))

	assert.NoError(t, f.SetCellValue("Sheet2", "G4", time.Now()))

	assert.NoError(t, f.SetCellValue("Sheet2", "G4", time.Now().UTC()))
	assert.EqualError(t, f.SetCellValue("SheetN", "A1", time.Now()), "sheet SheetN does not exist")
	// 02:46:40
	assert.NoError(t, f.SetCellValue("Sheet2", "G5", time.Duration(1e13)))
	// Test completion column
	assert.NoError(t, f.SetCellValue("Sheet2", "M2", nil))
	// Test read cell value with given cell reference large than exists row
	_, err = f.GetCellValue("Sheet2", "E231")
	assert.NoError(t, err)
	// Test get active worksheet of spreadsheet and get worksheet name of
	// spreadsheet by given worksheet index
	f.GetSheetName(f.GetActiveSheetIndex())
	// Test get worksheet index of spreadsheet by given worksheet name
	_, err = f.GetSheetIndex("Sheet1")
	assert.NoError(t, err)
	// Test get worksheet name of spreadsheet by given invalid worksheet index
	f.GetSheetName(4)
	// Test get worksheet map of workbook
	f.GetSheetMap()
	for i := 1; i <= 300; i++ {
		assert.NoError(t, f.SetCellStr("Sheet2", "c"+strconv.Itoa(i), strconv.Itoa(i)))
	}
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestOpenFile.xlsx")))
	assert.EqualError(t, f.SaveAs(filepath.Join("test", strings.Repeat("c", 199), ".xlsx")), ErrMaxFilePathLength.Error())
	assert.NoError(t, f.Close())
}

func TestSaveFile(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)
	assert.EqualError(t, f.SaveAs(filepath.Join("test", "TestSaveFile.xlsb")), ErrWorkbookFileFormat.Error())
	for _, ext := range []string{".xlam", ".xlsm", ".xlsx", ".xltm", ".xltx"} {
		assert.NoError(t, f.SaveAs(filepath.Join("test", fmt.Sprintf("TestSaveFile%s", ext))))
	}
	assert.NoError(t, f.Close())

	f, err = OpenFile(filepath.Join("test", "TestSaveFile.xlsx"))
	assert.NoError(t, err)
	assert.NoError(t, f.Save())
	assert.NoError(t, f.Close())

	t.Run("for_save_multiple_times", func(t *testing.T) {
		{
			f, err := OpenFile(filepath.Join("test", "TestSaveFile.xlsx"))
			assert.NoError(t, err)
			assert.NoError(t, f.SetCellValue("Sheet1", "A20", 20))
			assert.NoError(t, f.Save())

			assert.NoError(t, f.SetCellValue("Sheet1", "A21", 21))
			assert.NoError(t, f.Save())
			assert.NoError(t, f.Close())
		}
		{
			f, err := OpenFile(filepath.Join("test", "TestSaveFile.xlsx"))
			assert.NoError(t, err)
			val, err := f.GetCellValue("Sheet1", "A20")
			assert.NoError(t, err)
			assert.Equal(t, "20", val)
			val, err = f.GetCellValue("Sheet1", "A21")
			assert.NoError(t, err)
			assert.Equal(t, "21", val)
			assert.NoError(t, f.Close())
		}
	})
}

func TestSaveAsWrongPath(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)
	// Test write file to not exist directory
	assert.Error(t, f.SaveAs(filepath.Join("x", "Book1.xlsx")))
	assert.NoError(t, f.Close())
}

func TestCharsetTranscoder(t *testing.T) {
	f := NewFile()
	f.CharsetTranscoder(*new(charsetTranscoderFn))
}

func TestOpenReader(t *testing.T) {
	_, err := OpenReader(strings.NewReader(""))
	assert.EqualError(t, err, zip.ErrFormat.Error())
	_, err = OpenReader(bytes.NewReader(oleIdentifier), Options{Password: "password", UnzipXMLSizeLimit: UnzipSizeLimit + 1})
	assert.EqualError(t, err, ErrWorkbookFileFormat.Error())

	// Prepare unusual workbook, made the specified internal XML parts missing
	// or contain unsupported charset
	preset := func(filePath string, notExist bool) *bytes.Buffer {
		source, err := zip.OpenReader(filepath.Join("test", "Book1.xlsx"))
		assert.NoError(t, err)
		buf := new(bytes.Buffer)
		zw := zip.NewWriter(buf)
		for _, item := range source.File {
			// The following statements can be simplified as zw.Copy(item) in go1.17
			if notExist && item.Name == filePath {
				continue
			}
			writer, err := zw.Create(item.Name)
			assert.NoError(t, err)
			readerCloser, err := item.Open()
			assert.NoError(t, err)
			_, err = io.Copy(writer, readerCloser)
			assert.NoError(t, err)
		}
		if !notExist {
			fi, err := zw.Create(filePath)
			assert.NoError(t, err)
			_, err = fi.Write(MacintoshCyrillicCharset)
			assert.NoError(t, err)
		}
		assert.NoError(t, zw.Close())
		return buf
	}
	// Test open workbook with unsupported charset internal XML parts
	for _, defaultXMLPath := range []string{
		defaultXMLPathCalcChain,
		defaultXMLPathStyles,
		defaultXMLPathWorkbookRels,
	} {
		_, err = OpenReader(preset(defaultXMLPath, false))
		assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	}
	// Test open workbook without internal XML parts
	for _, defaultXMLPath := range []string{
		defaultXMLPathCalcChain,
		defaultXMLPathStyles,
		defaultXMLPathWorkbookRels,
	} {
		_, err = OpenReader(preset(defaultXMLPath, true))
		assert.NoError(t, err)
	}

	// Test open spreadsheet with unzip size limit
	_, err = OpenFile(filepath.Join("test", "Book1.xlsx"), Options{UnzipSizeLimit: 100})
	assert.EqualError(t, err, newUnzipSizeLimitError(100).Error())

	// Test open password protected spreadsheet created by Microsoft Office Excel 2010
	f, err := OpenFile(filepath.Join("test", "encryptSHA1.xlsx"), Options{Password: "password"})
	assert.NoError(t, err)
	val, err := f.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "SECRET", val)
	assert.NoError(t, f.Close())

	// Test open password protected spreadsheet created by LibreOffice 7.0.0.3
	f, err = OpenFile(filepath.Join("test", "encryptAES.xlsx"), Options{Password: "password"})
	assert.NoError(t, err)
	val, err = f.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "SECRET", val)
	assert.NoError(t, f.Close())

	// Test open spreadsheet with invalid options
	_, err = OpenReader(bytes.NewReader(oleIdentifier), Options{UnzipSizeLimit: 1, UnzipXMLSizeLimit: 2})
	assert.EqualError(t, err, ErrOptionsUnzipSizeLimit.Error())

	// Test unexpected EOF
	var b bytes.Buffer
	w := gzip.NewWriter(&b)
	defer w.Close()
	w.Flush()

	r, _ := gzip.NewReader(&b)
	defer r.Close()

	_, err = OpenReader(r)
	assert.EqualError(t, err, "unexpected EOF")

	_, err = OpenReader(bytes.NewReader([]byte{
		0x50, 0x4b, 0x03, 0x04, 0x0a, 0x00, 0x09, 0x00, 0x63, 0x00, 0x47, 0xa3, 0xb6, 0x50, 0x00, 0x00,
		0x00, 0x00, 0x1c, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x08, 0x00, 0x0b, 0x00, 0x70, 0x61,
		0x73, 0x73, 0x77, 0x6f, 0x72, 0x64, 0x01, 0x99, 0x07, 0x00, 0x02, 0x00, 0x41, 0x45, 0x03, 0x00,
		0x00, 0x21, 0x06, 0x59, 0xc0, 0x12, 0xf3, 0x19, 0xc7, 0x51, 0xd1, 0xc9, 0x31, 0xcb, 0xcc, 0x8a,
		0xe1, 0x44, 0xe1, 0x56, 0x20, 0x24, 0x1f, 0xba, 0x09, 0xda, 0x53, 0xd5, 0xef, 0x50, 0x4b, 0x07,
		0x08, 0x00, 0x00, 0x00, 0x00, 0x1c, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x50, 0x4b, 0x01,
		0x02, 0x1f, 0x00, 0x0a, 0x00, 0x09, 0x00, 0x63, 0x00, 0x47, 0xa3, 0xb6, 0x50, 0x00, 0x00, 0x00,
		0x00, 0x1c, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x08, 0x00, 0x0b, 0x00, 0x00, 0x00, 0x00,
		0x00, 0x00, 0x00, 0x20, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x70, 0x61, 0x73, 0x73, 0x77,
		0x6f, 0x72, 0x64, 0x01, 0x99, 0x07, 0x00, 0x02, 0x00, 0x41, 0x45, 0x03, 0x00, 0x00, 0x50, 0x4b,
		0x05, 0x06, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x41, 0x00, 0x00, 0x00, 0x5d, 0x00,
		0x00, 0x00, 0x00, 0x00,
	}))
	assert.EqualError(t, err, zip.ErrAlgorithm.Error())
}

func TestBrokenFile(t *testing.T) {
	// Test write file with broken file struct
	f := File{}

	t.Run("SaveWithoutName", func(t *testing.T) {
		assert.EqualError(t, f.Save(), "no path defined for file, consider File.WriteTo or File.Write")
	})

	t.Run("SaveAsEmptyStruct", func(t *testing.T) {
		// Test write file with broken file struct with given path
		assert.NoError(t, f.SaveAs(filepath.Join("test", "BadWorkbook.SaveAsEmptyStruct.xlsx")))
	})

	t.Run("OpenBadWorkbook", func(t *testing.T) {
		// Test set active sheet without BookViews and Sheets maps in xl/workbook.xml
		f3, err := OpenFile(filepath.Join("test", "BadWorkbook.xlsx"))
		f3.GetActiveSheetIndex()
		f3.SetActiveSheet(1)
		assert.NoError(t, err)
		assert.NoError(t, f3.Close())
	})

	t.Run("OpenNotExistsFile", func(t *testing.T) {
		// Test open a spreadsheet file with given illegal path
		_, err := OpenFile(filepath.Join("test", "NotExistsFile.xlsx"))
		if assert.Error(t, err) {
			assert.True(t, os.IsNotExist(err), "Expected os.IsNotExists(err) == true")
		}
	})
}

func TestNewFile(t *testing.T) {
	// Test create a spreadsheet file
	f := NewFile()
	_, err := f.NewSheet("Sheet1")
	assert.NoError(t, err)
	_, err = f.NewSheet("Sheet2")
	assert.NoError(t, err)
	_, err = f.NewSheet("Sheet3")
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellInt("Sheet2", "A23", 56))
	assert.NoError(t, f.SetCellStr("Sheet1", "B20", "42"))
	f.SetActiveSheet(0)

	// Test add picture to sheet with scaling and positioning
	assert.NoError(t, f.AddPicture("Sheet1", "H2", filepath.Join("test", "images", "excel.gif"),
		&GraphicOptions{ScaleX: 0.5, ScaleY: 0.5, Positioning: "absolute"}))

	// Test add picture to worksheet without options
	assert.NoError(t, f.AddPicture("Sheet1", "C2", filepath.Join("test", "images", "excel.png"), nil))

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestNewFile.xlsx")))
	assert.NoError(t, f.Save())
}

func TestSetCellHyperLink(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)
	// Test set cell hyperlink in a work sheet already have hyperlinks
	assert.NoError(t, f.SetCellHyperLink("Sheet1", "B19", "https://github.com/xuri/excelize", "External"))
	// Test add first hyperlink in a work sheet
	assert.NoError(t, f.SetCellHyperLink("Sheet2", "C1", "https://github.com/xuri/excelize", "External"))
	// Test add Location hyperlink in a work sheet
	assert.NoError(t, f.SetCellHyperLink("Sheet2", "D6", "Sheet1!D8", "Location"))
	// Test add Location hyperlink with display & tooltip in a work sheet
	display, tooltip := "Display value", "Hover text"
	assert.NoError(t, f.SetCellHyperLink("Sheet2", "D7", "Sheet1!D9", "Location", HyperlinkOpts{
		Display: &display,
		Tooltip: &tooltip,
	}))
	// Test set cell hyperlink with invalid sheet name
	assert.Equal(t, ErrSheetNameInvalid, f.SetCellHyperLink("Sheet:1", "A1", "Sheet1!D60", "Location"))
	assert.Equal(t, newInvalidLinkTypeError(""), f.SetCellHyperLink("Sheet2", "C3", "Sheet1!D8", ""))
	assert.EqualError(t, f.SetCellHyperLink("Sheet2", "", "Sheet1!D60", "Location"), `invalid cell name ""`)
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellHyperLink.xlsx")))
	assert.NoError(t, f.Close())

	f = NewFile()
	_, err = f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).Hyperlinks = &xlsxHyperlinks{Hyperlink: make([]xlsxHyperlink, 65530)}
	assert.EqualError(t, f.SetCellHyperLink("Sheet1", "A65531", "https://github.com/xuri/excelize", "External"), ErrTotalSheetHyperlinks.Error())

	f = NewFile()
	_, err = f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	ws, ok = f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).MergeCells = &xlsxMergeCells{Cells: []*xlsxMergeCell{{Ref: "A:A"}}}
	err = f.SetCellHyperLink("Sheet1", "A1", "https://github.com/xuri/excelize", "External")
	assert.EqualError(t, err, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())

	// Test update cell hyperlink
	f = NewFile()
	assert.NoError(t, f.SetCellHyperLink("Sheet1", "A1", "https://github.com", "External"))
	assert.NoError(t, f.SetCellHyperLink("Sheet1", "A1", "https://github.com/xuri/excelize", "External"))
	link, target, err := f.GetCellHyperLink("Sheet1", "A1")
	assert.Equal(t, link, true)
	assert.Equal(t, "https://github.com/xuri/excelize", target)
	assert.NoError(t, err)

	// Test remove hyperlink for a cell
	f = NewFile()
	assert.NoError(t, f.SetCellHyperLink("Sheet1", "A1", "Sheet1!D8", "Location"))
	ws, ok = f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).Hyperlinks.Hyperlink[0].Ref = "A1:D4"
	assert.NoError(t, f.SetCellHyperLink("Sheet1", "B2", "", "None"))
	// Test remove hyperlink for a cell with invalid cell reference
	assert.NoError(t, f.SetCellHyperLink("Sheet1", "A1", "Sheet1!D8", "Location"))
	ws.(*xlsxWorksheet).Hyperlinks.Hyperlink[0].Ref = "A:A"
	assert.Error(t, f.SetCellHyperLink("Sheet1", "B2", "", "None"), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")))
}

func TestGetCellHyperLink(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)

	_, _, err = f.GetCellHyperLink("Sheet1", "")
	assert.EqualError(t, err, `invalid cell name ""`)

	link, target, err := f.GetCellHyperLink("Sheet1", "A22")
	assert.NoError(t, err)
	assert.Equal(t, link, true)
	assert.Equal(t, target, "https://github.com/xuri/excelize")

	link, target, err = f.GetCellHyperLink("Sheet2", "D6")
	assert.NoError(t, err)
	assert.Equal(t, link, false)
	assert.Equal(t, target, "")

	link, target, err = f.GetCellHyperLink("Sheet3", "H3")
	assert.EqualError(t, err, "sheet Sheet3 does not exist")
	assert.Equal(t, link, false)
	assert.Equal(t, target, "")

	assert.NoError(t, f.Close())

	f = NewFile()
	_, err = f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).Hyperlinks = &xlsxHyperlinks{
		Hyperlink: []xlsxHyperlink{{Ref: "A1"}},
	}
	link, target, err = f.GetCellHyperLink("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, link, true)
	assert.Equal(t, target, "")

	ws, ok = f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).Hyperlinks = &xlsxHyperlinks{Hyperlink: []xlsxHyperlink{{Ref: "A:A"}}}
	link, target, err = f.GetCellHyperLink("Sheet1", "A1")
	assert.EqualError(t, err, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	assert.Equal(t, link, false)
	assert.Equal(t, target, "")

	// Test get cell hyperlink with invalid sheet name
	_, _, err = f.GetCellHyperLink("Sheet:1", "A1")
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())
}

func TestSetSheetBackground(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)
	assert.NoError(t, f.SetSheetBackground("Sheet2", filepath.Join("test", "images", "background.jpg")))
	assert.NoError(t, f.SetSheetBackground("Sheet2", filepath.Join("test", "images", "background.jpg")))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetSheetBackground.xlsx")))
	assert.NoError(t, f.Close())
}

func TestSetSheetBackgroundErrors(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)

	err = f.SetSheetBackground("Sheet2", filepath.Join("test", "not_exists", "not_exists.png"))
	if assert.Error(t, err) {
		assert.True(t, os.IsNotExist(err), "Expected os.IsNotExists(err) == true")
	}

	err = f.SetSheetBackground("Sheet2", filepath.Join("test", "Book1.xlsx"))
	assert.EqualError(t, err, ErrImgExt.Error())
	// Test set sheet background on not exist worksheet
	err = f.SetSheetBackground("SheetN", filepath.Join("test", "images", "background.jpg"))
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Test set sheet background with invalid sheet name
	assert.EqualError(t, f.SetSheetBackground("Sheet:1", filepath.Join("test", "images", "background.jpg")), ErrSheetNameInvalid.Error())
	assert.NoError(t, f.Close())

	// Test set sheet background with unsupported charset content types
	f = NewFile()
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, f.SetSheetBackground("Sheet1", filepath.Join("test", "images", "background.jpg")), "XML syntax error on line 1: invalid UTF-8")
}

// TestWriteArrayFormula tests the extended options of SetCellFormula by writing
// an array function to a workbook. In the resulting file, the lines 2 and 3 as
// well as 4 and 5 should have matching contents
func TestWriteArrayFormula(t *testing.T) {
	cell := func(col, row int) string {
		c, err := CoordinatesToCellName(col, row)
		if err != nil {
			t.Fatal(err)
		}
		return c
	}

	f := NewFile()

	sample := []string{"Sample 1", "Sample 2", "Sample 3"}
	values := []int{1855, 1709, 1462, 1115, 1524, 625, 773, 126, 1027, 1696, 1078, 1917, 1109, 1753, 1884, 659, 994, 1911, 1925, 899, 196, 244, 1488, 1056, 1986, 66, 784, 725, 767, 1722, 1541, 1026, 1455, 264, 1538, 877, 1581, 1098, 383, 762, 237, 493, 29, 1923, 474, 430, 585, 688, 308, 200, 1259, 622, 798, 1048, 996, 601, 582, 332, 377, 805, 250, 1860, 1360, 840, 911, 1346, 1651, 1651, 665, 584, 1057, 1145, 925, 1752, 202, 149, 1917, 1398, 1894, 818, 714, 624, 1085, 1566, 635, 78, 313, 1686, 1820, 494, 614, 1913, 271, 1016, 338, 1301, 489, 1733, 1483, 1141}
	assoc := []int{2, 0, 0, 0, 0, 1, 1, 0, 0, 1, 2, 2, 2, 1, 1, 1, 1, 0, 0, 0, 1, 0, 2, 0, 2, 1, 2, 2, 2, 1, 0, 1, 0, 1, 1, 2, 0, 2, 1, 0, 2, 1, 0, 1, 0, 0, 2, 0, 2, 2, 1, 2, 2, 1, 2, 2, 1, 2, 1, 2, 2, 1, 1, 1, 0, 1, 0, 2, 0, 0, 1, 2, 1, 0, 1, 0, 0, 2, 1, 1, 2, 0, 2, 1, 0, 2, 2, 2, 1, 0, 0, 1, 1, 1, 2, 0, 2, 0, 1, 1}
	if len(values) != len(assoc) {
		t.Fatal("values and assoc must be of same length")
	}

	// Average calculates the average of the n-th sample (0 <= n < len(sample)).
	average := func(n int) int {
		sum := 0
		count := 0
		for i := 0; i != len(values); i++ {
			if assoc[i] == n {
				sum += values[i]
				count++
			}
		}

		return int(math.Round(float64(sum) / float64(count)))
	}

	// Stdev calculates the standard deviation of the n-th sample (0 <= n < len(sample)).
	stdev := func(n int) int {
		avg := average(n)

		sum := 0
		count := 0
		for i := 0; i != len(values); i++ {
			if assoc[i] == n {
				sum += (values[i] - avg) * (values[i] - avg)
				count++
			}
		}

		return int(math.Round(math.Sqrt(float64(sum) / float64(count))))
	}

	// Line 2 contains the results of AVERAGEIF
	assert.NoError(t, f.SetCellStr("Sheet1", "A2", "Average"))

	// Line 3 contains the average that was calculated in Go
	assert.NoError(t, f.SetCellStr("Sheet1", "A3", "Average (calculated)"))

	// Line 4 contains the results of the array function that calculates the standard deviation
	assert.NoError(t, f.SetCellStr("Sheet1", "A4", "Std. deviation"))

	// Line 5 contains the standard deviations calculated in Go
	assert.NoError(t, f.SetCellStr("Sheet1", "A5", "Std. deviation (calculated)"))

	assert.NoError(t, f.SetCellStr("Sheet1", "B1", sample[0]))
	assert.NoError(t, f.SetCellStr("Sheet1", "C1", sample[1]))
	assert.NoError(t, f.SetCellStr("Sheet1", "D1", sample[2]))

	firstResLine := 8
	assert.NoError(t, f.SetCellStr("Sheet1", cell(1, firstResLine-1), "Result Values"))
	assert.NoError(t, f.SetCellStr("Sheet1", cell(2, firstResLine-1), "Sample"))

	for i := 0; i != len(values); i++ {
		valCell := cell(1, i+firstResLine)
		assocCell := cell(2, i+firstResLine)

		assert.NoError(t, f.SetCellInt("Sheet1", valCell, int64(values[i])))
		assert.NoError(t, f.SetCellStr("Sheet1", assocCell, sample[assoc[i]]))
	}

	valRange := fmt.Sprintf("$A$%d:$A$%d", firstResLine, len(values)+firstResLine-1)
	assocRange := fmt.Sprintf("$B$%d:$B$%d", firstResLine, len(values)+firstResLine-1)

	for i := 0; i != len(sample); i++ {
		nameCell := cell(i+2, 1)
		avgCell := cell(i+2, 2)
		calcAvgCell := cell(i+2, 3)
		stdevCell := cell(i+2, 4)
		calcStdevCell := cell(i+2, 5)

		assert.NoError(t, f.SetCellInt("Sheet1", calcAvgCell, int64(average(i))))
		assert.NoError(t, f.SetCellInt("Sheet1", calcStdevCell, int64(stdev(i))))

		// Average can be done with AVERAGEIF
		assert.NoError(t, f.SetCellFormula("Sheet1", avgCell, fmt.Sprintf("ROUND(AVERAGEIF(%s,%s,%s),0)", assocRange, nameCell, valRange)))

		ref := stdevCell + ":" + stdevCell
		arr := STCellFormulaTypeArray
		// Use an array formula for standard deviation
		assert.NoError(t, f.SetCellFormula("Sheet1", stdevCell, fmt.Sprintf("ROUND(STDEVP(IF(%s=%s,%s)),0)", assocRange, nameCell, valRange),
			FormulaOpts{}, FormulaOpts{Type: &arr}, FormulaOpts{Ref: &ref}))
	}

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestWriteArrayFormula.xlsx")))
}

func TestSetCellStyleAlignment(t *testing.T) {
	f, err := prepareTestBook1()
	assert.NoError(t, err)

	var style int
	style, err = f.NewStyle(&Style{Alignment: &Alignment{Horizontal: "center", Indent: 1, JustifyLastLine: true, ReadingOrder: 0, RelativeIndent: 1, ShrinkToFit: true, TextRotation: 45, Vertical: "top", WrapText: true}})
	assert.NoError(t, err)

	assert.NoError(t, f.SetCellStyle("Sheet1", "A22", "A22", style))

	// Test set cell style with given illegal rows number
	assert.EqualError(t, f.SetCellStyle("Sheet1", "A", "A22", style), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	assert.EqualError(t, f.SetCellStyle("Sheet1", "A22", "A", style), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	// Test set cell style with invalid sheet name
	assert.EqualError(t, f.SetCellStyle("Sheet:1", "A1", "A2", style), ErrSheetNameInvalid.Error())
	// Test get cell style with given illegal rows number
	index, err := f.GetCellStyle("Sheet1", "A")
	assert.Equal(t, 0, index)
	assert.EqualError(t, err, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())

	// Test get cell style with invalid sheet name
	_, err = f.GetCellStyle("Sheet:1", "A1")
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleAlignment.xlsx")))
}

func TestSetCellStyleBorder(t *testing.T) {
	f, err := prepareTestBook1()
	assert.NoError(t, err)

	var style int

	// Test set border on overlapping range with vertical variants shading styles gradient fill
	style, err = f.NewStyle(&Style{
		Border: []Border{
			{Type: "left", Color: "0000FF", Style: 3},
			{Type: "top", Color: "00FF00", Style: 4},
			{Type: "bottom", Color: "FFFF00", Style: 5},
			{Type: "right", Color: "FF0000", Style: 6},
			{Type: "diagonalDown", Color: "A020F0", Style: 7},
			{Type: "diagonalUp", Color: "A020F0", Style: 8},
		},
	})
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "J21", "L25", style))

	style, err = f.NewStyle(&Style{Border: []Border{{Type: "left", Color: "0000FF", Style: 2}, {Type: "top", Color: "00FF00", Style: 3}, {Type: "bottom", Color: "FFFF00", Style: 4}, {Type: "right", Color: "FF0000", Style: 5}, {Type: "diagonalDown", Color: "A020F0", Style: 6}, {Type: "diagonalUp", Color: "A020F0", Style: 7}}, Fill: Fill{Type: "gradient", Color: []string{"FFFFFF", "E0EBF5"}, Shading: 1}})
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "M28", "K24", style))

	style, err = f.NewStyle(&Style{Border: []Border{{Type: "left", Color: "0000FF", Style: 2}, {Type: "top", Color: "00FF00", Style: 3}, {Type: "bottom", Color: "FFFF00", Style: 4}, {Type: "right", Color: "FF0000", Style: 5}, {Type: "diagonalDown", Color: "A020F0", Style: 6}, {Type: "diagonalUp", Color: "A020F0", Style: 7}}, Fill: Fill{Type: "gradient", Color: []string{"FFFFFF", "E0EBF5"}, Shading: 4}})
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "M28", "K24", style))

	// Test set border and solid style pattern fill for a single cell
	style, err = f.NewStyle(&Style{
		Border: []Border{
			{
				Type:  "left",
				Color: "0000FF",
				Style: 8,
			},
			{
				Type:  "top",
				Color: "00FF00",
				Style: 9,
			},
			{
				Type:  "bottom",
				Color: "FFFF00",
				Style: 10,
			},
			{
				Type:  "right",
				Color: "FF0000",
				Style: 11,
			},
			{
				Type:  "diagonalDown",
				Color: "A020F0",
				Style: 12,
			},
			{
				Type:  "diagonalUp",
				Color: "A020F0",
				Style: 13,
			},
		},
		Fill: Fill{
			Type:    "pattern",
			Color:   []string{"E0EBF5"},
			Pattern: 1,
		},
	})
	assert.NoError(t, err)

	assert.NoError(t, f.SetCellStyle("Sheet1", "O22", "O22", style))

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleBorder.xlsx")))
}

func TestSetCellStyleBorderErrors(t *testing.T) {
	f, err := prepareTestBook1()
	assert.NoError(t, err)

	// Set border with invalid style index number
	_, err = f.NewStyle(&Style{Border: []Border{{Type: "left", Color: "0000FF", Style: -1}, {Type: "top", Color: "00FF00", Style: 14}, {Type: "bottom", Color: "FFFF00", Style: 5}, {Type: "right", Color: "FF0000", Style: 6}, {Type: "diagonalDown", Color: "A020F0", Style: 9}, {Type: "diagonalUp", Color: "A020F0", Style: 8}}})
	assert.NoError(t, err)
}

func TestSetCellStyleNumberFormat(t *testing.T) {
	f, err := prepareTestBook1()
	assert.NoError(t, err)

	// Test only set fill and number format for a cell
	col := []string{"L", "M", "N", "O", "P"}
	idxTbl := []int{0, 1, 2, 3, 4, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49}
	value := []string{"37947.7500001", "-37947.7500001", "0.007", "2.1", "String"}
	expected := [][]string{
		{"37947.75", "37948", "37947.75", "37,948", "37,947.75", "3794775%", "3794775.00%", "3.79E+04", "37947 3/4", "37947 3/4", "11-22-03", "22-Nov-03", "22-Nov", "Nov-03", "6:00 PM", "6:00:00 PM", "18:00", "18:00:00", "11/22/03 18:00", "37,948 ", "37,948 ", "37,947.75 ", "37,947.75 ", " 37,948 ", " $37,948 ", " 37,947.75 ", " $37,947.75 ", "00:00", "910746:00:00", "00:00.0", "37947.7500001", "37947.7500001"},
		{"-37947.75", "-37948", "-37947.75", "-37,948", "-37,947.75", "-3794775%", "-3794775.00%", "-3.79E+04", "-37947 3/4", "-37947 3/4", "-37947.7500001", "-37947.7500001", "-37947.7500001", "-37947.7500001", "-37947.7500001", "-37947.7500001", "-37947.7500001", "-37947.7500001", "-37947.7500001", "(37,948)", "(37,948)", "(37,947.75)", "(37,947.75)", " (37,948)", " $(37,948)", " (37,947.75)", " $(37,947.75)", "-37947.7500001", "-37947.7500001", "-37947.7500001", "-37947.7500001", "-37947.7500001"},
		{"0.007", "0", "0.01", "0", "0.01", "1%", "0.70%", "7.00E-03", "0    ", "0    ", "12-30-99", "30-Dec-99", "30-Dec", "Dec-99", "12:10 AM", "12:10:05 AM", "00:10", "00:10:05", "12/30/99 00:10", "0 ", "0 ", "0.01 ", "0.01 ", " 0 ", " $0 ", " 0.01 ", " $0.01 ", "10:05", "0:10:05", "10:04.8", "0.007", "0.007"},
		{"2.1", "2", "2.10", "2", "2.10", "210%", "210.00%", "2.10E+00", "2 1/9", "2 1/10", "01-01-00", "1-Jan-00", "1-Jan", "Jan-00", "2:24 AM", "2:24:00 AM", "02:24", "02:24:00", "1/1/00 02:24", "2 ", "2 ", "2.10 ", "2.10 ", " 2 ", " $2 ", " 2.10 ", " $2.10 ", "24:00", "50:24:00", "24:00.0", "2.1", "2.1"},
		{"String", "String", "String", "String", "String", "String", "String", "String", "String", "String", "String", "String", "String", "String", "String", "String", "String", "String", "String", "String", "String", "String", "String", " String ", " String ", " String ", " String ", "String", "String", "String", "String", "String"},
	}

	for c, v := range value {
		for r, idx := range idxTbl {
			cell := col[c] + strconv.Itoa(r+1)
			var val float64
			val, err = strconv.ParseFloat(v, 64)
			if err != nil {
				assert.NoError(t, f.SetCellValue("Sheet2", cell, v))
			} else {
				assert.NoError(t, f.SetCellValue("Sheet2", cell, val))
			}
			style, err := f.NewStyle(&Style{Fill: Fill{Type: "gradient", Color: []string{"FFFFFF", "E0EBF5"}, Shading: 5}, NumFmt: idx})
			if !assert.NoError(t, err) {
				t.FailNow()
			}
			assert.NoError(t, f.SetCellStyle("Sheet2", cell, cell, style))
			cellValue, err := f.GetCellValue("Sheet2", cell)
			assert.Equal(t, expected[c][r], cellValue, fmt.Sprintf("Sheet2!%s value: %s, number format: %s c: %d r: %d", cell, value[c], builtInNumFmt[idx], c, r))
			assert.NoError(t, err)
		}
	}
	var style int
	style, err = f.NewStyle(&Style{NumFmt: -1})
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet2", "L33", "L33", style))

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleNumberFormat.xlsx")))

	// Test get cell value with built-in number format code 22 with custom short date pattern
	f = NewFile(Options{ShortDatePattern: "yyyy-m-dd"})
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 45074.625694444447))
	style, err = f.NewStyle(&Style{NumFmt: 22})
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "A1", "A1", style))
	cellValue, err := f.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "2023-5-28 15:01", cellValue)
}

func TestSetCellStyleCurrencyNumberFormat(t *testing.T) {
	t.Run("TestBook3", func(t *testing.T) {
		f, err := prepareTestBook3()
		assert.NoError(t, err)

		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 56))
		assert.NoError(t, f.SetCellValue("Sheet1", "A2", -32.3))
		var style int
		style, err = f.NewStyle(&Style{NumFmt: 188, DecimalPlaces: intPtr(-1)})
		assert.NoError(t, err)

		assert.NoError(t, f.SetCellStyle("Sheet1", "A1", "A1", style))
		style, err = f.NewStyle(&Style{NumFmt: 188, DecimalPlaces: intPtr(31), NegRed: true})
		assert.NoError(t, err)

		assert.NoError(t, f.SetCellStyle("Sheet1", "A2", "A2", style))

		assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleCurrencyNumberFormat.TestBook3.xlsx")))
	})

	t.Run("TestBook4", func(t *testing.T) {
		f, err := prepareTestBook4()
		assert.NoError(t, err)
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 42920.5))
		assert.NoError(t, f.SetCellValue("Sheet1", "A2", 42920.5))

		_, err = f.NewStyle(&Style{NumFmt: 26})
		assert.NoError(t, err)

		style, err := f.NewStyle(&Style{NumFmt: 27})
		assert.NoError(t, err)

		assert.NoError(t, f.SetCellStyle("Sheet1", "A1", "A1", style))
		style, err = f.NewStyle(&Style{NumFmt: 31})
		assert.NoError(t, err)

		assert.NoError(t, f.SetCellStyle("Sheet1", "A2", "A2", style))

		style, err = f.NewStyle(&Style{NumFmt: 71})
		assert.NoError(t, err)
		assert.NoError(t, f.SetCellStyle("Sheet1", "A2", "A2", style))

		assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleCurrencyNumberFormat.TestBook4.xlsx")))
	})
}

func TestSetCellStyleLangNumberFormat(t *testing.T) {
	rawCellValues := make([][]string, 42)
	for i := 0; i < 42; i++ {
		rawCellValues[i] = []string{"45162"}
	}
	for lang, expected := range map[CultureName][][]string{
		CultureNameUnknown: rawCellValues,
		CultureNameEnUS:    {{"8/24/23"}, {"8/24/23"}, {"8/24/23"}, {"8/24/23"}, {"8/24/23"}, {"0:00:00"}, {"0:00:00"}, {"0:00:00"}, {"0:00:00"}, {"45162"}, {"8/24/23"}, {"8/24/23"}, {"8/24/23"}, {"8/24/23"}, {"8/24/23"}, {"8/24/23"}, {"8/24/23"}, {"8/24/23"}, {"8/24/23"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}},
		CultureNameJaJP:    {{"R5.8.24"}, {"令和5年8月24日"}, {"令和5年8月24日"}, {"8/24/23"}, {"2023年8月24日"}, {"0時00分"}, {"0時00分00秒"}, {"2023年8月"}, {"8月24日"}, {"R5.8.24"}, {"R5.8.24"}, {"令和5年8月24日"}, {"2023年8月"}, {"8月24日"}, {"令和5年8月24日"}, {"2023年8月"}, {"8月24日"}, {"R5.8.24"}, {"令和5年8月24日"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}},
		CultureNameKoKR:    {{"4356年 08月 24日"}, {"08-24"}, {"08-24"}, {"08-24-56"}, {"4356년 08월 24일"}, {"0시 00분"}, {"0시 00분 00초"}, {"4356-08-24"}, {"4356-08-24"}, {"4356年 08月 24日"}, {"4356年 08月 24日"}, {"08-24"}, {"4356-08-24"}, {"4356-08-24"}, {"08-24"}, {"4356-08-24"}, {"4356-08-24"}, {"4356年 08月 24日"}, {"08-24"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}},
		CultureNameZhCN:    {{"2023年8月"}, {"8月24日"}, {"8月24日"}, {"8/24/23"}, {"2023年8月24日"}, {"0时00分"}, {"0时00分00秒"}, {"上午12时00分"}, {"上午12时00分00秒"}, {"2023年8月"}, {"2023年8月"}, {"8月24日"}, {"2023年8月"}, {"8月24日"}, {"8月24日"}, {"上午12时00分"}, {"上午12时00分00秒"}, {"2023年8月"}, {"8月24日"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}},
		CultureNameZhTW:    {{"112/8/24"}, {"112年8月24日"}, {"112年8月24日"}, {"8/24/23"}, {"2023年8月24日"}, {"00時00分"}, {"00時00分00秒"}, {"上午12時00分"}, {"上午12時00分00秒"}, {"112/8/24"}, {"112/8/24"}, {"112年8月24日"}, {"上午12時00分"}, {"上午12時00分00秒"}, {"112年8月24日"}, {"上午12時00分"}, {"上午12時00分00秒"}, {"112/8/24"}, {"112年8月24日"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}},
	} {
		f, err := prepareTestBook5(Options{CultureInfo: lang})
		assert.NoError(t, err)
		rows, err := f.GetRows("Sheet1")
		assert.NoError(t, err)
		assert.Equal(t, expected, rows)
		assert.NoError(t, f.Close())
	}
	// Test apply language number format code with date and time pattern
	for lang, expected := range map[CultureName][][]string{
		CultureNameEnUS: {{"2023-8-24"}, {"2023-8-24"}, {"2023-8-24"}, {"2023-8-24"}, {"2023-8-24"}, {"00:00:00"}, {"00:00:00"}, {"00:00:00"}, {"00:00:00"}, {"45162"}, {"2023-8-24"}, {"2023-8-24"}, {"2023-8-24"}, {"2023-8-24"}, {"2023-8-24"}, {"2023-8-24"}, {"2023-8-24"}, {"2023-8-24"}, {"2023-8-24"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}},
		CultureNameJaJP: {{"R5.8.24"}, {"令和5年8月24日"}, {"令和5年8月24日"}, {"2023-8-24"}, {"2023年8月24日"}, {"00:00:00"}, {"00:00:00"}, {"2023年8月"}, {"8月24日"}, {"R5.8.24"}, {"R5.8.24"}, {"令和5年8月24日"}, {"2023年8月"}, {"8月24日"}, {"令和5年8月24日"}, {"2023年8月"}, {"8月24日"}, {"R5.8.24"}, {"令和5年8月24日"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}},
		CultureNameKoKR: {{"4356年 08月 24日"}, {"08-24"}, {"08-24"}, {"4356-8-24"}, {"4356년 08월 24일"}, {"00:00:00"}, {"00:00:00"}, {"4356-08-24"}, {"4356-08-24"}, {"4356年 08月 24日"}, {"4356年 08月 24日"}, {"08-24"}, {"4356-08-24"}, {"4356-08-24"}, {"08-24"}, {"4356-08-24"}, {"4356-08-24"}, {"4356年 08月 24日"}, {"08-24"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}},
		CultureNameZhCN: {{"2023年8月"}, {"8月24日"}, {"8月24日"}, {"2023-8-24"}, {"2023年8月24日"}, {"00:00:00"}, {"00:00:00"}, {"上午12时00分"}, {"上午12时00分00秒"}, {"2023年8月"}, {"2023年8月"}, {"8月24日"}, {"2023年8月"}, {"8月24日"}, {"8月24日"}, {"上午12时00分"}, {"上午12时00分00秒"}, {"2023年8月"}, {"8月24日"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}},
		CultureNameZhTW: {{"112/8/24"}, {"112年8月24日"}, {"112年8月24日"}, {"2023-8-24"}, {"2023年8月24日"}, {"00:00:00"}, {"00:00:00"}, {"上午12時00分"}, {"上午12時00分00秒"}, {"112/8/24"}, {"112/8/24"}, {"112年8月24日"}, {"上午12時00分"}, {"上午12時00分00秒"}, {"112年8月24日"}, {"上午12時00分"}, {"上午12時00分00秒"}, {"112/8/24"}, {"112年8月24日"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}, {"45162"}},
	} {
		f, err := prepareTestBook5(Options{CultureInfo: lang, ShortDatePattern: "yyyy-M-d", LongTimePattern: "hh:mm:ss"})
		assert.NoError(t, err)
		rows, err := f.GetRows("Sheet1")
		assert.NoError(t, err)
		assert.Equal(t, expected, rows)
		assert.NoError(t, f.Close())
	}
	// Test open workbook with invalid date and time pattern options
	_, err := OpenFile(filepath.Join("test", "Book1.xlsx"), Options{LongDatePattern: "0.00"})
	assert.Equal(t, ErrUnsupportedNumberFormat, err)
	_, err = OpenFile(filepath.Join("test", "Book1.xlsx"), Options{LongTimePattern: "0.00"})
	assert.Equal(t, ErrUnsupportedNumberFormat, err)
	_, err = OpenFile(filepath.Join("test", "Book1.xlsx"), Options{ShortDatePattern: "0.00"})
	assert.Equal(t, ErrUnsupportedNumberFormat, err)
}

func TestSetCellStyleCustomNumberFormat(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 42920.5))
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", 42920.5))
	customNumFmt := "[$-380A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@"
	style, err := f.NewStyle(&Style{CustomNumFmt: &customNumFmt})
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "A1", "A1", style))
	style, err = f.NewStyle(&Style{CustomNumFmt: &customNumFmt, Font: &Font{Color: "9A0511"}})
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "A2", "A2", style))

	customNumFmt = "[$-380A]dddd\\,\\ dd\" de \"mmmm\" de \"yy;@"
	_, err = f.NewStyle(&Style{CustomNumFmt: &customNumFmt})
	assert.NoError(t, err)
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleCustomNumberFormat.xlsx")))
}

func TestSetCellStyleFill(t *testing.T) {
	f, err := prepareTestBook1()
	assert.NoError(t, err)

	var style int
	// Test set fill for cell with invalid parameter
	style, err = f.NewStyle(&Style{Fill: Fill{Type: "gradient", Color: []string{"FFFFFF", "E0EBF5"}, Shading: 6}})
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "O23", "O23", style))

	style, err = f.NewStyle(&Style{Fill: Fill{Type: "gradient", Color: []string{"FFFFFF"}, Shading: 1}})
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "O23", "O23", style))

	style, err = f.NewStyle(&Style{Fill: Fill{Type: "pattern", Color: []string{}, Shading: 1}})
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "O23", "O23", style))

	style, err = f.NewStyle(&Style{Fill: Fill{Type: "pattern", Color: []string{"E0EBF5"}, Pattern: 19}})
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellStyle("Sheet1", "O23", "O23", style))

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleFill.xlsx")))
}

func TestSetCellStyleFont(t *testing.T) {
	f, err := prepareTestBook1()
	assert.NoError(t, err)

	var style int
	style, err = f.NewStyle(&Style{Font: &Font{Bold: true, Italic: true, Family: "Times New Roman", Size: 36, Color: "777777", Underline: "single"}})
	assert.NoError(t, err)

	assert.NoError(t, f.SetCellStyle("Sheet2", "A1", "A1", style))

	style, err = f.NewStyle(&Style{Font: &Font{Italic: true, Underline: "double"}})
	assert.NoError(t, err)

	assert.NoError(t, f.SetCellStyle("Sheet2", "A2", "A2", style))

	style, err = f.NewStyle(&Style{Font: &Font{Bold: true}})
	assert.NoError(t, err)

	assert.NoError(t, f.SetCellStyle("Sheet2", "A3", "A3", style))

	style, err = f.NewStyle(&Style{Font: &Font{Bold: true, Family: "", Size: 0, Color: "", Underline: ""}})
	assert.NoError(t, err)

	assert.NoError(t, f.SetCellStyle("Sheet2", "A4", "A4", style))

	style, err = f.NewStyle(&Style{Font: &Font{Color: "777777", Strike: true}})
	assert.NoError(t, err)

	assert.NoError(t, f.SetCellStyle("Sheet2", "A5", "A5", style))

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetCellStyleFont.xlsx")))
}

func TestSetCellStyleProtection(t *testing.T) {
	f, err := prepareTestBook1()
	assert.NoError(t, err)

	var style int
	style, err = f.NewStyle(&Style{Protection: &Protection{Hidden: true, Locked: true}})
	assert.NoError(t, err)

	assert.NoError(t, f.SetCellStyle("Sheet2", "A6", "A6", style))
	err = f.SaveAs(filepath.Join("test", "TestSetCellStyleProtection.xlsx"))
	assert.NoError(t, err)
}

func TestSetDeleteSheet(t *testing.T) {
	t.Run("TestBook3", func(t *testing.T) {
		f, err := prepareTestBook3()
		assert.NoError(t, err)

		assert.NoError(t, f.DeleteSheet("Sheet3"))
		assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetDeleteSheet.TestBook3.xlsx")))
	})

	t.Run("TestBook4", func(t *testing.T) {
		f, err := prepareTestBook4()
		assert.NoError(t, err)
		assert.NoError(t, f.DeleteSheet("Sheet1"))
		assert.NoError(t, f.AddComment("Sheet1", Comment{Cell: "A1", Author: "Excelize", Paragraph: []RichTextRun{{Text: "Excelize: ", Font: &Font{Bold: true}}, {Text: "This is a comment."}}}))
		assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetDeleteSheet.TestBook4.xlsx")))
	})
}

func TestSheetVisibility(t *testing.T) {
	f, err := prepareTestBook1()
	assert.NoError(t, err)

	assert.NoError(t, f.SetSheetVisible("Sheet2", false))
	assert.NoError(t, f.SetSheetVisible("Sheet2", false, true))
	assert.NoError(t, f.SetSheetVisible("Sheet1", false))
	assert.NoError(t, f.SetSheetVisible("Sheet1", true))
	visible, err := f.GetSheetVisible("Sheet1")
	assert.Equal(t, true, visible)
	assert.NoError(t, err)
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSheetVisibility.xlsx")))
}

func TestCopySheet(t *testing.T) {
	f, err := prepareTestBook1()
	assert.NoError(t, err)

	idx, err := f.NewSheet("CopySheet")
	assert.NoError(t, err)
	assert.NoError(t, f.CopySheet(0, idx))

	assert.NoError(t, f.SetCellValue("CopySheet", "F1", "Hello"))
	val, err := f.GetCellValue("Sheet1", "F1")
	assert.NoError(t, err)
	assert.NotEqual(t, "Hello", val)

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestCopySheet.xlsx")))
}

func TestCopySheetError(t *testing.T) {
	f, err := prepareTestBook1()
	assert.NoError(t, err)
	assert.EqualError(t, f.copySheet(-1, -2), ErrSheetNameBlank.Error())
	assert.EqualError(t, f.CopySheet(-1, -2), ErrSheetIdx.Error())
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestCopySheetError.xlsx")))
}

func TestGetSheetComments(t *testing.T) {
	f := NewFile()
	assert.Equal(t, "", f.getSheetComments("sheet0"))
}

func TestGetActiveSheetIndex(t *testing.T) {
	f := NewFile()
	f.WorkBook.BookViews = nil
	assert.Equal(t, 0, f.GetActiveSheetIndex())
}

func TestRelsWriter(t *testing.T) {
	f := NewFile()
	f.Relationships.Store("xl/worksheets/sheet/rels/sheet1.xml.rel", &xlsxRelationships{})
	f.relsWriter()
}

func TestConditionalFormat(t *testing.T) {
	f := NewFile()
	sheet1 := f.GetSheetName(0)

	assert.NoError(t, fillCells(f, sheet1, 10, 15))

	var format1, format2, format3, format4 int
	var err error
	// Rose format for bad conditional
	format1, err = f.NewConditionalStyle(&Style{Font: &Font{Color: "9A0511"}, Fill: Fill{Type: "pattern", Color: []string{"FEC7CE"}, Pattern: 1}})
	assert.NoError(t, err)

	// Light yellow format for neutral conditional
	format2, err = f.NewConditionalStyle(&Style{Fill: Fill{Type: "pattern", Color: []string{"FEEAA0"}, Pattern: 1}})
	assert.NoError(t, err)

	// Light green format for good conditional
	format3, err = f.NewConditionalStyle(&Style{Font: &Font{Color: "09600B"}, Fill: Fill{Type: "pattern", Color: []string{"C7EECF"}, Pattern: 1}})
	assert.NoError(t, err)

	// conditional style with align and left border
	format4, err = f.NewConditionalStyle(&Style{Alignment: &Alignment{WrapText: true}, Border: []Border{{Type: "left", Color: "000000", Style: 1}}})
	assert.NoError(t, err)

	// Color scales: 2 color
	assert.NoError(t, f.SetConditionalFormat(sheet1, "A1:A10",
		[]ConditionalFormatOptions{
			{
				Type:     "2_color_scale",
				Criteria: "=",
				MinType:  "min",
				MaxType:  "max",
				MinColor: "#F8696B",
				MaxColor: "#63BE7B",
			},
		},
	))
	// Color scales: 3 color
	assert.NoError(t, f.SetConditionalFormat(sheet1, "B1:B10",
		[]ConditionalFormatOptions{
			{
				Type:     "3_color_scale",
				Criteria: "=",
				MinType:  "min",
				MidType:  "percentile",
				MaxType:  "max",
				MinColor: "#F8696B",
				MidColor: "#FFEB84",
				MaxColor: "#63BE7B",
			},
		},
	))
	// Highlight cells rules: between...
	assert.NoError(t, f.SetConditionalFormat(sheet1, "C1:C10",
		[]ConditionalFormatOptions{
			{
				Type:     "cell",
				Criteria: "between",
				Format:   &format1,
				MinValue: "6",
				MaxValue: "8",
			},
		},
	))
	// Highlight cells rules: Greater Than...
	assert.NoError(t, f.SetConditionalFormat(sheet1, "D1:D10",
		[]ConditionalFormatOptions{
			{
				Type:     "cell",
				Criteria: ">",
				Format:   &format3,
				Value:    "6",
			},
		},
	))
	// Highlight cells rules: Equal To...
	assert.NoError(t, f.SetConditionalFormat(sheet1, "E1:E10",
		[]ConditionalFormatOptions{
			{
				Type:     "top",
				Criteria: "=",
				Format:   &format3,
			},
		},
	))
	// Highlight cells rules: Not Equal To...
	assert.NoError(t, f.SetConditionalFormat(sheet1, "F1:F10",
		[]ConditionalFormatOptions{
			{
				Type:     "unique",
				Criteria: "=",
				Format:   &format2,
			},
		},
	))
	// Highlight cells rules: Duplicate Values...
	assert.NoError(t, f.SetConditionalFormat(sheet1, "G1:G10",
		[]ConditionalFormatOptions{
			{
				Type:     "duplicate",
				Criteria: "=",
				Format:   &format2,
			},
		},
	))
	// Top/Bottom rules: Top 10%.
	assert.NoError(t, f.SetConditionalFormat(sheet1, "H1:H10",
		[]ConditionalFormatOptions{
			{
				Type:     "top",
				Criteria: "=",
				Format:   &format1,
				Value:    "6",
				Percent:  true,
			},
		},
	))
	// Top/Bottom rules: Above Average...
	assert.NoError(t, f.SetConditionalFormat(sheet1, "I1:I10",
		[]ConditionalFormatOptions{
			{
				Type:         "average",
				Criteria:     "=",
				Format:       &format3,
				AboveAverage: true,
			},
		},
	))
	// Top/Bottom rules: Below Average...
	assert.NoError(t, f.SetConditionalFormat(sheet1, "J1:J10",
		[]ConditionalFormatOptions{
			{
				Type:         "average",
				Criteria:     "=",
				Format:       &format1,
				AboveAverage: false,
			},
		},
	))
	// Data Bars: Gradient Fill
	assert.NoError(t, f.SetConditionalFormat(sheet1, "K1:K10",
		[]ConditionalFormatOptions{
			{
				Type:     "data_bar",
				Criteria: "=",
				MinType:  "min",
				MaxType:  "max",
				BarColor: "#638EC6",
			},
		},
	))
	// Use a formula to determine which cells to format
	assert.NoError(t, f.SetConditionalFormat(sheet1, "L1:L10",
		[]ConditionalFormatOptions{
			{
				Type:     "formula",
				Criteria: "L2<3",
				Format:   &format1,
			},
		},
	))
	// Alignment/Border cells rules
	assert.NoError(t, f.SetConditionalFormat(sheet1, "M1:M10",
		[]ConditionalFormatOptions{
			{
				Type:     "cell",
				Criteria: ">",
				Format:   &format4,
				Value:    "0",
			},
		},
	))
	// Test set conditional format with invalid cell reference
	assert.Equal(t, newCellNameToCoordinatesError("-", newInvalidCellNameError("-")), f.SetConditionalFormat("Sheet1", "A1:-", nil))
	// Test set conditional format on not exists worksheet
	assert.EqualError(t, f.SetConditionalFormat("SheetN", "L1:L10", nil), "sheet SheetN does not exist")
	// Test set conditional format with invalid sheet name
	assert.Equal(t, ErrSheetNameInvalid, f.SetConditionalFormat("Sheet:1", "L1:L10", nil))

	err = f.SaveAs(filepath.Join("test", "TestConditionalFormat.xlsx"))
	assert.NoError(t, err)

	// Set conditional format with illegal valid type
	assert.Equal(t, ErrParameterInvalid, f.SetConditionalFormat(sheet1, "K1:K10",
		[]ConditionalFormatOptions{
			{
				Type:     "",
				Criteria: "=",
				MinType:  "min",
				MaxType:  "max",
				BarColor: "#638EC6",
			},
		},
	))
	// Set conditional format with illegal criteria type
	assert.Equal(t, ErrParameterInvalid, f.SetConditionalFormat(sheet1, "K1:K10",
		[]ConditionalFormatOptions{
			{
				Type:     "data_bar",
				Criteria: "",
				MinType:  "min",
				MaxType:  "max",
				BarColor: "#638EC6",
			},
		},
	))
	// Test create conditional format with invalid custom number format
	var exp string
	_, err = f.NewConditionalStyle(&Style{CustomNumFmt: &exp})
	assert.Equal(t, ErrCustomNumFmt, err)

	// Set conditional format with file without dxfs element should not return error
	f, err = OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)

	_, err = f.NewConditionalStyle(&Style{Font: &Font{Color: "9A0511"}, Fill: Fill{Type: "", Color: []string{"FEC7CE"}, Pattern: 1}})
	assert.NoError(t, err)
	assert.NoError(t, f.Close())
}

func TestSharedStrings(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "SharedStrings.xlsx"))
	assert.NoError(t, err)
	rows, err := f.GetRows("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, "A", rows[0][0])
	rows, err = f.GetRows("Sheet2")
	assert.NoError(t, err)
	assert.Equal(t, "Test Weight (Kgs)", rows[0][0])
	assert.NoError(t, f.Close())
}

func TestSetSheetCol(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)

	assert.NoError(t, f.SetSheetCol("Sheet1", "B27", &[]interface{}{"cell", nil, int32(42), float64(42), time.Now().UTC()}))

	assert.EqualError(t, f.SetSheetCol("Sheet1", "", &[]interface{}{"cell", nil, 2}),
		newCellNameToCoordinatesError("", newInvalidCellNameError("")).Error())
	// Test set worksheet column values with invalid sheet name
	assert.EqualError(t, f.SetSheetCol("Sheet:1", "A1", &[]interface{}{nil}), ErrSheetNameInvalid.Error())
	assert.EqualError(t, f.SetSheetCol("Sheet1", "B27", []interface{}{}), ErrParameterInvalid.Error())
	assert.EqualError(t, f.SetSheetCol("Sheet1", "B27", &f), ErrParameterInvalid.Error())
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetSheetCol.xlsx")))
	assert.NoError(t, f.Close())
}

func TestSetSheetRow(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)

	assert.NoError(t, f.SetSheetRow("Sheet1", "B27", &[]interface{}{"cell", nil, int32(42), float64(42), time.Now().UTC()}))

	assert.EqualError(t, f.SetSheetRow("Sheet1", "", &[]interface{}{"cell", nil, 2}),
		newCellNameToCoordinatesError("", newInvalidCellNameError("")).Error())
	// Test set worksheet row with invalid sheet name
	assert.EqualError(t, f.SetSheetRow("Sheet:1", "A1", &[]interface{}{1}), ErrSheetNameInvalid.Error())
	assert.EqualError(t, f.SetSheetRow("Sheet1", "B27", []interface{}{}), ErrParameterInvalid.Error())
	assert.EqualError(t, f.SetSheetRow("Sheet1", "B27", &f), ErrParameterInvalid.Error())
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetSheetRow.xlsx")))
	assert.NoError(t, f.Close())
}

func TestHSL(t *testing.T) {
	var hsl HSL
	r, g, b, a := hsl.RGBA()
	assert.Equal(t, uint32(0), r)
	assert.Equal(t, uint32(0), g)
	assert.Equal(t, uint32(0), b)
	assert.Equal(t, uint32(0xffff), a)
	assert.Equal(t, HSL{0, 0, 0}, hslModel(hsl))
	assert.Equal(t, HSL{0, 0, 0}, hslModel(color.Gray16{Y: uint16(1)}))
	R, G, B := HSLToRGB(0, 1, 0.4)
	assert.Equal(t, uint8(204), R)
	assert.Equal(t, uint8(0), G)
	assert.Equal(t, uint8(0), B)
	R, G, B = HSLToRGB(0, 1, 0.6)
	assert.Equal(t, uint8(255), R)
	assert.Equal(t, uint8(51), G)
	assert.Equal(t, uint8(51), B)
	assert.Equal(t, 0.0, hueToRGB(0, 0, -1))
	assert.Equal(t, 0.0, hueToRGB(0, 0, 2))
	assert.Equal(t, 0.0, hueToRGB(0, 0, 1.0/7))
	assert.Equal(t, 0.0, hueToRGB(0, 0, 0.4))
	assert.Equal(t, 0.0, hueToRGB(0, 0, 2.0/4))
	h, s, l := RGBToHSL(255, 255, 0)
	assert.Equal(t, 0.16666666666666666, h)
	assert.Equal(t, 1.0, s)
	assert.Equal(t, 0.5, l)
	h, s, l = RGBToHSL(0, 255, 255)
	assert.Equal(t, 0.5, h)
	assert.Equal(t, 1.0, s)
	assert.Equal(t, 0.5, l)
	h, s, l = RGBToHSL(250, 100, 50)
	assert.Equal(t, 0.041666666666666664, h)
	assert.Equal(t, 0.9523809523809524, s)
	assert.Equal(t, 0.5882352941176471, l)
	h, s, l = RGBToHSL(50, 100, 250)
	assert.Equal(t, 0.625, h)
	assert.Equal(t, 0.9523809523809524, s)
	assert.Equal(t, 0.5882352941176471, l)
	h, s, l = RGBToHSL(250, 50, 100)
	assert.Equal(t, 0.9583333333333334, h)
	assert.Equal(t, 0.9523809523809524, s)
	assert.Equal(t, 0.5882352941176471, l)
}

func TestProtectSheet(t *testing.T) {
	f := NewFile()
	sheetName := f.GetSheetName(0)
	assert.EqualError(t, f.ProtectSheet(sheetName, nil), ErrParameterInvalid.Error())
	// Test protect worksheet with XOR hash algorithm
	assert.NoError(t, f.ProtectSheet(sheetName, &SheetProtectionOptions{
		Password:      "password",
		EditScenarios: false,
	}))
	ws, err := f.workSheetReader(sheetName)
	assert.NoError(t, err)
	assert.Equal(t, "83AF", ws.SheetProtection.Password)
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestProtectSheet.xlsx")))
	// Test protect worksheet with SHA-512 hash algorithm
	assert.NoError(t, f.ProtectSheet(sheetName, &SheetProtectionOptions{
		AlgorithmName: "SHA-512",
		Password:      "password",
	}))
	ws, err = f.workSheetReader(sheetName)
	assert.NoError(t, err)
	assert.Len(t, ws.SheetProtection.SaltValue, 24)
	assert.Len(t, ws.SheetProtection.HashValue, 88)
	assert.Equal(t, int(sheetProtectionSpinCount), ws.SheetProtection.SpinCount)
	// Test remove sheet protection with an incorrect password
	assert.EqualError(t, f.UnprotectSheet(sheetName, "wrongPassword"), ErrUnprotectSheetPassword.Error())
	// Test remove sheet protection with invalid sheet name
	assert.EqualError(t, f.UnprotectSheet("Sheet:1", "wrongPassword"), ErrSheetNameInvalid.Error())
	// Test remove sheet protection with password verification
	assert.NoError(t, f.UnprotectSheet(sheetName, "password"))
	// Test protect worksheet with empty password
	assert.NoError(t, f.ProtectSheet(sheetName, &SheetProtectionOptions{}))
	assert.Equal(t, "", ws.SheetProtection.Password)
	// Test protect worksheet with password exceeds the limit length
	assert.EqualError(t, f.ProtectSheet(sheetName, &SheetProtectionOptions{
		AlgorithmName: "MD4",
		Password:      strings.Repeat("s", MaxFieldLength+1),
	}), ErrPasswordLengthInvalid.Error())
	// Test protect worksheet with unsupported hash algorithm
	assert.EqualError(t, f.ProtectSheet(sheetName, &SheetProtectionOptions{
		AlgorithmName: "RIPEMD-160",
		Password:      "password",
	}), ErrUnsupportedHashAlgorithm.Error())
	// Test protect not exists worksheet
	assert.EqualError(t, f.ProtectSheet("SheetN", nil), "sheet SheetN does not exist")
	// Test protect sheet with invalid sheet name
	assert.EqualError(t, f.ProtectSheet("Sheet:1", nil), ErrSheetNameInvalid.Error())
}

func TestUnprotectSheet(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)
	// Test remove protection on not exists worksheet
	assert.EqualError(t, f.UnprotectSheet("SheetN"), "sheet SheetN does not exist")

	assert.NoError(t, f.UnprotectSheet("Sheet1"))
	assert.EqualError(t, f.UnprotectSheet("Sheet1", "password"), ErrUnprotectSheet.Error())
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestUnprotectSheet.xlsx")))
	assert.NoError(t, f.Close())

	f = NewFile()
	sheetName := f.GetSheetName(0)
	assert.NoError(t, f.ProtectSheet(sheetName, &SheetProtectionOptions{Password: "password"}))
	// Test remove sheet protection with an incorrect password
	assert.EqualError(t, f.UnprotectSheet(sheetName, "wrongPassword"), ErrUnprotectSheetPassword.Error())
	// Test remove sheet protection with password verification
	assert.NoError(t, f.UnprotectSheet(sheetName, "password"))
	// Test with invalid salt value
	assert.NoError(t, f.ProtectSheet(sheetName, &SheetProtectionOptions{
		AlgorithmName: "SHA-512",
		Password:      "password",
	}))
	ws, err := f.workSheetReader(sheetName)
	assert.NoError(t, err)
	ws.SheetProtection.SaltValue = "YWJjZA====="
	assert.EqualError(t, f.UnprotectSheet(sheetName, "wrongPassword"), "illegal base64 data at input byte 8")
}

func TestProtectWorkbook(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.ProtectWorkbook(nil))
	// Test protect workbook with default hash algorithm
	assert.NoError(t, f.ProtectWorkbook(&WorkbookProtectionOptions{
		Password:      "password",
		LockStructure: true,
	}))
	wb, err := f.workbookReader()
	assert.NoError(t, err)
	assert.Equal(t, "SHA-512", wb.WorkbookProtection.WorkbookAlgorithmName)
	assert.Len(t, wb.WorkbookProtection.WorkbookSaltValue, 24)
	assert.Len(t, wb.WorkbookProtection.WorkbookHashValue, 88)
	assert.Equal(t, int(workbookProtectionSpinCount), wb.WorkbookProtection.WorkbookSpinCount)

	// Test protect workbook with password exceeds the limit length
	assert.EqualError(t, f.ProtectWorkbook(&WorkbookProtectionOptions{
		AlgorithmName: "MD4",
		Password:      strings.Repeat("s", MaxFieldLength+1),
	}), ErrPasswordLengthInvalid.Error())
	// Test protect workbook with unsupported hash algorithm
	assert.EqualError(t, f.ProtectWorkbook(&WorkbookProtectionOptions{
		AlgorithmName: "RIPEMD-160",
		Password:      "password",
	}), ErrUnsupportedHashAlgorithm.Error())
	// Test protect workbook with unsupported charset workbook
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	assert.EqualError(t, f.ProtectWorkbook(nil), "XML syntax error on line 1: invalid UTF-8")
}

func TestUnprotectWorkbook(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)

	assert.NoError(t, f.UnprotectWorkbook())
	assert.EqualError(t, f.UnprotectWorkbook("password"), ErrUnprotectWorkbook.Error())
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestUnprotectWorkbook.xlsx")))
	assert.NoError(t, f.Close())

	f = NewFile()
	assert.NoError(t, f.ProtectWorkbook(&WorkbookProtectionOptions{Password: "password"}))
	// Test remove workbook protection with an incorrect password
	assert.EqualError(t, f.UnprotectWorkbook("wrongPassword"), ErrUnprotectWorkbookPassword.Error())
	// Test remove workbook protection with password verification
	assert.NoError(t, f.UnprotectWorkbook("password"))
	// Test with invalid salt value
	assert.NoError(t, f.ProtectWorkbook(&WorkbookProtectionOptions{
		AlgorithmName: "SHA-512",
		Password:      "password",
	}))
	wb, err := f.workbookReader()
	assert.NoError(t, err)
	wb.WorkbookProtection.WorkbookSaltValue = "YWJjZA====="
	assert.EqualError(t, f.UnprotectWorkbook("wrongPassword"), "illegal base64 data at input byte 8")
	// Test remove workbook protection with unsupported charset workbook
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	assert.EqualError(t, f.UnprotectWorkbook(), "XML syntax error on line 1: invalid UTF-8")
}

func TestSetDefaultTimeStyle(t *testing.T) {
	f := NewFile()
	// Test set default time style on not exists worksheet.
	assert.EqualError(t, f.setDefaultTimeStyle("SheetN", "", 0), "sheet SheetN does not exist")

	// Test set default time style on invalid cell
	assert.EqualError(t, f.setDefaultTimeStyle("Sheet1", "", 42), newCellNameToCoordinatesError("", newInvalidCellNameError("")).Error())
}

func TestAddVBAProject(t *testing.T) {
	f := NewFile()
	file, err := os.ReadFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)
	assert.NoError(t, f.SetSheetProps("Sheet1", &SheetPropsOptions{CodeName: stringPtr("Sheet1")}))
	assert.EqualError(t, f.AddVBAProject(file), ErrAddVBAProject.Error())
	file, err = os.ReadFile(filepath.Join("test", "vbaProject.bin"))
	assert.NoError(t, err)
	assert.NoError(t, f.AddVBAProject(file))
	// Test add VBA project twice
	assert.NoError(t, f.AddVBAProject(file))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddVBAProject.xlsm")))
	// Test add VBA with unsupported charset workbook relationships
	f.Relationships.Delete(defaultXMLPathWorkbookRels)
	f.Pkg.Store(defaultXMLPathWorkbookRels, MacintoshCyrillicCharset)
	assert.EqualError(t, f.AddVBAProject(file), "XML syntax error on line 1: invalid UTF-8")
}

func TestContentTypesReader(t *testing.T) {
	// Test unsupported charset
	f := NewFile()
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	_, err := f.contentTypesReader()
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
}

func TestWorkbookReader(t *testing.T) {
	// Test unsupported charset
	f := NewFile()
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	_, err := f.workbookReader()
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
}

func TestWorkSheetReader(t *testing.T) {
	// Test unsupported charset
	f := NewFile()
	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", MacintoshCyrillicCharset)
	_, err := f.workSheetReader("Sheet1")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	assert.EqualError(t, f.UpdateLinkedValue(), "XML syntax error on line 1: invalid UTF-8")

	// Test on no checked worksheet
	f = NewFile()
	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`))
	f.checked = sync.Map{}
	_, err = f.workSheetReader("Sheet1")
	assert.NoError(t, err)
}

func TestRelsReader(t *testing.T) {
	// Test unsupported charset
	f := NewFile()
	rels := defaultXMLPathWorkbookRels
	f.Relationships.Store(rels, nil)
	f.Pkg.Store(rels, MacintoshCyrillicCharset)
	_, err := f.relsReader(rels)
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
}

func TestDeleteSheetFromWorkbookRels(t *testing.T) {
	f := NewFile()
	rels := defaultXMLPathWorkbookRels
	f.Relationships.Store(rels, nil)
	assert.Equal(t, f.deleteSheetFromWorkbookRels("rID"), "")
}

func TestUpdateLinkedValue(t *testing.T) {
	f := NewFile()
	// Test update lined value with unsupported charset workbook
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	assert.EqualError(t, f.UpdateLinkedValue(), "XML syntax error on line 1: invalid UTF-8")
}

func TestAttrValToInt(t *testing.T) {
	_, err := attrValToInt("r", []xml.Attr{
		{Name: xml.Name{Local: "r"}, Value: "s"},
	})
	assert.EqualError(t, err, `strconv.Atoi: parsing "s": invalid syntax`)
}

func prepareTestBook1() (*File, error) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if err != nil {
		return nil, err
	}

	if err = f.AddPicture("Sheet2", "I9", filepath.Join("test", "images", "excel.jpg"),
		&GraphicOptions{OffsetX: 140, OffsetY: 120, Hyperlink: "#Sheet2!D8", HyperlinkType: "Location"}); err != nil {
		return nil, err
	}

	// Test add picture to worksheet with offset, external hyperlink and positioning
	if err := f.AddPicture("Sheet1", "F21", filepath.Join("test", "images", "excel.png"),
		&GraphicOptions{
			OffsetX:       10,
			OffsetY:       10,
			Hyperlink:     "https://github.com/xuri/excelize",
			HyperlinkType: "External",
			Positioning:   "oneCell",
		},
	); err != nil {
		return nil, err
	}

	file, err := os.ReadFile(filepath.Join("test", "images", "excel.jpg"))
	if err != nil {
		return nil, err
	}

	err = f.AddPictureFromBytes("Sheet1", "Q1", &Picture{Extension: ".jpg", File: file, Format: &GraphicOptions{AltText: "Excel Logo"}})
	if err != nil {
		return nil, err
	}

	return f, nil
}

func prepareTestBook3() (*File, error) {
	f := NewFile()
	if _, err := f.NewSheet("Sheet2"); err != nil {
		return nil, err
	}
	if _, err := f.NewSheet("Sheet3"); err != nil {
		return nil, err
	}
	if err := f.SetCellInt("Sheet2", "A23", 56); err != nil {
		return nil, err
	}
	if err := f.SetCellStr("Sheet1", "B20", "42"); err != nil {
		return nil, err
	}
	f.SetActiveSheet(0)
	if err := f.AddPicture("Sheet1", "H2", filepath.Join("test", "images", "excel.gif"),
		&GraphicOptions{ScaleX: 0.5, ScaleY: 0.5, Positioning: "absolute"}); err != nil {
		return nil, err
	}
	if err := f.AddPicture("Sheet1", "C2", filepath.Join("test", "images", "excel.png"), nil); err != nil {
		return nil, err
	}
	return f, nil
}

func prepareTestBook4() (*File, error) {
	f := NewFile()
	if err := f.SetColWidth("Sheet1", "B", "A", 12); err != nil {
		return f, err
	}
	if err := f.SetColWidth("Sheet1", "A", "B", 12); err != nil {
		return f, err
	}
	if _, err := f.GetColWidth("Sheet1", "A"); err != nil {
		return f, err
	}
	if _, err := f.GetColWidth("Sheet1", "C"); err != nil {
		return f, err
	}

	return f, nil
}

func prepareTestBook5(opts Options) (*File, error) {
	f := NewFile(opts)
	var rowNum int
	for _, idxRange := range [][]int{{27, 36}, {50, 81}} {
		for numFmtIdx := idxRange[0]; numFmtIdx <= idxRange[1]; numFmtIdx++ {
			rowNum++
			styleID, err := f.NewStyle(&Style{NumFmt: numFmtIdx})
			if err != nil {
				return f, err
			}
			cell, err := CoordinatesToCellName(1, rowNum)
			if err != nil {
				return f, err
			}
			if err := f.SetCellValue("Sheet1", cell, 45162); err != nil {
				return f, err
			}
			if err := f.SetCellStyle("Sheet1", cell, cell, styleID); err != nil {
				return f, err
			}
		}
	}
	return f, nil
}

func fillCells(f *File, sheet string, colCount, rowCount int) error {
	for col := 1; col <= colCount; col++ {
		for row := 1; row <= rowCount; row++ {
			cell, _ := CoordinatesToCellName(col, row)
			if err := f.SetCellStr(sheet, cell, cell); err != nil {
				return err
			}
		}
	}
	return nil
}

func BenchmarkOpenFile(b *testing.B) {
	for i := 0; i < b.N; i++ {
		f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
		if err != nil {
			b.Error(err)
		}
		if err := f.Close(); err != nil {
			b.Error(err)
		}
	}
}
