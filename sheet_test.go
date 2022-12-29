package excelize

import (
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestNewSheet(t *testing.T) {
	f := NewFile()
	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)
	sheetID, err := f.NewSheet("sheet2")
	assert.NoError(t, err)
	f.SetActiveSheet(sheetID)
	// Test delete original sheet
	idx, err := f.GetSheetIndex("Sheet1")
	assert.NoError(t, err)
	assert.NoError(t, f.DeleteSheet(f.GetSheetName(idx)))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestNewSheet.xlsx")))
	// Test create new worksheet with already exists name
	sheetID, err = f.NewSheet("Sheet2")
	assert.NoError(t, err)
	idx, err = f.GetSheetIndex("Sheet2")
	assert.NoError(t, err)
	assert.Equal(t, idx, sheetID)
	// Test create new worksheet with empty sheet name
	sheetID, err = f.NewSheet(":\\/?*[]")
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())
	assert.Equal(t, -1, sheetID)
}

func TestSetPanes(t *testing.T) {
	f := NewFile()

	assert.NoError(t, f.SetPanes("Sheet1", &Panes{Freeze: false, Split: false}))
	_, err := f.NewSheet("Panes 2")
	assert.NoError(t, err)
	assert.NoError(t, f.SetPanes("Panes 2",
		&Panes{
			Freeze:      true,
			Split:       false,
			XSplit:      1,
			YSplit:      0,
			TopLeftCell: "B1",
			ActivePane:  "topRight",
			Panes: []PaneOptions{
				{SQRef: "K16", ActiveCell: "K16", Pane: "topRight"},
			},
		},
	))
	_, err = f.NewSheet("Panes 3")
	assert.NoError(t, err)
	assert.NoError(t, f.SetPanes("Panes 3",
		&Panes{
			Freeze:      false,
			Split:       true,
			XSplit:      3270,
			YSplit:      1800,
			TopLeftCell: "N57",
			ActivePane:  "bottomLeft",
			Panes: []PaneOptions{
				{SQRef: "I36", ActiveCell: "I36"},
				{SQRef: "G33", ActiveCell: "G33", Pane: "topRight"},
				{SQRef: "J60", ActiveCell: "J60", Pane: "bottomLeft"},
				{SQRef: "O60", ActiveCell: "O60", Pane: "bottomRight"},
			},
		},
	))
	_, err = f.NewSheet("Panes 4")
	assert.NoError(t, err)
	assert.NoError(t, f.SetPanes("Panes 4",
		&Panes{
			Freeze:      true,
			Split:       false,
			XSplit:      0,
			YSplit:      9,
			TopLeftCell: "A34",
			ActivePane:  "bottomLeft",
			Panes: []PaneOptions{
				{SQRef: "A11:XFD11", ActiveCell: "A11", Pane: "bottomLeft"},
			},
		},
	))
	assert.EqualError(t, f.SetPanes("Panes 4", nil), ErrParameterInvalid.Error())
	assert.EqualError(t, f.SetPanes("SheetN", nil), "sheet SheetN does not exist")
	// Test set panes with invalid sheet name
	assert.EqualError(t, f.SetPanes("Sheet:1", &Panes{Freeze: false, Split: false}), ErrSheetNameInvalid.Error())
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetPane.xlsx")))
	// Test add pane on empty sheet views worksheet
	f = NewFile()
	f.checked = nil
	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`))
	assert.NoError(t, f.SetPanes("Sheet1",
		&Panes{
			Freeze:      true,
			Split:       false,
			XSplit:      1,
			YSplit:      0,
			TopLeftCell: "B1",
			ActivePane:  "topRight",
			Panes: []PaneOptions{
				{SQRef: "K16", ActiveCell: "K16", Pane: "topRight"},
			},
		},
	))
}

func TestSearchSheet(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "SharedStrings.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	// Test search in a not exists worksheet
	_, err = f.SearchSheet("Sheet4", "")
	assert.EqualError(t, err, "sheet Sheet4 does not exist")
	// Test search sheet with invalid sheet name
	_, err = f.SearchSheet("Sheet:1", "")
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())
	var expected []string
	// Test search a not exists value
	result, err := f.SearchSheet("Sheet1", "X")
	assert.NoError(t, err)
	assert.EqualValues(t, expected, result)
	result, err = f.SearchSheet("Sheet1", "A")
	assert.NoError(t, err)
	assert.EqualValues(t, []string{"A1"}, result)
	// Test search the coordinates where the numerical value in the range of
	// "0-9" of Sheet1 is described by regular expression:
	result, err = f.SearchSheet("Sheet1", "[0-9]", true)
	assert.NoError(t, err)
	assert.EqualValues(t, expected, result)
	assert.NoError(t, f.Close())

	// Test search worksheet data after set cell value
	f = NewFile()
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", true))
	_, err = f.SearchSheet("Sheet1", "")
	assert.NoError(t, err)

	f = NewFile()
	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(`<worksheet><sheetData><row r="A"><c r="2" t="inlineStr"><is><t>A</t></is></c></row></sheetData></worksheet>`))
	f.checked = nil
	result, err = f.SearchSheet("Sheet1", "A")
	assert.EqualError(t, err, "strconv.Atoi: parsing \"A\": invalid syntax")
	assert.Equal(t, []string(nil), result)

	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(`<worksheet><sheetData><row r="2"><c r="A" t="inlineStr"><is><t>A</t></is></c></row></sheetData></worksheet>`))
	result, err = f.SearchSheet("Sheet1", "A")
	assert.EqualError(t, err, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	assert.Equal(t, []string(nil), result)

	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(`<worksheet><sheetData><row r="0"><c r="A1" t="inlineStr"><is><t>A</t></is></c></row></sheetData></worksheet>`))
	result, err = f.SearchSheet("Sheet1", "A")
	assert.EqualError(t, err, "invalid cell reference [1, 0]")
	assert.Equal(t, []string(nil), result)

	// Test search sheet with unsupported charset shared strings table
	f.SharedStrings = nil
	f.Pkg.Store(defaultXMLPathSharedStrings, MacintoshCyrillicCharset)
	_, err = f.SearchSheet("Sheet1", "A")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
}

func TestSetPageLayout(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetPageLayout("Sheet1", nil))
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).PageSetUp = nil
	expected := PageLayoutOptions{
		Size:            intPtr(1),
		Orientation:     stringPtr("landscape"),
		FirstPageNumber: uintPtr(1),
		AdjustTo:        uintPtr(120),
		FitToHeight:     intPtr(2),
		FitToWidth:      intPtr(2),
		BlackAndWhite:   boolPtr(true),
	}
	assert.NoError(t, f.SetPageLayout("Sheet1", &expected))
	opts, err := f.GetPageLayout("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, expected, opts)
	// Test set page layout on not exists worksheet
	assert.EqualError(t, f.SetPageLayout("SheetN", nil), "sheet SheetN does not exist")
	// Test set page layout with invalid sheet name
	assert.EqualError(t, f.SetPageLayout("Sheet:1", nil), ErrSheetNameInvalid.Error())
}

func TestGetPageLayout(t *testing.T) {
	f := NewFile()
	// Test get page layout on not exists worksheet
	_, err := f.GetPageLayout("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Test get page layout with invalid sheet name
	_, err = f.GetPageLayout("Sheet:1")
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())
}

func TestSetHeaderFooter(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellStr("Sheet1", "A1", "Test SetHeaderFooter"))
	// Test set header and footer on not exists worksheet
	assert.EqualError(t, f.SetHeaderFooter("SheetN", nil), "sheet SheetN does not exist")
	// Test Sheet:1 with invalid sheet name
	assert.EqualError(t, f.SetHeaderFooter("Sheet:1", nil), ErrSheetNameInvalid.Error())
	// Test set header and footer with illegal setting
	assert.EqualError(t, f.SetHeaderFooter("Sheet1", &HeaderFooterOptions{
		OddHeader: strings.Repeat("c", MaxFieldLength+1),
	}), newFieldLengthError("OddHeader").Error())

	assert.NoError(t, f.SetHeaderFooter("Sheet1", nil))
	text := strings.Repeat("一", MaxFieldLength)
	assert.NoError(t, f.SetHeaderFooter("Sheet1", &HeaderFooterOptions{
		OddHeader:   text,
		OddFooter:   text,
		EvenHeader:  text,
		EvenFooter:  text,
		FirstHeader: text,
	}))
	assert.NoError(t, f.SetHeaderFooter("Sheet1", &HeaderFooterOptions{
		DifferentFirst:   true,
		DifferentOddEven: true,
		OddHeader:        "&R&P",
		OddFooter:        "&C&F",
		EvenHeader:       "&L&P",
		EvenFooter:       "&L&D&R&T",
		FirstHeader:      `&CCenter &"-,Bold"Bold&"-,Regular"HeaderU+000A&D`,
	}))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetHeaderFooter.xlsx")))
}

func TestDefinedName(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetDefinedName(&DefinedName{
		Name:     "Amount",
		RefersTo: "Sheet1!$A$2:$D$5",
		Comment:  "defined name comment",
		Scope:    "Sheet1",
	}))
	assert.NoError(t, f.SetDefinedName(&DefinedName{
		Name:     "Amount",
		RefersTo: "Sheet1!$A$2:$D$5",
		Comment:  "defined name comment",
	}))
	assert.EqualError(t, f.SetDefinedName(&DefinedName{
		Name:     "Amount",
		RefersTo: "Sheet1!$A$2:$D$5",
		Comment:  "defined name comment",
	}), ErrDefinedNameDuplicate.Error())
	assert.EqualError(t, f.DeleteDefinedName(&DefinedName{
		Name: "No Exist Defined Name",
	}), ErrDefinedNameScope.Error())
	// Test set defined name without name
	assert.EqualError(t, f.SetDefinedName(&DefinedName{
		RefersTo: "Sheet1!$A$2:$D$5",
	}), ErrParameterInvalid.Error())
	// Test set defined name without reference
	assert.EqualError(t, f.SetDefinedName(&DefinedName{
		Name: "Amount",
	}), ErrParameterInvalid.Error())
	assert.Exactly(t, "Sheet1!$A$2:$D$5", f.GetDefinedName()[1].RefersTo)
	assert.NoError(t, f.DeleteDefinedName(&DefinedName{
		Name: "Amount",
	}))
	assert.Exactly(t, "Sheet1!$A$2:$D$5", f.GetDefinedName()[0].RefersTo)
	assert.Exactly(t, 1, len(f.GetDefinedName()))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestDefinedName.xlsx")))
	// Test set defined name with unsupported charset workbook
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	assert.EqualError(t, f.SetDefinedName(&DefinedName{
		Name: "Amount", RefersTo: "Sheet1!$A$2:$D$5",
	}), "XML syntax error on line 1: invalid UTF-8")
	// Test delete defined name with unsupported charset workbook
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	assert.EqualError(t, f.DeleteDefinedName(&DefinedName{Name: "Amount"}),
		"XML syntax error on line 1: invalid UTF-8")
}

func TestGroupSheets(t *testing.T) {
	f := NewFile()
	sheets := []string{"Sheet2", "Sheet3"}
	for _, sheet := range sheets {
		_, err := f.NewSheet(sheet)
		assert.NoError(t, err)
	}
	assert.EqualError(t, f.GroupSheets([]string{"Sheet1", "SheetN"}), "sheet SheetN does not exist")
	assert.EqualError(t, f.GroupSheets([]string{"Sheet2", "Sheet3"}), "group worksheet must contain an active worksheet")
	// Test group sheets with invalid sheet name
	assert.EqualError(t, f.GroupSheets([]string{"Sheet:1", "Sheet1"}), ErrSheetNameInvalid.Error())
	assert.NoError(t, f.GroupSheets([]string{"Sheet1", "Sheet2"}))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestGroupSheets.xlsx")))
}

func TestUngroupSheets(t *testing.T) {
	f := NewFile()
	sheets := []string{"Sheet2", "Sheet3", "Sheet4", "Sheet5"}
	for _, sheet := range sheets {
		_, err := f.NewSheet(sheet)
		assert.NoError(t, err)
	}
	assert.NoError(t, f.UngroupSheets())
}

func TestInsertPageBreak(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.InsertPageBreak("Sheet1", "A1"))
	assert.NoError(t, f.InsertPageBreak("Sheet1", "B2"))
	assert.NoError(t, f.InsertPageBreak("Sheet1", "C3"))
	assert.NoError(t, f.InsertPageBreak("Sheet1", "C3"))
	assert.EqualError(t, f.InsertPageBreak("Sheet1", "A"), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	assert.EqualError(t, f.InsertPageBreak("SheetN", "C3"), "sheet SheetN does not exist")
	// Test insert page break with invalid sheet name
	assert.EqualError(t, f.InsertPageBreak("Sheet:1", "C3"), ErrSheetNameInvalid.Error())
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestInsertPageBreak.xlsx")))
}

func TestRemovePageBreak(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.RemovePageBreak("Sheet1", "A2"))

	assert.NoError(t, f.InsertPageBreak("Sheet1", "A2"))
	assert.NoError(t, f.InsertPageBreak("Sheet1", "B2"))
	assert.NoError(t, f.RemovePageBreak("Sheet1", "A1"))
	assert.NoError(t, f.RemovePageBreak("Sheet1", "B2"))

	assert.NoError(t, f.InsertPageBreak("Sheet1", "C3"))
	assert.NoError(t, f.RemovePageBreak("Sheet1", "C3"))

	assert.NoError(t, f.InsertPageBreak("Sheet1", "A3"))
	assert.NoError(t, f.RemovePageBreak("Sheet1", "B3"))
	assert.NoError(t, f.RemovePageBreak("Sheet1", "A3"))

	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)
	assert.NoError(t, f.InsertPageBreak("Sheet2", "B2"))
	assert.NoError(t, f.InsertPageBreak("Sheet2", "C2"))
	assert.NoError(t, f.RemovePageBreak("Sheet2", "B2"))

	assert.EqualError(t, f.RemovePageBreak("Sheet1", "A"), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	assert.EqualError(t, f.RemovePageBreak("SheetN", "C3"), "sheet SheetN does not exist")
	// Test remove page break with invalid sheet name
	assert.EqualError(t, f.RemovePageBreak("Sheet:1", "A3"), ErrSheetNameInvalid.Error())
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestRemovePageBreak.xlsx")))
}

func TestGetSheetName(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)
	assert.Equal(t, "Sheet1", f.GetSheetName(0))
	assert.Equal(t, "Sheet2", f.GetSheetName(1))
	assert.Equal(t, "", f.GetSheetName(-1))
	assert.Equal(t, "", f.GetSheetName(2))
	assert.NoError(t, f.Close())
}

func TestGetSheetMap(t *testing.T) {
	expectedMap := map[int]string{
		1: "Sheet1",
		2: "Sheet2",
	}
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)
	sheetMap := f.GetSheetMap()
	for idx, name := range sheetMap {
		assert.Equal(t, expectedMap[idx], name)
	}
	assert.Equal(t, len(sheetMap), 2)
	assert.NoError(t, f.Close())
}

func TestSetActiveSheet(t *testing.T) {
	f := NewFile()
	f.WorkBook.BookViews = nil
	f.SetActiveSheet(1)
	f.WorkBook.BookViews = &xlsxBookViews{WorkBookView: []xlsxWorkBookView{}}
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).SheetViews = &xlsxSheetViews{SheetView: []xlsxSheetView{}}
	f.SetActiveSheet(1)
	ws, ok = f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).SheetViews = nil
	f.SetActiveSheet(1)
	f = NewFile()
	f.SetActiveSheet(-1)
	assert.Equal(t, f.GetActiveSheetIndex(), 0)

	f = NewFile()
	f.WorkBook.BookViews = nil
	idx, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)
	ws, ok = f.Sheet.Load("xl/worksheets/sheet2.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).SheetViews = &xlsxSheetViews{SheetView: []xlsxSheetView{}}
	f.SetActiveSheet(idx)
}

func TestSetSheetName(t *testing.T) {
	f := NewFile()
	// Test set worksheet with the same name
	assert.NoError(t, f.SetSheetName("Sheet1", "Sheet1"))
	assert.Equal(t, "Sheet1", f.GetSheetName(0))
	// Test set sheet name with invalid sheet name
	assert.EqualError(t, f.SetSheetName("Sheet:1", "Sheet1"), ErrSheetNameInvalid.Error())
}

func TestWorksheetWriter(t *testing.T) {
	f := NewFile()
	// Test set cell value with alternate content
	f.Sheet.Delete("xl/worksheets/sheet1.xml")
	worksheet := xml.Header + `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1"><v>%d</v></c></row></sheetData><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14"><xdr:twoCellAnchor editAs="oneCell"></xdr:twoCellAnchor></mc:Choice><mc:Fallback/></mc:AlternateContent></worksheet>`
	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(fmt.Sprintf(worksheet, 1)))
	f.checked = nil
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 2))
	f.workSheetWriter()
	value, ok := f.Pkg.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	assert.Equal(t, fmt.Sprintf(worksheet, 2), string(value.([]byte)))
}

func TestGetWorkbookPath(t *testing.T) {
	f := NewFile()
	f.Pkg.Delete("_rels/.rels")
	assert.Equal(t, "", f.getWorkbookPath())
}

func TestGetWorkbookRelsPath(t *testing.T) {
	f := NewFile()
	f.Pkg.Delete("xl/_rels/.rels")
	f.Pkg.Store("_rels/.rels", []byte(xml.Header+`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="/workbook.xml"/></Relationships>`))
	assert.Equal(t, "_rels/workbook.xml.rels", f.getWorkbookRelsPath())
}

func TestDeleteSheet(t *testing.T) {
	f := NewFile()
	idx, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)
	f.SetActiveSheet(idx)
	_, err = f.NewSheet("Sheet3")
	assert.NoError(t, err)
	assert.NoError(t, f.DeleteSheet("Sheet1"))
	assert.Equal(t, "Sheet2", f.GetSheetName(f.GetActiveSheetIndex()))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestDeleteSheet.xlsx")))
	// Test with auto filter defined names
	f = NewFile()
	_, err = f.NewSheet("Sheet2")
	assert.NoError(t, err)
	_, err = f.NewSheet("Sheet3")
	assert.NoError(t, err)
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "A"))
	assert.NoError(t, f.SetCellValue("Sheet2", "A1", "A"))
	assert.NoError(t, f.SetCellValue("Sheet3", "A1", "A"))
	assert.NoError(t, f.AutoFilter("Sheet1", "A1:A1", nil))
	assert.NoError(t, f.AutoFilter("Sheet2", "A1:A1", nil))
	assert.NoError(t, f.AutoFilter("Sheet3", "A1:A1", nil))
	assert.NoError(t, f.DeleteSheet("Sheet2"))
	assert.NoError(t, f.DeleteSheet("Sheet1"))
	// Test delete sheet with invalid sheet name
	assert.EqualError(t, f.DeleteSheet("Sheet:1"), ErrSheetNameInvalid.Error())
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestDeleteSheet2.xlsx")))
}

func TestDeleteAndAdjustDefinedNames(t *testing.T) {
	deleteAndAdjustDefinedNames(nil, 0)
	deleteAndAdjustDefinedNames(&xlsxWorkbook{}, 0)
}

func TestGetSheetID(t *testing.T) {
	f := NewFile()
	_, err := f.NewSheet("Sheet1")
	assert.NoError(t, err)
	id := f.getSheetID("sheet1")
	assert.NotEqual(t, -1, id)
}

func TestSetSheetVisible(t *testing.T) {
	f := NewFile()
	// Test set sheet visible with invalid sheet name
	assert.EqualError(t, f.SetSheetVisible("Sheet:1", false), ErrSheetNameInvalid.Error())
	f.WorkBook.Sheets.Sheet[0].Name = "SheetN"
	assert.EqualError(t, f.SetSheetVisible("Sheet1", false), "sheet SheetN does not exist")
	// Test set sheet visible with unsupported charset workbook
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	assert.EqualError(t, f.SetSheetVisible("Sheet1", false), "XML syntax error on line 1: invalid UTF-8")
}

func TestGetSheetVisible(t *testing.T) {
	f := NewFile()
	// Test get sheet visible with invalid sheet name
	visible, err := f.GetSheetVisible("Sheet:1")
	assert.Equal(t, false, visible)
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())
}

func TestGetSheetIndex(t *testing.T) {
	f := NewFile()
	// Test get sheet index with invalid sheet name
	idx, err := f.GetSheetIndex("Sheet:1")
	assert.Equal(t, -1, idx)
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())
}

func TestSetContentTypes(t *testing.T) {
	f := NewFile()
	// Test set content type with unsupported charset content types
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, f.setContentTypes("/xl/worksheets/sheet1.xml", ContentTypeSpreadSheetMLWorksheet), "XML syntax error on line 1: invalid UTF-8")
}

func TestDeleteSheetFromContentTypes(t *testing.T) {
	f := NewFile()
	// Test delete sheet from content types with unsupported charset content types
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, f.deleteSheetFromContentTypes("/xl/worksheets/sheet1.xml"), "XML syntax error on line 1: invalid UTF-8")
}

func BenchmarkNewSheet(b *testing.B) {
	b.RunParallel(func(pb *testing.PB) {
		for pb.Next() {
			newSheetWithSet()
		}
	})
}

func newSheetWithSet() {
	file := NewFile()
	for i := 0; i < 1000; i++ {
		_ = file.SetCellInt("Sheet1", "A"+strconv.Itoa(i+1), i)
	}
	file = nil
}

func BenchmarkFile_SaveAs(b *testing.B) {
	b.RunParallel(func(pb *testing.PB) {
		for pb.Next() {
			newSheetWithSave()
		}
	})
}

func newSheetWithSave() {
	file := NewFile()
	for i := 0; i < 1000; i++ {
		_ = file.SetCellInt("Sheet1", "A"+strconv.Itoa(i+1), i)
	}
	_ = file.Save()
}

func TestAttrValToBool(t *testing.T) {
	_, err := attrValToBool("hidden", []xml.Attr{
		{Name: xml.Name{Local: "hidden"}},
	})
	assert.EqualError(t, err, `strconv.ParseBool: parsing "": invalid syntax`)

	got, err := attrValToBool("hidden", []xml.Attr{
		{Name: xml.Name{Local: "hidden"}, Value: "1"},
	})
	assert.NoError(t, err)
	assert.Equal(t, true, got)
}

func TestAttrValToFloat(t *testing.T) {
	_, err := attrValToFloat("ht", []xml.Attr{
		{Name: xml.Name{Local: "ht"}},
	})
	assert.EqualError(t, err, `strconv.ParseFloat: parsing "": invalid syntax`)

	got, err := attrValToFloat("ht", []xml.Attr{
		{Name: xml.Name{Local: "ht"}, Value: "42.1"},
	})
	assert.NoError(t, err)
	assert.Equal(t, 42.1, got)
}

func TestSetSheetBackgroundFromBytes(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetSheetName("Sheet1", ".svg"))
	for i, imageTypes := range []string{".svg", ".emf", ".emz", ".gif", ".jpg", ".png", ".tif", ".wmf", ".wmz"} {
		file := fmt.Sprintf("excelize%s", imageTypes)
		if i > 0 {
			file = filepath.Join("test", "images", fmt.Sprintf("excel%s", imageTypes))
			_, err := f.NewSheet(imageTypes)
			assert.NoError(t, err)
		}
		img, err := os.Open(file)
		assert.NoError(t, err)
		content, err := io.ReadAll(img)
		assert.NoError(t, err)
		assert.NoError(t, img.Close())
		assert.NoError(t, f.SetSheetBackgroundFromBytes(imageTypes, imageTypes, content))
	}
	// Test set worksheet background with invalid sheet name
	img, err := os.Open(filepath.Join("test", "images", "excel.png"))
	assert.NoError(t, err)
	content, err := io.ReadAll(img)
	assert.NoError(t, err)
	assert.EqualError(t, f.SetSheetBackgroundFromBytes("Sheet:1", ".png", content), ErrSheetNameInvalid.Error())

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetSheetBackgroundFromBytes.xlsx")))
	assert.NoError(t, f.Close())

	assert.EqualError(t, f.SetSheetBackgroundFromBytes("Sheet1", ".svg", nil), ErrParameterInvalid.Error())
}

func TestCheckSheetName(t *testing.T) {
	// Test valid sheet name
	assert.NoError(t, checkSheetName("Sheet1"))
	assert.NoError(t, checkSheetName("She'et1"))
	// Test invalid sheet name, empty name
	assert.EqualError(t, checkSheetName(""), ErrSheetNameBlank.Error())
	// Test invalid sheet name, include :\/?*[]
	assert.EqualError(t, checkSheetName("Sheet:"), ErrSheetNameInvalid.Error())
	assert.EqualError(t, checkSheetName(`Sheet\`), ErrSheetNameInvalid.Error())
	assert.EqualError(t, checkSheetName("Sheet/"), ErrSheetNameInvalid.Error())
	assert.EqualError(t, checkSheetName("Sheet?"), ErrSheetNameInvalid.Error())
	assert.EqualError(t, checkSheetName("Sheet*"), ErrSheetNameInvalid.Error())
	assert.EqualError(t, checkSheetName("Sheet["), ErrSheetNameInvalid.Error())
	assert.EqualError(t, checkSheetName("Sheet]"), ErrSheetNameInvalid.Error())
	// Test invalid sheet name, single quotes at the front or at the end
	assert.EqualError(t, checkSheetName("'Sheet"), ErrSheetNameSingleQuote.Error())
	assert.EqualError(t, checkSheetName("Sheet'"), ErrSheetNameSingleQuote.Error())
}
