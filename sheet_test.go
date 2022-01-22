package excelize

import (
	"fmt"
	"path/filepath"
	"strconv"
	"strings"
	"testing"

	"github.com/mohae/deepcopy"
	"github.com/stretchr/testify/assert"
)

func ExampleFile_SetPageLayout() {
	f := NewFile()
	if err := f.SetPageLayout(
		"Sheet1",
		BlackAndWhite(true),
		FirstPageNumber(2),
		PageLayoutOrientation(OrientationLandscape),
		PageLayoutPaperSize(10),
		FitToHeight(2),
		FitToWidth(2),
		PageLayoutScale(50),
	); err != nil {
		fmt.Println(err)
	}
	// Output:
}

func ExampleFile_GetPageLayout() {
	f := NewFile()
	var (
		blackAndWhite   BlackAndWhite
		firstPageNumber FirstPageNumber
		orientation     PageLayoutOrientation
		paperSize       PageLayoutPaperSize
		fitToHeight     FitToHeight
		fitToWidth      FitToWidth
		scale           PageLayoutScale
	)
	if err := f.GetPageLayout("Sheet1", &blackAndWhite); err != nil {
		fmt.Println(err)
	}
	if err := f.GetPageLayout("Sheet1", &firstPageNumber); err != nil {
		fmt.Println(err)
	}
	if err := f.GetPageLayout("Sheet1", &orientation); err != nil {
		fmt.Println(err)
	}
	if err := f.GetPageLayout("Sheet1", &paperSize); err != nil {
		fmt.Println(err)
	}
	if err := f.GetPageLayout("Sheet1", &fitToHeight); err != nil {
		fmt.Println(err)
	}
	if err := f.GetPageLayout("Sheet1", &fitToWidth); err != nil {
		fmt.Println(err)
	}
	if err := f.GetPageLayout("Sheet1", &scale); err != nil {
		fmt.Println(err)
	}
	fmt.Println("Defaults:")
	fmt.Printf("- print black and white: %t\n", blackAndWhite)
	fmt.Printf("- page number for first printed page: %d\n", firstPageNumber)
	fmt.Printf("- orientation: %q\n", orientation)
	fmt.Printf("- paper size: %d\n", paperSize)
	fmt.Printf("- fit to height: %d\n", fitToHeight)
	fmt.Printf("- fit to width: %d\n", fitToWidth)
	fmt.Printf("- scale: %d\n", scale)
	// Output:
	// Defaults:
	// - print black and white: false
	// - page number for first printed page: 1
	// - orientation: "portrait"
	// - paper size: 1
	// - fit to height: 1
	// - fit to width: 1
	// - scale: 100
}

func TestNewSheet(t *testing.T) {
	f := NewFile()
	f.NewSheet("Sheet2")
	sheetID := f.NewSheet("sheet2")
	f.SetActiveSheet(sheetID)
	// delete original sheet
	f.DeleteSheet(f.GetSheetName(f.GetSheetIndex("Sheet1")))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestNewSheet.xlsx")))
	// create new worksheet with already exists name
	assert.Equal(t, f.GetSheetIndex("Sheet2"), f.NewSheet("Sheet2"))
}

func TestSetPane(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetPanes("Sheet1", `{"freeze":false,"split":false}`))
	f.NewSheet("Panes 2")
	assert.NoError(t, f.SetPanes("Panes 2", `{"freeze":true,"split":false,"x_split":1,"y_split":0,"top_left_cell":"B1","active_pane":"topRight","panes":[{"sqref":"K16","active_cell":"K16","pane":"topRight"}]}`))
	f.NewSheet("Panes 3")
	assert.NoError(t, f.SetPanes("Panes 3", `{"freeze":false,"split":true,"x_split":3270,"y_split":1800,"top_left_cell":"N57","active_pane":"bottomLeft","panes":[{"sqref":"I36","active_cell":"I36"},{"sqref":"G33","active_cell":"G33","pane":"topRight"},{"sqref":"J60","active_cell":"J60","pane":"bottomLeft"},{"sqref":"O60","active_cell":"O60","pane":"bottomRight"}]}`))
	f.NewSheet("Panes 4")
	assert.NoError(t, f.SetPanes("Panes 4", `{"freeze":true,"split":false,"x_split":0,"y_split":9,"top_left_cell":"A34","active_pane":"bottomLeft","panes":[{"sqref":"A11:XFD11","active_cell":"A11","pane":"bottomLeft"}]}`))
	assert.NoError(t, f.SetPanes("Panes 4", ""))
	assert.EqualError(t, f.SetPanes("SheetN", ""), "sheet SheetN is not exist")
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestSetPane.xlsx")))
}

func TestPageLayoutOption(t *testing.T) {
	const sheet = "Sheet1"

	testData := []struct {
		container  PageLayoutOptionPtr
		nonDefault PageLayoutOption
	}{
		{new(BlackAndWhite), BlackAndWhite(true)},
		{new(FirstPageNumber), FirstPageNumber(2)},
		{new(PageLayoutOrientation), PageLayoutOrientation(OrientationLandscape)},
		{new(PageLayoutPaperSize), PageLayoutPaperSize(10)},
		{new(FitToHeight), FitToHeight(2)},
		{new(FitToWidth), FitToWidth(2)},
		{new(PageLayoutScale), PageLayoutScale(50)},
	}

	for i, test := range testData {
		t.Run(fmt.Sprintf("TestData%d", i), func(t *testing.T) {
			opt := test.nonDefault
			t.Logf("option %T", opt)

			def := deepcopy.Copy(test.container).(PageLayoutOptionPtr)
			val1 := deepcopy.Copy(def).(PageLayoutOptionPtr)
			val2 := deepcopy.Copy(def).(PageLayoutOptionPtr)

			f := NewFile()
			// Get the default value
			assert.NoError(t, f.GetPageLayout(sheet, def), opt)
			// Get again and check
			assert.NoError(t, f.GetPageLayout(sheet, val1), opt)
			if !assert.Equal(t, val1, def, opt) {
				t.FailNow()
			}
			// Set the same value
			assert.NoError(t, f.SetPageLayout(sheet, val1), opt)
			// Get again and check
			assert.NoError(t, f.GetPageLayout(sheet, val1), opt)
			if !assert.Equal(t, val1, def, "%T: value should not have changed", opt) {
				t.FailNow()
			}
			// Set a different value
			assert.NoError(t, f.SetPageLayout(sheet, test.nonDefault), opt)
			assert.NoError(t, f.GetPageLayout(sheet, val1), opt)
			// Get again and compare
			assert.NoError(t, f.GetPageLayout(sheet, val2), opt)
			if !assert.Equal(t, val1, val2, "%T: value should not have changed", opt) {
				t.FailNow()
			}
			// Value should not be the same as the default
			if !assert.NotEqual(t, def, val1, "%T: value should have changed from default", opt) {
				t.FailNow()
			}
			// Restore the default value
			assert.NoError(t, f.SetPageLayout(sheet, def), opt)
			assert.NoError(t, f.GetPageLayout(sheet, val1), opt)
			if !assert.Equal(t, def, val1) {
				t.FailNow()
			}
		})
	}
}

func TestSearchSheet(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "SharedStrings.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	// Test search in a not exists worksheet.
	_, err = f.SearchSheet("Sheet4", "")
	assert.EqualError(t, err, "sheet Sheet4 is not exist")
	var expected []string
	// Test search a not exists value.
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
	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(`<worksheet><sheetData><row r="A"><c r="2" t="str"><v>A</v></c></row></sheetData></worksheet>`))
	f.checked = nil
	result, err = f.SearchSheet("Sheet1", "A")
	assert.EqualError(t, err, "strconv.Atoi: parsing \"A\": invalid syntax")
	assert.Equal(t, []string(nil), result)

	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(`<worksheet><sheetData><row r="2"><c r="A" t="str"><v>A</v></c></row></sheetData></worksheet>`))
	result, err = f.SearchSheet("Sheet1", "A")
	assert.EqualError(t, err, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	assert.Equal(t, []string(nil), result)

	f.Pkg.Store("xl/worksheets/sheet1.xml", []byte(`<worksheet><sheetData><row r="0"><c r="A1" t="str"><v>A</v></c></row></sheetData></worksheet>`))
	result, err = f.SearchSheet("Sheet1", "A")
	assert.EqualError(t, err, "invalid cell coordinates [1, 0]")
	assert.Equal(t, []string(nil), result)
}

func TestSetPageLayout(t *testing.T) {
	f := NewFile()
	// Test set page layout on not exists worksheet.
	assert.EqualError(t, f.SetPageLayout("SheetN"), "sheet SheetN is not exist")
}

func TestGetPageLayout(t *testing.T) {
	f := NewFile()
	// Test get page layout on not exists worksheet.
	assert.EqualError(t, f.GetPageLayout("SheetN"), "sheet SheetN is not exist")
}

func TestSetHeaderFooter(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellStr("Sheet1", "A1", "Test SetHeaderFooter"))
	// Test set header and footer on not exists worksheet.
	assert.EqualError(t, f.SetHeaderFooter("SheetN", nil), "sheet SheetN is not exist")
	// Test set header and footer with illegal setting.
	assert.EqualError(t, f.SetHeaderFooter("Sheet1", &FormatHeaderFooter{
		OddHeader: strings.Repeat("c", MaxFieldLength+1),
	}), newFieldLengthError("OddHeader").Error())

	assert.NoError(t, f.SetHeaderFooter("Sheet1", nil))
	text := strings.Repeat("一", MaxFieldLength)
	assert.NoError(t, f.SetHeaderFooter("Sheet1", &FormatHeaderFooter{
		OddHeader:   text,
		OddFooter:   text,
		EvenHeader:  text,
		EvenFooter:  text,
		FirstHeader: text,
	}))
	assert.NoError(t, f.SetHeaderFooter("Sheet1", &FormatHeaderFooter{
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
	}), ErrDefinedNameduplicate.Error())
	assert.EqualError(t, f.DeleteDefinedName(&DefinedName{
		Name: "No Exist Defined Name",
	}), ErrDefinedNameScope.Error())
	assert.Exactly(t, "Sheet1!$A$2:$D$5", f.GetDefinedName()[1].RefersTo)
	assert.NoError(t, f.DeleteDefinedName(&DefinedName{
		Name: "Amount",
	}))
	assert.Exactly(t, "Sheet1!$A$2:$D$5", f.GetDefinedName()[0].RefersTo)
	assert.Exactly(t, 1, len(f.GetDefinedName()))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestDefinedName.xlsx")))
}

func TestGroupSheets(t *testing.T) {
	f := NewFile()
	sheets := []string{"Sheet2", "Sheet3"}
	for _, sheet := range sheets {
		f.NewSheet(sheet)
	}
	assert.EqualError(t, f.GroupSheets([]string{"Sheet1", "SheetN"}), "sheet SheetN is not exist")
	assert.EqualError(t, f.GroupSheets([]string{"Sheet2", "Sheet3"}), "group worksheet must contain an active worksheet")
	assert.NoError(t, f.GroupSheets([]string{"Sheet1", "Sheet2"}))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestGroupSheets.xlsx")))
}

func TestUngroupSheets(t *testing.T) {
	f := NewFile()
	sheets := []string{"Sheet2", "Sheet3", "Sheet4", "Sheet5"}
	for _, sheet := range sheets {
		f.NewSheet(sheet)
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
	assert.EqualError(t, f.InsertPageBreak("SheetN", "C3"), "sheet SheetN is not exist")
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

	f.NewSheet("Sheet2")
	assert.NoError(t, f.InsertPageBreak("Sheet2", "B2"))
	assert.NoError(t, f.InsertPageBreak("Sheet2", "C2"))
	assert.NoError(t, f.RemovePageBreak("Sheet2", "B2"))

	assert.EqualError(t, f.RemovePageBreak("Sheet1", "A"), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	assert.EqualError(t, f.RemovePageBreak("SheetN", "C3"), "sheet SheetN is not exist")
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
	idx := f.NewSheet("Sheet2")
	ws, ok = f.Sheet.Load("xl/worksheets/sheet2.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).SheetViews = &xlsxSheetViews{SheetView: []xlsxSheetView{}}
	f.SetActiveSheet(idx)
}

func TestSetSheetName(t *testing.T) {
	f := NewFile()
	// Test set worksheet with the same name.
	f.SetSheetName("Sheet1", "Sheet1")
	assert.Equal(t, "Sheet1", f.GetSheetName(0))
}

func TestGetWorkbookPath(t *testing.T) {
	f := NewFile()
	f.Pkg.Delete("_rels/.rels")
	assert.Equal(t, "", f.getWorkbookPath())
}

func TestGetWorkbookRelsPath(t *testing.T) {
	f := NewFile()
	f.Pkg.Delete("xl/_rels/.rels")
	f.Pkg.Store("_rels/.rels", []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="/workbook.xml"/></Relationships>`))
	assert.Equal(t, "_rels/workbook.xml.rels", f.getWorkbookRelsPath())
}

func TestDeleteSheet(t *testing.T) {
	f := NewFile()
	f.SetActiveSheet(f.NewSheet("Sheet2"))
	f.NewSheet("Sheet3")
	f.DeleteSheet("Sheet1")
	assert.Equal(t, "Sheet2", f.GetSheetName(f.GetActiveSheetIndex()))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestDeleteSheet.xlsx")))
	// Test with auto filter defined names
	f = NewFile()
	f.NewSheet("Sheet2")
	f.NewSheet("Sheet3")
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "A"))
	assert.NoError(t, f.SetCellValue("Sheet2", "A1", "A"))
	assert.NoError(t, f.SetCellValue("Sheet3", "A1", "A"))
	assert.NoError(t, f.AutoFilter("Sheet1", "A1", "A1", ""))
	assert.NoError(t, f.AutoFilter("Sheet2", "A1", "A1", ""))
	assert.NoError(t, f.AutoFilter("Sheet3", "A1", "A1", ""))
	f.DeleteSheet("Sheet2")
	f.DeleteSheet("Sheet1")
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestDeleteSheet2.xlsx")))
}

func TestDeleteAndAdjustDefinedNames(t *testing.T) {
	deleteAndAdjustDefinedNames(nil, 0)
	deleteAndAdjustDefinedNames(&xlsxWorkbook{}, 0)
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
	file.NewSheet("sheet1")
	for i := 0; i < 1000; i++ {
		_ = file.SetCellInt("sheet1", "A"+strconv.Itoa(i+1), i)
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
	file.NewSheet("sheet1")
	for i := 0; i < 1000; i++ {
		_ = file.SetCellInt("sheet1", "A"+strconv.Itoa(i+1), i)
	}
	_ = file.Save()
}
