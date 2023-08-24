package excelize

import (
	"fmt"
	"image"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	_ "golang.org/x/image/bmp"
	_ "golang.org/x/image/tiff"

	"github.com/stretchr/testify/assert"
)

func BenchmarkAddPictureFromBytes(b *testing.B) {
	f := NewFile()
	imgFile, err := os.ReadFile(filepath.Join("test", "images", "excel.png"))
	if err != nil {
		b.Error("unable to load image for benchmark")
	}
	b.ResetTimer()
	for i := 1; i <= b.N; i++ {
		if err := f.AddPictureFromBytes("Sheet1", fmt.Sprint("A", i), &Picture{Extension: ".png", File: imgFile, Format: &GraphicOptions{AltText: "Excel"}}); err != nil {
			b.Error(err)
		}
	}
}

func TestAddPicture(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)

	// Test add picture to worksheet with offset and location hyperlink
	assert.NoError(t, f.AddPicture("Sheet2", "I9", filepath.Join("test", "images", "excel.jpg"),
		&GraphicOptions{OffsetX: 140, OffsetY: 120, Hyperlink: "#Sheet2!D8", HyperlinkType: "Location"}))
	// Test add picture to worksheet with offset, external hyperlink and positioning
	assert.NoError(t, f.AddPicture("Sheet1", "F21", filepath.Join("test", "images", "excel.jpg"),
		&GraphicOptions{OffsetX: 10, OffsetY: 10, Hyperlink: "https://github.com/xuri/excelize", HyperlinkType: "External", Positioning: "oneCell"}))

	file, err := os.ReadFile(filepath.Join("test", "images", "excel.png"))
	assert.NoError(t, err)

	// Test add picture to worksheet with autofit
	assert.NoError(t, f.AddPicture("Sheet1", "A30", filepath.Join("test", "images", "excel.jpg"), &GraphicOptions{AutoFit: true}))
	assert.NoError(t, f.AddPicture("Sheet1", "B30", filepath.Join("test", "images", "excel.jpg"), &GraphicOptions{OffsetX: 10, OffsetY: 10, AutoFit: true}))
	_, err = f.NewSheet("AddPicture")
	assert.NoError(t, err)
	assert.NoError(t, f.SetRowHeight("AddPicture", 10, 30))
	assert.NoError(t, f.MergeCell("AddPicture", "B3", "D9"))
	assert.NoError(t, f.MergeCell("AddPicture", "B1", "D1"))
	assert.NoError(t, f.AddPicture("AddPicture", "C6", filepath.Join("test", "images", "excel.jpg"), &GraphicOptions{AutoFit: true}))
	assert.NoError(t, f.AddPicture("AddPicture", "A1", filepath.Join("test", "images", "excel.jpg"), &GraphicOptions{AutoFit: true}))

	// Test add picture to worksheet from bytes
	assert.NoError(t, f.AddPictureFromBytes("Sheet1", "Q1", &Picture{Extension: ".png", File: file, Format: &GraphicOptions{AltText: "Excel Logo"}}))
	// Test add picture to worksheet from bytes with illegal cell reference
	assert.EqualError(t, f.AddPictureFromBytes("Sheet1", "A", &Picture{Extension: ".png", File: file, Format: &GraphicOptions{AltText: "Excel Logo"}}), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())

	for cell, ext := range map[string]string{"Q8": "gif", "Q15": "jpg", "Q22": "tif", "Q28": "bmp"} {
		assert.NoError(t, f.AddPicture("Sheet1", cell, filepath.Join("test", "images", fmt.Sprintf("excel.%s", ext)), nil))
	}

	// Test write file to given path
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddPicture1.xlsx")))
	assert.NoError(t, f.Close())

	// Test get pictures after inserting a new picture from a workbook which contains existing pictures
	f, err = OpenFile(filepath.Join("test", "TestAddPicture1.xlsx"))
	assert.NoError(t, err)
	assert.NoError(t, f.AddPicture("Sheet1", "A30", filepath.Join("test", "images", "excel.jpg"), nil))
	pics, err := f.GetPictures("Sheet1", "A30")
	assert.NoError(t, err)
	assert.Len(t, pics, 2)
	assert.NoError(t, f.Close())

	// Test add picture with unsupported charset content types
	f = NewFile()
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, f.AddPictureFromBytes("Sheet1", "Q1", &Picture{Extension: ".png", File: file, Format: &GraphicOptions{AltText: "Excel Logo"}}), "XML syntax error on line 1: invalid UTF-8")

	// Test add picture with invalid sheet name
	assert.EqualError(t, f.AddPicture("Sheet:1", "A1", filepath.Join("test", "images", "excel.jpg"), nil), ErrSheetNameInvalid.Error())
}

func TestAddPictureErrors(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)

	// Test add picture to worksheet with invalid file path
	assert.Error(t, f.AddPicture("Sheet1", "G21", filepath.Join("test", "not_exists_dir", "not_exists.icon"), nil))

	// Test add picture to worksheet with unsupported file type
	assert.EqualError(t, f.AddPicture("Sheet1", "G21", filepath.Join("test", "Book1.xlsx"), nil), ErrImgExt.Error())

	assert.EqualError(t, f.AddPictureFromBytes("Sheet1", "G21", &Picture{Extension: "jpg", File: make([]byte, 1), Format: &GraphicOptions{AltText: "Excel Logo"}}), ErrImgExt.Error())

	// Test add picture to worksheet with invalid file data
	assert.EqualError(t, f.AddPictureFromBytes("Sheet1", "G21", &Picture{Extension: ".jpg", File: make([]byte, 1), Format: &GraphicOptions{AltText: "Excel Logo"}}), image.ErrFormat.Error())

	// Test add picture with custom image decoder and encoder
	decode := func(r io.Reader) (image.Image, error) { return nil, nil }
	decodeConfig := func(r io.Reader) (image.Config, error) { return image.Config{Height: 100, Width: 90}, nil }
	for cell, ext := range map[string]string{"Q1": "emf", "Q7": "wmf", "Q13": "emz", "Q19": "wmz"} {
		image.RegisterFormat(ext, "", decode, decodeConfig)
		assert.NoError(t, f.AddPicture("Sheet1", cell, filepath.Join("test", "images", fmt.Sprintf("excel.%s", ext)), nil))
	}
	assert.NoError(t, f.AddPicture("Sheet1", "Q25", "excelize.svg", &GraphicOptions{ScaleX: 2.8}))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddPicture2.xlsx")))
	assert.NoError(t, f.Close())
}

func TestGetPicture(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.AddPicture("Sheet1", "A1", filepath.Join("test", "images", "excel.png"), nil))
	pics, err := f.GetPictures("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Len(t, pics[0].File, 13233)
	assert.Empty(t, pics[0].Format.AltText)

	f, err = prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	pics, err = f.GetPictures("Sheet1", "F21")
	assert.NoError(t, err)
	if !assert.NotEmpty(t, filepath.Join("test", fmt.Sprintf("image1%s", pics[0].Extension))) || !assert.NotEmpty(t, pics[0].File) ||
		!assert.NoError(t, os.WriteFile(filepath.Join("test", fmt.Sprintf("image1%s", pics[0].Extension)), pics[0].File, 0o644)) {
		t.FailNow()
	}

	// Try to get picture from a worksheet with illegal cell reference
	_, err = f.GetPictures("Sheet1", "A")
	assert.EqualError(t, err, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())

	// Try to get picture from a worksheet that doesn't contain any images
	pics, err = f.GetPictures("Sheet3", "I9")
	assert.EqualError(t, err, "sheet Sheet3 does not exist")
	assert.Len(t, pics, 0)

	// Try to get picture from a cell that doesn't contain an image
	pics, err = f.GetPictures("Sheet2", "A2")
	assert.NoError(t, err)
	assert.Len(t, pics, 0)

	// Test get picture with invalid sheet name
	_, err = f.GetPictures("Sheet:1", "A2")
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())

	f.getDrawingRelationships("xl/worksheets/_rels/sheet1.xml.rels", "rId8")
	f.getDrawingRelationships("", "")
	f.getSheetRelationshipsTargetByID("", "")
	f.deleteSheetRelationships("", "")

	// Try to get picture from a local storage file.
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestGetPicture.xlsx")))

	f, err = OpenFile(filepath.Join("test", "TestGetPicture.xlsx"))
	assert.NoError(t, err)

	pics, err = f.GetPictures("Sheet1", "F21")
	assert.NoError(t, err)
	if !assert.NotEmpty(t, filepath.Join("test", fmt.Sprintf("image1%s", pics[0].Extension))) || !assert.NotEmpty(t, pics[0].File) ||
		!assert.NoError(t, os.WriteFile(filepath.Join("test", fmt.Sprintf("image1%s", pics[0].Extension)), pics[0].File, 0o644)) {
		t.FailNow()
	}

	// Try to get picture from a local storage file that doesn't contain an image
	pics, err = f.GetPictures("Sheet1", "F22")
	assert.NoError(t, err)
	assert.Len(t, pics, 0)
	assert.NoError(t, f.Close())

	// Try to get picture with one cell anchor
	f, err = OpenFile(filepath.Join("test", "TestGetPicture.xlsx"))
	assert.NoError(t, err)
	f.Pkg.Store("xl/drawings/drawing2.xml", []byte(`<xdr:wsDr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:oneCellAnchor><xdr:from><xdr:col>10</xdr:col><xdr:row>15</xdr:row></xdr:from><xdr:to><xdr:col>13</xdr:col><xdr:row>22</xdr:row></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2"></xdr:cNvPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"></a:blip></xdr:blipFill></xdr:pic></xdr:oneCellAnchor></xdr:wsDr>`))
	pics, err = f.GetPictures("Sheet2", "K16")
	assert.NoError(t, err)
	assert.Len(t, pics, 1)

	// Test get picture from none drawing worksheet
	f = NewFile()
	pics, err = f.GetPictures("Sheet1", "F22")
	assert.NoError(t, err)
	assert.Len(t, pics, 0)
	f, err = prepareTestBook1()
	assert.NoError(t, err)

	// Test get pictures with unsupported charset
	path := "xl/drawings/drawing1.xml"
	f.Pkg.Store(path, MacintoshCyrillicCharset)
	_, err = f.getPicture(20, 5, path, "xl/drawings/_rels/drawing2.xml.rels")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	f.Drawings.Delete(path)
	_, err = f.getPicture(20, 5, path, "xl/drawings/_rels/drawing2.xml.rels")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
}

func TestAddDrawingPicture(t *testing.T) {
	// Test addDrawingPicture with illegal cell reference
	f := NewFile()
	opts := &GraphicOptions{PrintObject: boolPtr(true), Locked: boolPtr(false)}
	assert.EqualError(t, f.addDrawingPicture("sheet1", "", "A", "", 0, 0, image.Config{}, opts), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	// Test addDrawingPicture with invalid positioning types
	assert.Equal(t, f.addDrawingPicture("sheet1", "", "A1", "", 0, 0, image.Config{}, &GraphicOptions{Positioning: "x"}), ErrParameterInvalid)

	path := "xl/drawings/drawing1.xml"
	f.Pkg.Store(path, MacintoshCyrillicCharset)
	assert.EqualError(t, f.addDrawingPicture("sheet1", path, "A1", "", 0, 0, image.Config{}, opts), "XML syntax error on line 1: invalid UTF-8")
}

func TestAddPictureFromBytes(t *testing.T) {
	f := NewFile()
	imgFile, err := os.ReadFile("logo.png")
	assert.NoError(t, err, "Unable to load logo for test")

	assert.NoError(t, f.AddPictureFromBytes("Sheet1", fmt.Sprint("A", 1), &Picture{Extension: ".png", File: imgFile, Format: &GraphicOptions{AltText: "logo"}}))
	assert.NoError(t, f.AddPictureFromBytes("Sheet1", fmt.Sprint("A", 50), &Picture{Extension: ".png", File: imgFile, Format: &GraphicOptions{AltText: "logo"}}))
	imageCount := 0
	f.Pkg.Range(func(fileName, v interface{}) bool {
		if strings.Contains(fileName.(string), "media/image") {
			imageCount++
		}
		return true
	})
	assert.Equal(t, 1, imageCount, "Duplicate image should only be stored once.")
	assert.EqualError(t, f.AddPictureFromBytes("SheetN", fmt.Sprint("A", 1), &Picture{Extension: ".png", File: imgFile, Format: &GraphicOptions{AltText: "logo"}}), "sheet SheetN does not exist")
	// Test add picture from bytes with invalid sheet name
	assert.EqualError(t, f.AddPictureFromBytes("Sheet:1", fmt.Sprint("A", 1), &Picture{Extension: ".png", File: imgFile, Format: &GraphicOptions{AltText: "logo"}}), ErrSheetNameInvalid.Error())
}

func TestDeletePicture(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	assert.NoError(t, err)
	assert.NoError(t, f.DeletePicture("Sheet1", "A1"))
	assert.NoError(t, f.AddPicture("Sheet1", "P1", filepath.Join("test", "images", "excel.jpg"), nil))
	assert.NoError(t, f.DeletePicture("Sheet1", "P1"))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestDeletePicture.xlsx")))
	// Test delete picture on not exists worksheet
	assert.EqualError(t, f.DeletePicture("SheetN", "A1"), "sheet SheetN does not exist")
	// Test delete picture with invalid sheet name
	assert.EqualError(t, f.DeletePicture("Sheet:1", "A1"), ErrSheetNameInvalid.Error())
	// Test delete picture with invalid coordinates
	assert.EqualError(t, f.DeletePicture("Sheet1", ""), newCellNameToCoordinatesError("", newInvalidCellNameError("")).Error())
	assert.NoError(t, f.Close())
	// Test delete picture on no chart worksheet
	assert.NoError(t, NewFile().DeletePicture("Sheet1", "A1"))
}

func TestDrawingResize(t *testing.T) {
	f := NewFile()
	// Test calculate drawing resize on not exists worksheet
	_, _, _, _, err := f.drawingResize("SheetN", "A1", 1, 1, nil)
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Test calculate drawing resize with invalid coordinates
	_, _, _, _, err = f.drawingResize("Sheet1", "", 1, 1, nil)
	assert.EqualError(t, err, newCellNameToCoordinatesError("", newInvalidCellNameError("")).Error())
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).MergeCells = &xlsxMergeCells{Cells: []*xlsxMergeCell{{Ref: "A:A"}}}
	assert.EqualError(t, f.AddPicture("Sheet1", "A1", filepath.Join("test", "images", "excel.jpg"), &GraphicOptions{AutoFit: true}), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
}

func TestSetContentTypePartImageExtensions(t *testing.T) {
	f := NewFile()
	// Test set content type part image extensions with unsupported charset content types
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, f.setContentTypePartImageExtensions(), "XML syntax error on line 1: invalid UTF-8")
}

func TestSetContentTypePartVMLExtensions(t *testing.T) {
	f := NewFile()
	// Test set content type part VML extensions with unsupported charset content types
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, f.setContentTypePartVMLExtensions(), "XML syntax error on line 1: invalid UTF-8")
}

func TestAddContentTypePart(t *testing.T) {
	f := NewFile()
	// Test add content type part with unsupported charset content types
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, f.addContentTypePart(0, "unknown"), "XML syntax error on line 1: invalid UTF-8")
}
