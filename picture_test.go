package excelize

import (
	"fmt"
	_ "image/png"
	"io/ioutil"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

func BenchmarkAddPictureFromBytes(b *testing.B) {
	f := NewFile()
	imgFile, err := ioutil.ReadFile(filepath.Join("test", "images", "excel.png"))
	if err != nil {
		b.Error("unable to load image for benchmark")
	}
	b.ResetTimer()
	for i := 1; i <= b.N; i++ {
		f.AddPictureFromBytes("Sheet1", fmt.Sprint("A", i), "", "excel", ".png", imgFile)
	}
}

func TestAddPicture(t *testing.T) {
	xlsx, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Test add picture to worksheet with offset and location hyperlink.
	err = xlsx.AddPicture("Sheet2", "I9", filepath.Join("test", "images", "excel.jpg"),
		`{"x_offset": 140, "y_offset": 120, "hyperlink": "#Sheet2!D8", "hyperlink_type": "Location"}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Test add picture to worksheet with offset, external hyperlink and positioning.
	err = xlsx.AddPicture("Sheet1", "F21", filepath.Join("test", "images", "excel.jpg"),
		`{"x_offset": 10, "y_offset": 10, "hyperlink": "https://github.com/360EntSecGroup-Skylar/excelize", "hyperlink_type": "External", "positioning": "oneCell"}`)
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	file, err := ioutil.ReadFile(filepath.Join("test", "images", "excel.jpg"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Test add picture to worksheet from bytes.
	assert.NoError(t, xlsx.AddPictureFromBytes("Sheet1", "Q1", "", "Excel Logo", ".jpg", file))
	// Test add picture to worksheet from bytes with illegal cell coordinates.
	assert.EqualError(t, xlsx.AddPictureFromBytes("Sheet1", "A", "", "Excel Logo", ".jpg", file), `cannot convert cell "A" to coordinates: invalid cell name "A"`)

	// Test write file to given path.
	assert.NoError(t, xlsx.SaveAs(filepath.Join("test", "TestAddPicture.xlsx")))
}

func TestAddPictureErrors(t *testing.T) {
	xlsx, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// Test add picture to worksheet with invalid file path.
	err = xlsx.AddPicture("Sheet1", "G21", filepath.Join("test", "not_exists_dir", "not_exists.icon"), "")
	if assert.Error(t, err) {
		assert.True(t, os.IsNotExist(err), "Expected os.IsNotExist(err) == true")
	}

	// Test add picture to worksheet with unsupport file type.
	err = xlsx.AddPicture("Sheet1", "G21", filepath.Join("test", "Book1.xlsx"), "")
	assert.EqualError(t, err, "unsupported image extension")

	err = xlsx.AddPictureFromBytes("Sheet1", "G21", "", "Excel Logo", "jpg", make([]byte, 1))
	assert.EqualError(t, err, "unsupported image extension")

	// Test add picture to worksheet with invalid file data.
	err = xlsx.AddPictureFromBytes("Sheet1", "G21", "", "Excel Logo", ".jpg", make([]byte, 1))
	assert.EqualError(t, err, "image: unknown format")
}

func TestGetPicture(t *testing.T) {
	xlsx, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	file, raw, err := xlsx.GetPicture("Sheet1", "F21")
	assert.NoError(t, err)
	if !assert.NotEmpty(t, filepath.Join("test", file)) || !assert.NotEmpty(t, raw) ||
		!assert.NoError(t, ioutil.WriteFile(filepath.Join("test", file), raw, 0644)) {

		t.FailNow()
	}

	// Try to get picture from a worksheet with illegal cell coordinates.
	_, _, err = xlsx.GetPicture("Sheet1", "A")
	assert.EqualError(t, err, `cannot convert cell "A" to coordinates: invalid cell name "A"`)

	// Try to get picture from a worksheet that doesn't contain any images.
	file, raw, err = xlsx.GetPicture("Sheet3", "I9")
	assert.EqualError(t, err, "sheet Sheet3 is not exist")
	assert.Empty(t, file)
	assert.Empty(t, raw)

	// Try to get picture from a cell that doesn't contain an image.
	file, raw, err = xlsx.GetPicture("Sheet2", "A2")
	assert.NoError(t, err)
	assert.Empty(t, file)
	assert.Empty(t, raw)

	xlsx.getDrawingRelationships("xl/worksheets/_rels/sheet1.xml.rels", "rId8")
	xlsx.getDrawingRelationships("", "")
	xlsx.getSheetRelationshipsTargetByID("", "")
	xlsx.deleteSheetRelationships("", "")

	// Try to get picture from a local storage file.
	if !assert.NoError(t, xlsx.SaveAs(filepath.Join("test", "TestGetPicture.xlsx"))) {
		t.FailNow()
	}

	xlsx, err = OpenFile(filepath.Join("test", "TestGetPicture.xlsx"))
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	file, raw, err = xlsx.GetPicture("Sheet1", "F21")
	assert.NoError(t, err)
	if !assert.NotEmpty(t, filepath.Join("test", file)) || !assert.NotEmpty(t, raw) ||
		!assert.NoError(t, ioutil.WriteFile(filepath.Join("test", file), raw, 0644)) {

		t.FailNow()
	}

	// Try to get picture from a local storage file that doesn't contain an image.
	file, raw, err = xlsx.GetPicture("Sheet1", "F22")
	assert.NoError(t, err)
	assert.Empty(t, file)
	assert.Empty(t, raw)
}

func TestAddDrawingPicture(t *testing.T) {
	// testing addDrawingPicture with illegal cell coordinates.
	f := NewFile()
	assert.EqualError(t, f.addDrawingPicture("sheet1", "", "A", "", 0, 0, 0, 0, nil), `cannot convert cell "A" to coordinates: invalid cell name "A"`)
}

func TestAddPictureFromBytes(t *testing.T) {
	f := NewFile()
	imgFile, err := ioutil.ReadFile("logo.png")
	if err != nil {
		t.Error("Unable to load logo for test")
	}
	f.AddPictureFromBytes("Sheet1", fmt.Sprint("A", 1), "", "logo", ".png", imgFile)
	f.AddPictureFromBytes("Sheet1", fmt.Sprint("A", 50), "", "logo", ".png", imgFile)
	imageCount := 0
	for fileName := range f.XLSX {
		if strings.Contains(fileName, "media/image") {
			imageCount++
		}
	}
	assert.Equal(t, 1, imageCount, "Duplicate image should only be stored once.")
}
