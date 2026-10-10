package excelize

import (
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestAddShape(t *testing.T) {
	f, err := prepareTestBook1()
	if !assert.NoError(t, err) {
		t.FailNow()
	}
	assert.NoError(t, f.AddShape("Sheet1", &Shape{
		Cell: "A30",
		Type: "rect",
		Paragraph: []RichTextRun{
			{Text: "Rectangle", Font: &Font{Color: "CD5C5C"}},
			{Text: "Shape", Font: &Font{Bold: true, Color: "2980B9"}},
		},
	}))
	assert.NoError(t, f.AddShape("Sheet1", &Shape{Cell: "B30", Type: "rect", Paragraph: []RichTextRun{{Text: "Rectangle"}, {}}}))
	shape1 := Shape{Cell: "C30", Type: "rect", Width: 160, Height: 160, Format: GraphicOptions{Name: "Shape 8"}}
	assert.NoError(t, f.AddShape("Sheet1", &shape1))
	shapes, err := f.GetShapes("Sheet1")
	assert.NoError(t, err)
	assert.Len(t, shapes, 3)
	assert.Equal(t, shape1, shapes[2])
	// Test add shape with invalid positioning types
	assert.Equal(t, newInvalidOptionalValue("Positioning", "x", supportedPositioning), f.AddShape("Sheet1", &Shape{Cell: "C30", Type: "rect", Format: GraphicOptions{Positioning: "x"}}))
	assert.EqualError(t, f.AddShape("Sheet3", &Shape{Cell: "C30", Type: "rect"}), "sheet Sheet3 does not exist")
	assert.Equal(t, ErrParameterInvalid, f.AddShape("Sheet3", nil))
	assert.Equal(t, ErrParameterInvalid, f.AddShape("Sheet1", &Shape{Cell: "A1"}))
	assert.Equal(t, newCellNameToCoordinatesError("A", newInvalidCellNameError("A")), f.AddShape("Sheet1", &Shape{
		Cell: "A",
		Type: "rect",
		Paragraph: []RichTextRun{
			{Text: "Rectangle", Font: &Font{Color: "CD5C5C"}},
			{Text: "Shape", Font: &Font{Bold: true, Color: "2980B9"}},
		},
	}))
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddShape1.xlsx")))

	// Test add first shape for given sheet
	f = NewFile()
	width := 1.2
	shape2 := Shape{
		Cell: "A1",
		Type: "ellipseRibbon",
		Line: LineOptions{
			Fill: Fill{
				Type: "pattern", Color: []string{"4286F4"}, Pattern: 1,
			},
			Width: width,
		},
		Fill: Fill{Color: []string{"8EB9FF"}, Transparency: 60},
		Format: GraphicOptions{
			AltText:     "Shape",
			Name:        "Shape 1",
			PrintObject: boolPtr(true),
			Locked:      boolPtr(false),
			ScaleX:      0.8,
			ScaleY:      0.8,
			Positioning: "oneCell",
		},
		Paragraph: []RichTextRun{
			{
				Font: &Font{
					Bold:      true,
					Italic:    true,
					Family:    "Times New Roman",
					Size:      18,
					Color:     "777777",
					Underline: "sng",
				},
				Text: "Shape",
			},
		},
		Height: 90,
	}
	assert.NoError(t, f.AddShape("Sheet1", &shape2))
	shapes, err = f.GetShapes("Sheet1")
	assert.NoError(t, err)
	assert.Len(t, shapes, 1)
	assert.Equal(t, shape2, shapes[0])
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddShape2.xlsx")))
	// Test add shape with invalid sheet name
	assert.Equal(t, ErrSheetNameInvalid, f.AddShape("Sheet:1", &Shape{
		Cell: "A30",
		Type: "rect",
		Paragraph: []RichTextRun{
			{Text: "Rectangle", Font: &Font{Color: "CD5C5C"}},
			{Text: "Shape", Font: &Font{Bold: true, Color: "2980B9"}},
		},
	}))
	// Test add shape with transparency value exceeds limit
	assert.Equal(t, ErrTransparency, f.AddShape("Sheet1", &Shape{Cell: "B30", Type: "rect", Fill: Fill{Color: []string{"8EB9FF"}, Transparency: 110}}))
	// Test add shape with unsupported charset style sheet
	f.Styles = nil
	f.Pkg.Store(defaultXMLPathStyles, MacintoshCyrillicCharset)
	assert.EqualError(t, f.AddShape("Sheet1", &Shape{Cell: "B30", Type: "rect", Paragraph: []RichTextRun{{Text: "Rectangle"}, {}}}), "XML syntax error on line 1: invalid UTF-8")
	// Test add shape with unsupported charset content types
	f = NewFile()
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, f.AddShape("Sheet1", &Shape{Cell: "B30", Type: "rect", Paragraph: []RichTextRun{{Text: "Rectangle"}, {}}}), "XML syntax error on line 1: invalid UTF-8")

	// Test get shapes after from a worksheet which contains existing shapes
	f, err = OpenFile(filepath.Join("test", "TestAddShape1.xlsx"))
	assert.NoError(t, err)
	shapes, err = f.GetShapes("Sheet1")
	assert.NoError(t, err)
	assert.Len(t, shapes, 3)
	assert.Equal(t, shape1, shapes[2])
	// Test get shapes from a worksheet without shapes
	shapes, err = f.GetShapes("Sheet2")
	assert.NoError(t, err)
	assert.Empty(t, shapes)
	// Test get shapes after from a worksheet which contains existing shapes
	f, err = OpenFile(filepath.Join("test", "TestAddShape2.xlsx"))
	assert.NoError(t, err)
	shapes, err = f.GetShapes("Sheet1")
	assert.NoError(t, err)
	assert.Len(t, shapes, 1)
	assert.Equal(t, shape2, shapes[0])
	// Test get shapes with unsupported charset
	path := "xl/drawings/drawing1.xml"
	f.Drawings.Delete(path)
	f.Pkg.Store(path, MacintoshCyrillicCharset)
	_, err = f.GetShapes("Sheet1")
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	// Test get shapes from a sheet without shapes
	f = NewFile()
	shapes, err = f.GetShapes("Sheet1")
	assert.NoError(t, err)
	assert.Empty(t, shapes)
	// Test get shapes on not exists worksheet
	_, err = f.GetShapes("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	assert.Nil(t, f.extractShapeFromAnchor("Sheet1", &xdrCellAnchor{Sp: &xdrSp{}, From: &xlsxFrom{Col: -1, Row: 0}}))
	assert.Nil(t, f.extractShapeFromDecodeAnchor("Sheet1", "<from><col>-1</col><row>0</row></from><sp/>"))
}

func TestAddDrawingShape(t *testing.T) {
	f := NewFile()
	path := "xl/drawings/drawing1.xml"
	f.Pkg.Store(path, MacintoshCyrillicCharset)
	assert.EqualError(t, f.addDrawingShape("sheet1", path, "A1",
		&Shape{
			Width:  defaultShapeSize,
			Height: defaultShapeSize,
			Format: GraphicOptions{
				PrintObject: boolPtr(true),
				Locked:      boolPtr(false),
			},
		},
	), "XML syntax error on line 1: invalid UTF-8")
}
