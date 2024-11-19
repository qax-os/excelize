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
	assert.NoError(t, f.AddShape("Sheet1", &Shape{Cell: "C30", Type: "rect"}))
	assert.EqualError(t, f.AddShape("Sheet3",
		&Shape{
			Cell: "H1",
			Type: "ellipseRibbon",
			Line: ShapeLine{Color: "4286F4"},
			Fill: Fill{Color: []string{"8EB9FF"}},
			Paragraph: []RichTextRun{
				{
					Font: &Font{
						Bold:      true,
						Italic:    true,
						Family:    "Times New Roman",
						Size:      36,
						Color:     "777777",
						Underline: "single",
					},
				},
			},
		},
	), "sheet Sheet3 does not exist")
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
	lineWidth := 1.2
	assert.NoError(t, f.AddShape("Sheet1",
		&Shape{
			Cell: "A1",
			Type: "ellipseRibbon",
			Line: ShapeLine{Color: "4286F4", Width: &lineWidth},
			Fill: Fill{Color: []string{"8EB9FF"}},
			Paragraph: []RichTextRun{
				{
					Font: &Font{
						Bold:      true,
						Italic:    true,
						Family:    "Times New Roman",
						Size:      36,
						Color:     "777777",
						Underline: "single",
					},
				},
			},
			Height: 90,
		}))
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
	// Test add shape with unsupported charset style sheet
	f.Styles = nil
	f.Pkg.Store(defaultXMLPathStyles, MacintoshCyrillicCharset)
	assert.EqualError(t, f.AddShape("Sheet1", &Shape{Cell: "B30", Type: "rect", Paragraph: []RichTextRun{{Text: "Rectangle"}, {}}}), "XML syntax error on line 1: invalid UTF-8")
	// Test add shape with unsupported charset content types
	f = NewFile()
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
	assert.EqualError(t, f.AddShape("Sheet1", &Shape{Cell: "B30", Type: "rect", Paragraph: []RichTextRun{{Text: "Rectangle"}, {}}}), "XML syntax error on line 1: invalid UTF-8")
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
