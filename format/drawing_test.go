package format

import (
	"testing"

	"github.com/stretchr/testify/require"
)

func TestNewPicture(t *testing.T) {
	target :=  Picture{
		FPrintsWithSheet: true,
		FLocksWithSheet:  false,
		NoChangeAspect:   false,
		OffsetX:          140,
		OffsetY:          120,
		XScale:           1.0,
		YScale:           1.0,
	}

	//string
	s, err := NewPicture(`{"x_offset": 140, "y_offset": 120}`)

	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Picture{}, s)
	require.Equal(t, &target, s)

	//struct
	s, err = NewPicture(target)
	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Picture{}, s)
	require.Equal(t, &target, s)

	//ptr to struct
	s, err = NewPicture(&target)
	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Picture{}, s)
	require.Equal(t, &target, s)
}

func TestNewShape(t *testing.T) {
	target :=  Shape{
		Type: "rect",
		Width:  160,
		Height: 160,
		Format: Picture{
			FPrintsWithSheet: true,
			FLocksWithSheet:  false,
			NoChangeAspect:   false,
			OffsetX:          0,
			OffsetY:          0,
			XScale:           1.0,
			YScale:           1.0,
		},
		Paragraph:[]ShapeParagraph{
			{
				Text: "Rectangle",
				Font: Font{
					Color: "CD5C5C",
				},
			},
			{
				Text: "Shape",
				Font: Font{
					Bold:true,
					Color: "2980B9",
				},
			},
		},
	}

	//string
	s, err := NewShape(`{"type":"rect","paragraph":[{"text":"Rectangle","font":{"color":"CD5C5C"}},{"text":"Shape","font":{"bold":true,"color":"2980B9"}}]}`)

	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Shape{}, s)
	require.Equal(t, &target, s)

	//struct
	s, err = NewShape(target)
	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Shape{}, s)
	require.Equal(t, &target, s)

	//ptr to struct
	s, err = NewShape(&target)
	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Shape{}, s)
	require.Equal(t, &target, s)
}
