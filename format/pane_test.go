package format

import (
	"testing"

	"github.com/stretchr/testify/require"
)

func TestNewPanes(t *testing.T) {
	target :=  Panes{
		Freeze:true,
		Split:false,
		XSplit:1,
		YSplit:0,
		TopLeftCell:"B1",
		ActivePane:"topRight",
		Panes:[]Pane{
			{
				SQRef:"K16",
				ActiveCell:"K16",
				Pane:"topRight",
			},
		},
	}

	//string
	s, err := NewPanes(`{"freeze":true,"split":false,"x_split":1,"y_split":0,"top_left_cell":"B1","active_pane":"topRight","panes":[{"sqref":"K16","active_cell":"K16","pane":"topRight"}]}`)

	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Panes{}, s)
	require.Equal(t, &target, s)

	//struct
	s, err = NewPanes(target)
	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Panes{}, s)
	require.Equal(t, &target, s)

	//ptr to struct
	s, err = NewPanes(&target)
	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Panes{}, s)
	require.Equal(t, &target, s)
}

