package format

import (
	"testing"

	"github.com/stretchr/testify/require"
)

func TestNewStyleSet(t *testing.T) {
	target :=  Style{
		Border: []Border{
			{
				Type:  "left",
				Color: "0000FF",
				Style: -1,
			},
			{
				Type:  "top",
				Color: "00FF00",
				Style: 14,
			},
			{
				Type:  "bottom",
				Color: "FFFF00",
				Style: 5,
			},
			{
				Type:  "right",
				Color: "FF0000",
				Style: 6,
			},
			{
				Type:  "diagonalDown",
				Color: "A020F0",
				Style: 9,
			},
			{
				Type:  "diagonalUp",
				Color: "A020F0",
				Style: 8,
			},
		},
	}

	//string
	s, err := NewStyleSet(`{"border":[{"type":"left","color":"0000FF","style":-1},{"type":"top","color":"00FF00","style":14},{"type":"bottom","color":"FFFF00","style":5},{"type":"right","color":"FF0000","style":6},{"type":"diagonalDown","color":"A020F0","style":9},{"type":"diagonalUp","color":"A020F0","style":8}]}`)

	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Style{}, s)
	require.Equal(t, &target, s)

	//struct
	s, err = NewStyleSet(target)
	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Style{}, s)
	require.Equal(t, &target, s)

	//ptr to struct
	s, err = NewStyleSet(&target)
	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Style{}, s)
	require.Equal(t, &target, s)
}
