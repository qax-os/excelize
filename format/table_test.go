package format

import (
	"testing"

	"github.com/stretchr/testify/require"
)

func TestNewTable(t *testing.T) {
	target :=  Table{
		TableStyle: "TableStyleMedium2",
		ShowFirstColumn:true,
		ShowLastColumn:true,
		ShowRowStripes:false,
		ShowColumnStripes:true,
	}

	//string
	s, err := NewTable(`{"table_style":"TableStyleMedium2", "show_first_column":true,"show_last_column":true,"show_row_stripes":false,"show_column_stripes":true}`)

	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Table{}, s)
	require.Equal(t, &target, s)

	//struct
	s, err = NewTable(target)
	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Table{}, s)
	require.Equal(t, &target, s)

	//ptr to struct
	s, err = NewTable(&target)
	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Table{}, s)
	require.Equal(t, &target, s)
}

func TestNewAutoFilter(t *testing.T) {
	target :=  AutoFilter{
		Column: "B",
		Expression:"x <= 1 and x >= blanks",
	}

	//string
	s, err := NewAutoFilter(`{"column":"B","expression":"x <= 1 and x >= blanks"}`)

	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &AutoFilter{}, s)
	require.Equal(t, &target, s)

	//struct
	s, err = NewAutoFilter(target)
	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &AutoFilter{}, s)
	require.Equal(t, &target, s)

	//ptr to struct
	s, err = NewAutoFilter(&target)
	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &AutoFilter{}, s)
	require.Equal(t, &target, s)
}

