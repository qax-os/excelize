package format

import (
	"testing"
	"github.com/stretchr/testify/require"
)

func TestNewConditional(t *testing.T) {
	target :=  Conditional{
		Type: "2_color_scale",
		Criteria: "=",
		MinType: "min",
		MaxType: "max",
		MinColor: "#F8696B",
		MaxColor: "#63BE7B",
	}

	//string
	s, err := NewConditional(`[{"type":"2_color_scale","criteria":"=","min_type":"min","max_type":"max","min_color":"#F8696B","max_color":"#63BE7B"}]`)

	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, []*Conditional{}, s)
	require.Equal(t, []*Conditional{&target}, s)

	//struct
	s, err = NewConditional(target)
	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, []*Conditional{}, s)
	require.Equal(t, []*Conditional{&target}, s)

	//ptr to struct
	s, err = NewConditional(target)
	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, []*Conditional{}, s)
	require.Equal(t, []*Conditional{&target}, s)

	//slice of struct
	s, err = NewConditional([]Conditional{target})
	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, []*Conditional{}, s)
	require.Equal(t, []*Conditional{&target}, s)

	//slice of ptr struct
	s, err = NewConditional([]*Conditional{&target})
	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, []*Conditional{}, s)
	require.Equal(t, []*Conditional{&target}, s)
}

