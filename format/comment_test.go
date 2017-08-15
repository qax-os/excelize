package format

import (
	"testing"

	"github.com/stretchr/testify/require"
)

func TestNewComment(t *testing.T) {
	target :=  Comment{
		Author: "Excelize: ",
		Text: "This is a comment.",
	}

	//string
	s, err := NewComment(`{"author":"Excelize: ","text":"This is a comment."}`)

	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Comment{}, s)
	require.Equal(t, &target, s)

	//struct
	s, err = NewComment(target)
	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Comment{}, s)
	require.Equal(t, &target, s)

	//ptr to struct
	s, err = NewComment(&target)
	require.NotNil(t, s)
	require.Nil(t, err)

	require.IsType(t, &Comment{}, s)
	require.Equal(t, &target, s)
}

