package excelize

import (
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
)

func TestWriteStructsIntoFile(t *testing.T) {
	t.Run("Should write structs into file with header", func(t *testing.T) {
		f := NewFile()
		type testStruct struct {
			Column1 string `column:"A" columnHeader:"Column 1"`
			Column2 string `column:"B" columnHeader:"Column 2"`
		}
		var testStructs = []testStruct{
			{Column1: "1", Column2: "2"},
			{Column1: "3", Column2: "4"},
		}
		assert.NoError(t, WriteStructsIntoFile(f, testStructs, &ModelTableOptions{HasHeader: true}))

		rows, err := f.GetRows("Sheet1")
		assert.NoError(t, err)
		assert.Equal(t, 3, len(rows))
	})

	t.Run("Should write structs into file without header", func(t *testing.T) {
		f := NewFile()
		type testStruct struct {
			Column1 string `column:"A" columnHeader:"Column 1"`
			Column2 string `column:"B"`
		}
		var testStructs = []testStruct{
			{Column1: "1", Column2: "2"},
			{Column1: "3", Column2: "4"},
		}
		assert.NoError(t, WriteStructsIntoFile(f, testStructs, &ModelTableOptions{HasHeader: true}))

		rows, err := f.GetRows("Sheet1")
		assert.NoError(t, err)
		assert.Equal(t, 3, len(rows))
	})

	t.Run("Should write structs into file with header and nested values", func(t *testing.T) {
		f := NewFile()
		amount := 5
		testString := "test string"
		// fields with no tag should be ignored
		type address struct {
			City    string `columnInnerValue:"City:"`
			Country string `columnInnerValue:"Country:"`
			Number  int
		}
		type user struct {
			Name       string    `column:"A"`
			Age        int       `column:"B"`
			AddressStr string    `column:"C"`
			Birthdate  time.Time `column:"D" columnHeader:"Birthday"`
			Pointer    *string   `column:"E"`
			Address    address   `column:"F" columnHeader:"Indirizzo"`
			Tst        string
		}
		testUser := user{
			Name:       "John Doe",
			Age:        30,
			AddressStr: "123 Main St",
			Birthdate:  time.Date(1993, 1, 1, 0, 0, 0, 0, time.UTC),
			Pointer:    &testString,
			Address: address{
				City:    "New York",
				Country: "USA",
				Number:  123,
			},
		}
		var users []user
		for i := 0; i < amount; i++ {
			users = append(users, testUser)
		}

		assert.NoError(t, WriteStructsIntoFile[user](f, users, &ModelTableOptions{HasHeader: true}))

		rows, err := f.GetRows("Sheet1")
		assert.NoError(t, err)
		// adding 1 for header
		assert.Equal(t, amount+1, len(rows))

	})

	t.Run("Should write structs into file with no options", func(t *testing.T) {
		f := NewFile()
		type testStruct struct {
			Column1 string `column:"A" columnHeader:"Column 1"`
			Column2 string `column:"B" columnHeader:"Column 2"`
		}
		var testStructs = []testStruct{
			{Column1: "1", Column2: "2"},
			{Column1: "3", Column2: "4"},
		}
		assert.NoError(t, WriteStructsIntoFile(f, testStructs, nil))

		rows, err := f.GetRows("Sheet1")
		assert.NoError(t, err)
		assert.Equal(t, len(testStructs), len(rows))
	})

	t.Run("Should return error if the file is nil", func(t *testing.T) {
		type testStruct struct {
			Column1 string `column:"A" columnHeader:"Column 1"`
			Column2 string `column:"B" columnHeader:"Column 2"`
		}
		var testStructs = []testStruct{
			{Column1: "1", Column2: "2"},
			{Column1: "3", Column2: "4"},
		}
		assert.Error(t, WriteStructsIntoFile(nil, testStructs, &ModelTableOptions{HasHeader: true}))
	})

}

func TestGetTagValues(t *testing.T) {
	t.Run("Should get tag values", func(t *testing.T) {
		type testStruct struct {
			Column1 string  `column:"A" columnHeader:"Column 1"`
			Column2 string  `column:"B" columnHeader:"Column 2"`
			Column3 *string `column:"C" columnHeader:"Column 3"`
		}
		var t1 testStruct
		fieldColumnMap := getTagValues(t1, "column")
		columnAliasMap := getTagValues(t1, "columnHeader")

		assert.Equal(t, 3, len(fieldColumnMap))
		assert.Equal(t, 3, len(columnAliasMap))
	})
	t.Run("Should get tag values using pointers", func(t *testing.T) {
		type testStruct struct {
			Column1 string  `column:"A" columnHeader:"Column 1"`
			Column2 string  `column:"B" columnHeader:"Column 2"`
			Column3 *string `column:"C" columnHeader:"Column 3"`
		}
		var t1 testStruct
		fieldColumnMap := getTagValues(&t1, "column")
		columnAliasMap := getTagValues(&t1, "columnHeader")

		assert.Equal(t, 3, len(fieldColumnMap))
		assert.Equal(t, 3, len(columnAliasMap))
	})
}

func TestConstructRows(t *testing.T) {
	type testStruct struct {
		Column1 string `column:"A" columnHeader:"Column 1"`
		Column2 string `column:"B" columnHeader:"Column 2"`
	}
	var testStructs = []testStruct{
		{Column1: "1", Column2: "2"},
		{Column1: "3", Column2: "4"},
	}
	rows := constructRows(testStructs)
	assert.Equal(t, 2, len(rows))
}

func TestGetFieldValue(t *testing.T) {
	type sample struct {
		A string
		B int
	}

	t.Run("should return true with a string", func(t *testing.T) {
		s := sample{A: "test", B: 42}
		val, ok := getFieldValue(s, "A")
		assert.True(t, ok)
		assert.Equal(t, "test", val)
	})

	t.Run("should return true with an int", func(t *testing.T) {
		s := sample{A: "test", B: 42}
		val, ok := getFieldValue(s, "B")
		assert.True(t, ok)
		assert.Equal(t, 42, val)
	})

	t.Run("should return false with no field", func(t *testing.T) {
		s := sample{A: "test", B: 42}
		val, ok := getFieldValue(s, "C")
		assert.False(t, ok)
		assert.Nil(t, val)
	})

	t.Run("should return true with a pointer to a struct", func(t *testing.T) {
		s := sample{A: "test", B: 42}
		ps := &s
		val, ok := getFieldValue(ps, "A")
		assert.True(t, ok)
		assert.Equal(t, "test", val)
	})

	t.Run("should return false", func(t *testing.T) {
		val, ok := getFieldValue(123, "A")
		assert.False(t, ok)
		assert.Nil(t, val)
	})
}
