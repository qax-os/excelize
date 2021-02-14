package excelize

import (
	"bufio"
	"bytes"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

func BenchmarkWrite(b *testing.B) {
	const s = "This is test data"
	for i := 0; i < b.N; i++ {
		f := NewFile()
		for row := 1; row <= 10000; row++ {
			for col := 1; col <= 20; col++ {
				val, err := CoordinatesToCellName(col, row)
				if err != nil {
					b.Error(err)
				}
				if err := f.SetCellValue("Sheet1", val, s); err != nil {
					b.Error(err)
				}
			}
		}
		// Save spreadsheet by the given path.
		err := f.SaveAs("./test.xlsx")
		if err != nil {
			b.Error(err)
		}
	}
}

func TestWriteTo(t *testing.T) {
	f := File{}
	buf := bytes.Buffer{}
	f.XLSX = make(map[string][]byte)
	f.XLSX["/d/"] = []byte("s")
	_, err := f.WriteTo(bufio.NewWriter(&buf))
	assert.EqualError(t, err, "zip: write to directory")
	delete(f.XLSX, "/d/")
	// Test file path overflow
	const maxUint16 = 1<<16 - 1
	f.XLSX[strings.Repeat("s", maxUint16+1)] = nil
	_, err = f.WriteTo(bufio.NewWriter(&buf))
	assert.EqualError(t, err, "zip: FileHeader.Name too long")
}
