package excelize

import (
	"bufio"
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
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
	// Test WriteToBuffer err
	{
		f, buf := File{Pkg: sync.Map{}}, bytes.Buffer{}
		f.Pkg.Store("/d/", []byte("s"))
		_, err := f.WriteTo(bufio.NewWriter(&buf))
		assert.EqualError(t, err, "zip: write to directory")
		f.Pkg.Delete("/d/")
	}
	// Test file path overflow
	{
		f, buf := File{Pkg: sync.Map{}}, bytes.Buffer{}
		const maxUint16 = 1<<16 - 1
		f.Pkg.Store(strings.Repeat("s", maxUint16+1), nil)
		_, err := f.WriteTo(bufio.NewWriter(&buf))
		assert.EqualError(t, err, "zip: FileHeader.Name too long")
	}
	// Test StreamsWriter err
	{
		f, buf := File{Pkg: sync.Map{}}, bytes.Buffer{}
		f.Pkg.Store("s", nil)
		f.streams = make(map[string]*StreamWriter)
		file, _ := os.Open("123")
		f.streams["s"] = &StreamWriter{rawData: bufferedWriter{tmp: file}}
		_, err := f.WriteTo(bufio.NewWriter(&buf))
		assert.Nil(t, err)
	}
	// Test write with temporary file
	{
		f, buf := File{tempFiles: sync.Map{}}, bytes.Buffer{}
		const maxUint16 = 1<<16 - 1
		f.tempFiles.Store("s", "")
		f.tempFiles.Store(strings.Repeat("s", maxUint16+1), "")
		_, err := f.WriteTo(bufio.NewWriter(&buf))
		assert.EqualError(t, err, "zip: FileHeader.Name too long")
	}
	// Test write with unsupported workbook file format
	{
		f, buf := File{Pkg: sync.Map{}}, bytes.Buffer{}
		f.Pkg.Store("/d", []byte("s"))
		f.Path = "Book1.xls"
		_, err := f.WriteTo(bufio.NewWriter(&buf))
		assert.EqualError(t, err, ErrWorkbookFileFormat.Error())
	}
	// Test write with unsupported charset content types.
	{
		f, buf := NewFile(), bytes.Buffer{}
		f.ContentTypes, f.Path = nil, filepath.Join("test", "TestWriteTo.xlsx")
		f.Pkg.Store(defaultXMLPathContentTypes, MacintoshCyrillicCharset)
		_, err := f.WriteTo(bufio.NewWriter(&buf))
		assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	}
}

func TestClose(t *testing.T) {
	f := NewFile()
	f.tempFiles.Store("/d/", "/d/")
	require.Error(t, f.Close())
}
