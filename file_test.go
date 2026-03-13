package excelize

import (
	"archive/zip"
	"bufio"
	"bytes"
	"encoding/binary"
	"io"
	"math"
	"os"
	"path/filepath"
	"strconv"
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
		f.SetZipWriter(func(w io.Writer) ZipWriter { return zip.NewWriter(w) })
		f.Pkg.Store("/d/", []byte("s"))
		_, err := f.WriteTo(bufio.NewWriter(&buf))
		assert.EqualError(t, err, "zip: write to directory")
		f.Pkg.Delete("/d/")
	}
	// Test file path overflow
	{
		f, buf := File{Pkg: sync.Map{}}, bytes.Buffer{}
		f.SetZipWriter(func(w io.Writer) ZipWriter { return zip.NewWriter(w) })
		const maxUint16 = 1<<16 - 1
		f.Pkg.Store(strings.Repeat("s", maxUint16+1), nil)
		_, err := f.WriteTo(bufio.NewWriter(&buf))
		assert.EqualError(t, err, "zip: FileHeader.Name too long")
	}
	// Test StreamsWriter err
	{
		f, buf := File{Pkg: sync.Map{}}, bytes.Buffer{}
		f.SetZipWriter(func(w io.Writer) ZipWriter { return zip.NewWriter(w) })
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
		f.SetZipWriter(func(w io.Writer) ZipWriter { return zip.NewWriter(w) })
		const maxUint16 = 1<<16 - 1
		f.tempFiles.Store("s", "")
		f.tempFiles.Store(strings.Repeat("s", maxUint16+1), "")
		_, err := f.WriteTo(bufio.NewWriter(&buf))
		assert.EqualError(t, err, "zip: FileHeader.Name too long")
	}
	// Test write with unsupported workbook file format
	{
		f, buf := File{Pkg: sync.Map{}}, bytes.Buffer{}
		f.SetZipWriter(func(w io.Writer) ZipWriter { return zip.NewWriter(w) })
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

	// Test non-password writes use file-backed path.
	{
		f := NewFile()
		var gotWriterType string
		f.SetZipWriter(func(w io.Writer) ZipWriter {
			switch w.(type) {
			case *os.File:
				gotWriterType = "file"
			case *bytes.Buffer:
				gotWriterType = "buffer"
			default:
				gotWriterType = "other"
			}
			return zip.NewWriter(w)
		})
		var out bytes.Buffer
		_, err := f.WriteTo(&out)
		require.NoError(t, err)
		assert.Equal(t, "file", gotWriterType)
		assert.NotZero(t, out.Len())
	}
	// Test password writes use buffered path.
	{
		f := NewFile()
		var gotWriterType string
		f.SetZipWriter(func(w io.Writer) ZipWriter {
			switch w.(type) {
			case *os.File:
				gotWriterType = "file"
			case *bytes.Buffer:
				gotWriterType = "buffer"
			default:
				gotWriterType = "other"
			}
			return zip.NewWriter(w)
		})
		var out bytes.Buffer
		_, err := f.WriteTo(&out, Options{Password: "123"})
		require.NoError(t, err)
		assert.Equal(t, "buffer", gotWriterType)
		assert.NotZero(t, out.Len())
	}
	// Test stream writer non-password writes use file-backed path.
	{
		f := NewFile()
		sw, err := f.NewStreamWriter("Sheet1")
		require.NoError(t, err)
		for r := 1; r <= 64; r++ {
			row := make([]interface{}, 64)
			for c := 0; c < len(row); c++ {
				row[c] = strings.Repeat("x", 512)
			}
			cell, err := CoordinatesToCellName(1, r)
			require.NoError(t, err)
			require.NoError(t, sw.SetRow(cell, row))
		}
		require.NoError(t, sw.Flush())

		var gotWriterType string
		f.SetZipWriter(func(w io.Writer) ZipWriter {
			switch w.(type) {
			case *os.File:
				gotWriterType = "file"
			case *bytes.Buffer:
				gotWriterType = "buffer"
			default:
				gotWriterType = "other"
			}
			return zip.NewWriter(w)
		})
		_, err = f.WriteTo(io.Discard)
		require.NoError(t, err)
		assert.Equal(t, "file", gotWriterType)
		require.NoError(t, f.Close())
	}
	// Test writeDirectToWriter returns error when the temporary directory does not exist.
	{
		f := NewFile()
		_, err := f.WriteTo(io.Discard, Options{TmpDir: filepath.Join(os.TempDir(), "excelize_nonexistent_dir")})
		assert.Error(t, err)
	}
}

func TestClose(t *testing.T) {
	f := NewFile()
	f.tempFiles.Store("/d/", "/d/")
	require.Error(t, f.Close())
}

func TestZip64(t *testing.T) {
	f := NewFile()
	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)
	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	for r := range 131 {
		rowData := make([]interface{}, 1000)
		for c := range 1000 {
			rowData[c] = strings.Repeat("c", TotalCellChars)
		}
		cell, err := CoordinatesToCellName(1, r+1)
		assert.NoError(t, err)
		assert.NoError(t, sw.SetRow(cell, rowData))
	}
	assert.NoError(t, sw.Flush())
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestZip64.xlsx")))
	assert.NoError(t, f.Close())

	// Test with filename length overflow
	f = NewFile()
	f.zip64Entries = append(f.zip64Entries, defaultXMLPathSharedStrings)
	buf := new(bytes.Buffer)
	buf.Write([]byte{0x50, 0x4b, 0x03, 0x04})
	buf.Write(make([]byte, 20))
	assert.NoError(t, f.writeZip64LFH(buf))

	// Test with file header less than the required 30 for the fixed header part
	f = NewFile()
	f.zip64Entries = append(f.zip64Entries, defaultXMLPathSharedStrings)
	buf.Reset()
	buf.Write([]byte{0x50, 0x4b, 0x03, 0x04})
	buf.Write(make([]byte, 22))
	binary.Write(buf, binary.LittleEndian, uint16(10))
	buf.Write(make([]byte, 2))
	buf.WriteString("test")
	assert.NoError(t, f.writeZip64LFH(buf))

	// Test that stale zip64Entries from a previous writeToZip call are not carried over.
	f = NewFile()
	f.zip64Entries = []string{"stale/entry.xml"}
	buf.Reset()
	zw := f.ZipWriter(buf)
	assert.NoError(t, f.writeToZip(zw))
	_ = zw.Close()
	assert.Equal(t, -1, inStrSlice(f.zip64Entries, "stale/entry.xml", true))

	t.Run("for_save_zip64_with_in_memory_file_over_4GB", func(t *testing.T) {
		// Test save workbook in ZIP64 format with in memory file with size over 4GB.
		f := NewFile()
		f.Sheet.Delete("xl/worksheets/sheet1.xml")
		f.Pkg.Store("xl/worksheets/sheet1.xml", make([]byte, math.MaxUint32+1))
		_, err := f.WriteToBuffer()
		assert.NoError(t, err)
		assert.NoError(t, f.Close())
	})

	t.Run("for_save_zip64_with_in_temporary_file_over_4GB", func(t *testing.T) {
		// Test save workbook in ZIP64 format with temporary file with size over 4GB.
		if os.Getenv("GITHUB_ACTIONS") == "true" {
			t.Skip()
		}
		f := NewFile()
		f.Pkg.Delete("xl/worksheets/sheet1.xml")
		f.Sheet.Delete("xl/worksheets/sheet1.xml")
		tmp, err := os.CreateTemp(os.TempDir(), "excelize-")
		assert.NoError(t, err)
		assert.NoError(t, tmp.Truncate(math.MaxUint32+1))
		f.tempFiles.Store("xl/worksheets/sheet1.xml", tmp.Name())
		assert.NoError(t, tmp.Close())
		_, err = f.WriteToBuffer()
		assert.NoError(t, err)
		assert.NoError(t, f.Close())
	})

	t.Run("write_zip64_lfh_file_patch_parity", func(t *testing.T) {
		makeLFH := func(name string) []byte {
			h := make([]byte, 30+len(name))
			copy(h[0:4], []byte{0x50, 0x4b, 0x03, 0x04})
			binary.LittleEndian.PutUint16(h[4:6], 20)
			binary.LittleEndian.PutUint16(h[26:28], uint16(len(name)))
			copy(h[30:], []byte(name))
			return h
		}

		entryA := "xl/worksheets/sheet1.xml"
		entryB := "docProps/core.xml"
		initial := append(makeLFH(entryA), makeLFH(entryB)...)

		f := NewFile()
		f.zip64Entries = append(f.zip64Entries, entryA)

		buf := bytes.NewBuffer(append([]byte(nil), initial...))
		require.NoError(t, f.writeZip64LFH(buf))

		tmp, err := os.CreateTemp("", "excelize-zip64-lfh-*")
		require.NoError(t, err)
		defer os.Remove(tmp.Name())
		_, err = tmp.Write(initial)
		require.NoError(t, err)
		require.NoError(t, f.writeZip64LFHToFile(tmp))
		require.NoError(t, tmp.Close())

		got, err := os.ReadFile(tmp.Name())
		require.NoError(t, err)
		assert.Equal(t, buf.Bytes(), got)
		assert.EqualValues(t, 45, binary.LittleEndian.Uint16(got[4:6]))
		offsetSecond := len(makeLFH(entryA))
		assert.EqualValues(t, 20, binary.LittleEndian.Uint16(got[offsetSecond+4:offsetSecond+6]))
	})
}

func TestRemoveTempFiles(t *testing.T) {
	tmp, err := os.CreateTemp("", "excelize-*")
	if err != nil {
		t.Fatal(err)
	}
	tmpName := tmp.Name()
	tmp.Close()
	f := NewFile()
	// fill the tempFiles map with non-existing (erroring on Remove) "files"
	for i := 0; i < 1000; i++ {
		f.tempFiles.Store(strconv.Itoa(i), "/hopefully not existing")
	}
	f.tempFiles.Store("existing", tmpName)

	require.Error(t, f.Close())
	if _, err := os.Stat(tmpName); err == nil {
		t.Errorf("temp file %q still exist", tmpName)
		os.Remove(tmpName)
	}
}
