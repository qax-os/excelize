package excelize

import (
	"bufio"
	"bytes"
	"encoding/binary"
	"errors"
	"io"
	"io/fs"
	"math"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"sync"
	"testing"

	"github.com/klauspost/compress/zip"
	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

// errZipWriter is a mock ZipWriter whose Create and Close can be configured to
// return errors.
type errZipWriter struct {
	createFunc func(string) (io.Writer, error)
	closeErr   error
}

func (m *errZipWriter) Create(name string) (io.Writer, error) {
	if m.createFunc != nil {
		return m.createFunc(name)
	}
	return &bytes.Buffer{}, nil
}

func (m *errZipWriter) AddFS(fs.FS) error { return nil }

func (m *errZipWriter) Close() error { return m.closeErr }

type errWriter struct{ err error }

func (e *errWriter) Write([]byte) (int, error) { return 0, e.err }

func BenchmarkWrite(b *testing.B) {
	const s = "This is test data"
	for b.Loop() {
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
		if err := f.SaveAs("test.xlsx"); err != nil {
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
	// Test WriteToBuffer with ZipWriter Close error
	{
		f := NewFile()
		f.SetZipWriter(func(w io.Writer) ZipWriter {
			return &errZipWriter{closeErr: errors.New("close error")}
		})
		_, err := f.WriteTo(bufio.NewWriter(&bytes.Buffer{}))
		assert.EqualError(t, err, "close error")
	}
	// Test writeToZip with stream Create error
	{
		f := NewFile()
		f.streams = make(map[string]*StreamWriter)
		f.streams["s"] = &StreamWriter{rawData: bufferedWriter{}}
		f.SetZipWriter(func(w io.Writer) ZipWriter {
			return &errZipWriter{
				createFunc: func(name string) (io.Writer, error) {
					if name == "s" {
						return nil, errors.New("create stream error")
					}
					return &bytes.Buffer{}, nil
				},
			}
		})
		_, err := f.WriteTo(bufio.NewWriter(&bytes.Buffer{}))
		assert.EqualError(t, err, "create stream error")
	}
	// Test writeToZip with stream rawData.Reader() error
	{
		f := NewFile()
		f.streams = make(map[string]*StreamWriter)
		tmp, err := os.CreateTemp("", "excelize-test-*")
		assert.NoError(t, err)
		assert.NoError(t, tmp.Close())
		f.streams["s"] = &StreamWriter{rawData: bufferedWriter{tmp: tmp}}
		_, err = f.WriteTo(bufio.NewWriter(&bytes.Buffer{}))
		assert.Error(t, err)
	}
	// Test writeToZip with io.Copy error on stream
	{
		f := NewFile()
		f.streams = make(map[string]*StreamWriter)
		sw := &StreamWriter{}
		sw.rawData.WriteString("test data")
		f.streams["s"] = sw
		f.SetZipWriter(func(w io.Writer) ZipWriter {
			return &errZipWriter{
				createFunc: func(name string) (io.Writer, error) {
					if name == "s" {
						return &errWriter{err: errors.New("copy error")}, nil
					}
					return &bytes.Buffer{}, nil
				},
			}
		})
		_, err := f.WriteTo(bufio.NewWriter(&bytes.Buffer{}))
		assert.EqualError(t, err, "copy error")
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
	assert.NoError(t, binary.Write(buf, binary.LittleEndian, uint16(10)))
	buf.Write(make([]byte, 2))
	buf.WriteString("test")
	assert.NoError(t, f.writeZip64LFH(buf))

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
}

func TestRemoveTempFiles(t *testing.T) {
	tmp, err := os.CreateTemp("", "excelize-*")
	if err != nil {
		t.Fatal(err)
	}
	tmpName := tmp.Name()
	assert.NoError(t, tmp.Close())
	f := NewFile()
	// Fill the tempFiles map with non-existing files
	for i := range 1000 {
		f.tempFiles.Store(strconv.Itoa(i), "/hopefully not existing")
	}
	f.tempFiles.Store("existing", tmpName)

	require.Error(t, f.Close())
	if _, err := os.Stat(tmpName); err == nil {
		t.Errorf("temp file %q still exist", tmpName)
		assert.NoError(t, os.Remove(tmpName))
	}
}

func TestStreamingWriteTo(t *testing.T) {
	// Verify that WriteTo streams directly when no password is set,
	// producing a valid XLSX file.
	f := NewFile()
	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	for row := 1; row <= 100; row++ {
		rowData := make([]interface{}, 10)
		for col := range 10 {
			rowData[col] = "test"
		}
		cell, _ := CoordinatesToCellName(1, row)
		assert.NoError(t, sw.SetRow(cell, rowData))
	}
	assert.NoError(t, sw.Flush())
	// WriteTo a buffer (exercises the streaming path, no password)
	var buf bytes.Buffer
	_, err = f.WriteTo(&buf)
	assert.NoError(t, err)
	assert.Greater(t, buf.Len(), 0)
	assert.NoError(t, f.Close())
	// Verify the output is a valid ZIP/XLSX by reading it back
	f2, err := OpenReader(bytes.NewReader(buf.Bytes()))
	assert.NoError(t, err)
	val, err := f2.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "test", val)
	assert.NoError(t, f2.Close())
}

func TestCompressionOption(t *testing.T) {
	// Generate a file with known content, then save with different
	// compression levels and compare sizes.
	makeFile := func() *File {
		f := NewFile()
		sw, _ := f.NewStreamWriter("Sheet1")
		for row := 1; row <= 500; row++ {
			rowData := make([]interface{}, 20)
			for col := range 20 {
				rowData[col] = "Hello World"
			}
			cell, _ := CoordinatesToCellName(1, row)
			_ = sw.SetRow(cell, rowData)
		}
		_ = sw.Flush()
		return f
	}

	// Default compression
	f1 := makeFile()
	var buf1 bytes.Buffer
	_, err := f1.WriteTo(&buf1)
	assert.NoError(t, err)
	assert.NoError(t, f1.Close())

	// No compression
	f2 := makeFile()
	f2.options.Compression = CompressionNone
	var buf2 bytes.Buffer
	_, err = f2.WriteTo(&buf2)
	assert.NoError(t, err)
	assert.NoError(t, f2.Close())

	// Best speed
	f3 := makeFile()
	f3.options.Compression = CompressionBestSpeed
	var buf3 bytes.Buffer
	_, err = f3.WriteTo(&buf3)
	assert.NoError(t, err)
	assert.NoError(t, f3.Close())

	// No compression should produce the largest file
	assert.Greater(t, buf2.Len(), buf1.Len(), "uncompressed should be larger than default")
	assert.Greater(t, buf2.Len(), buf3.Len(), "uncompressed should be larger than best-speed")

	// All should be valid XLSX files
	for _, buf := range []*bytes.Buffer{&buf1, &buf2, &buf3} {
		f, err := OpenReader(bytes.NewReader(buf.Bytes()))
		assert.NoError(t, err)
		val, err := f.GetCellValue("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Equal(t, "Hello World", val)
		assert.NoError(t, f.Close())
	}
}

func TestWriteToBufferCompression(t *testing.T) {
	// Ensure WriteToBuffer also respects the Compression option
	f := NewFile(Options{Compression: CompressionNone})
	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	for row := 1; row <= 100; row++ {
		cell, _ := CoordinatesToCellName(1, row)
		_ = sw.SetRow(cell, []interface{}{"data"})
	}
	_ = sw.Flush()
	buf, err := f.WriteToBuffer()
	assert.NoError(t, err)
	assert.Greater(t, buf.Len(), 0)
	assert.NoError(t, f.Close())
}

func TestWriteToWithPassword(t *testing.T) {
	// Test that WriteTo with password uses temp file approach
	f := NewFile()
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "Encrypted Data"))

	// Write with password encryption to a buffer
	var buf bytes.Buffer
	n, err := f.WriteTo(&buf, Options{Password: "testpass"})
	assert.NoError(t, err)
	assert.Greater(t, n, int64(0))
	assert.Greater(t, buf.Len(), 0)

	// Verify it can be opened with the password
	encrypted := buf.Bytes()
	f2, err := OpenReader(bytes.NewReader(encrypted), Options{Password: "testpass"})
	assert.NoError(t, err)
	val, err := f2.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "Encrypted Data", val)
	assert.NoError(t, f2.Close())
	assert.NoError(t, f.Close())
}

func TestWriteToWithPasswordAndCompression(t *testing.T) {
	// Test that compression settings work with password encryption
	f := NewFile(Options{Compression: CompressionBestSpeed})
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "Test"))

	var buf bytes.Buffer
	_, err := f.WriteTo(&buf, Options{Password: "pass", Compression: CompressionBestSpeed})
	assert.NoError(t, err)
	assert.Greater(t, buf.Len(), 0)

	// Verify it opens correctly
	f2, err := OpenReader(bytes.NewReader(buf.Bytes()), Options{Password: "pass"})
	assert.NoError(t, err)
	assert.NoError(t, f2.Close())
	assert.NoError(t, f.Close())
}
