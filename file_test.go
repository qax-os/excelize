package excelize

import (
	"archive/zip"
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
		sw.rawData.buf.WriteString("test data")
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

func TestWriteDirect(t *testing.T) {
	// Test stream written worksheet round-trip through a plain io.Writer
	f := NewFile()
	sw, err := f.NewStreamWriter("Sheet1")
	assert.NoError(t, err)
	for r := 1; r <= 100; r++ {
		cell, err := CoordinatesToCellName(1, r)
		assert.NoError(t, err)
		assert.NoError(t, sw.SetRow(cell, []interface{}{r, "value-" + strconv.Itoa(r)}))
	}
	assert.NoError(t, sw.Flush())
	buf := new(bytes.Buffer)
	written, err := f.WriteTo(buf)
	assert.NoError(t, err)
	assert.Equal(t, int64(buf.Len()), written)
	assert.NoError(t, f.Close())

	result, err := OpenReader(bytes.NewReader(buf.Bytes()))
	assert.NoError(t, err)
	val, err := result.GetCellValue("Sheet1", "B100")
	assert.NoError(t, err)
	assert.Equal(t, "value-100", val)
	assert.NoError(t, result.Close())

	// Test WriteTo with password still returns an encrypted workbook
	f = NewFile()
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 1))
	buf.Reset()
	_, err = f.WriteTo(buf, Options{Password: "password"})
	assert.NoError(t, err)
	assert.NoError(t, f.Close())
	result, err = OpenReader(bytes.NewReader(buf.Bytes()), Options{Password: "password"})
	assert.NoError(t, err)
	val, err = result.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "1", val)
	assert.NoError(t, result.Close())

	// Test write with part in temporary file streamed from disk
	f = NewFile()
	content := []byte("<custom>data</custom>")
	tmp, err := os.CreateTemp(t.TempDir(), "excelize-")
	assert.NoError(t, err)
	_, err = tmp.Write(content)
	assert.NoError(t, err)
	assert.NoError(t, tmp.Close())
	f.tempFiles.Store("xl/custom.xml", tmp.Name())
	// Test write with unreadable temporary file part returns the read error
	f.tempFiles.Store("xl/missing.xml", filepath.Join(t.TempDir(), "missing"))
	buf.Reset()
	_, err = f.WriteTo(buf)
	assert.Error(t, err)
	f.tempFiles.Delete("xl/missing.xml")
	buf.Reset()
	_, err = f.WriteTo(buf)
	assert.NoError(t, err)
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	assert.NoError(t, err)
	entry, err := zr.Open("xl/custom.xml")
	assert.NoError(t, err)
	entryData, err := io.ReadAll(entry)
	assert.NoError(t, err)
	assert.Equal(t, content, entryData)
	assert.NoError(t, entry.Close())
	// The part should be streamed from disk, not cached into memory
	_, ok := f.Pkg.Load("xl/custom.xml")
	assert.False(t, ok)
	assert.NoError(t, f.Close())
}

func TestWriteSpooled(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "spooled"))
	// Test with stale ZIP64 entries of the previous save are reset
	f.zip64Entries = []string{"stale"}
	f.prepareToWrite()
	buf := new(bytes.Buffer)
	written, err := f.writeSpooled(buf)
	assert.NoError(t, err)
	assert.Equal(t, int64(buf.Len()), written)
	assert.Empty(t, f.zip64Entries)
	result, err := OpenReader(bytes.NewReader(buf.Bytes()))
	assert.NoError(t, err)
	val, err := result.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "spooled", val)
	assert.NoError(t, result.Close())
	assert.NoError(t, f.Close())

	// Test with unavailable temporary directory
	f = NewFile(Options{TmpDir: filepath.Join(t.TempDir(), "unavailable")})
	_, err = f.writeSpooled(new(bytes.Buffer))
	assert.Error(t, err)
	assert.NoError(t, f.Close())
}

func TestZip64Needed(t *testing.T) {
	f := NewFile()
	assert.False(t, f.zip64Needed())

	// Test with stream data size over the ZIP64 size threshold
	tmp, err := os.CreateTemp(t.TempDir(), "excelize-")
	assert.NoError(t, err)
	assert.NoError(t, tmp.Truncate(math.MaxUint32+1))
	f.streams = map[string]*StreamWriter{"xl/worksheets/sheet1.xml": {rawData: bufferedWriter{tmp: tmp}}}
	assert.True(t, f.zip64Needed())
	f.streams = nil
	assert.NoError(t, tmp.Close())

	// Test with temporary file part size over the ZIP64 size threshold, the
	// part name must not be in the Pkg, parts in the Pkg take precedence over
	// temporary files
	tmp, err = os.CreateTemp(t.TempDir(), "excelize-")
	assert.NoError(t, err)
	assert.NoError(t, tmp.Truncate(math.MaxUint32+1))
	assert.NoError(t, tmp.Close())
	f.tempFiles.Store("xl/worksheets/sheet2.xml", tmp.Name())
	assert.True(t, f.zip64Needed())
	f.tempFiles.Delete("xl/worksheets/sheet2.xml")

	// Test with in memory part size over the ZIP64 size threshold
	f.Pkg.Store("xl/worksheets/sheet1.xml", make([]byte, math.MaxUint32+1))
	assert.True(t, f.zip64Needed())
	f.Pkg.Delete("xl/worksheets/sheet1.xml")
	assert.NoError(t, f.Close())
}

func TestPatchZip64LFH(t *testing.T) {
	// Test with no ZIP64 entries is a no-op
	f := NewFile()
	assert.NoError(t, f.patchZip64LFH(nil, 0))

	lfh := func(name string) []byte {
		b := make([]byte, 30)
		copy(b, []byte{0x50, 0x4b, 0x03, 0x04})
		binary.LittleEndian.PutUint16(b[4:6], 20)
		binary.LittleEndian.PutUint16(b[26:28], uint16(len(name)))
		return append(b, name...)
	}
	const chunkSize = 1 << 20
	f.zip64Entries = append(f.zip64Entries, defaultXMLPathSharedStrings)
	// The first header will be patched, the second straddles the chunk
	// scanning boundary and doesn't match the ZIP64 entries, the third is
	// beyond the boundary and will be patched by the next chunk scan
	data := lfh(defaultXMLPathSharedStrings)
	data = append(data, make([]byte, chunkSize-2-len(data))...)
	secondOffset := len(data)
	data = append(data, lfh("xl/media/image1.png")...)
	thirdOffset := len(data)
	data = append(data, lfh(defaultXMLPathSharedStrings)...)

	tmp, err := os.CreateTemp(t.TempDir(), "excelize-")
	assert.NoError(t, err)
	_, err = tmp.Write(data)
	assert.NoError(t, err)
	assert.NoError(t, f.patchZip64LFH(tmp, int64(len(data))))
	patched, err := os.ReadFile(tmp.Name())
	assert.NoError(t, err)
	assert.Equal(t, uint16(45), binary.LittleEndian.Uint16(patched[4:6]))
	assert.Equal(t, uint16(20), binary.LittleEndian.Uint16(patched[secondOffset+4:secondOffset+6]))
	assert.Equal(t, uint16(45), binary.LittleEndian.Uint16(patched[thirdOffset+4:thirdOffset+6]))
	assert.NoError(t, tmp.Close())

	// Test with file header less than the required 30 for the fixed header part
	tmp, err = os.CreateTemp(t.TempDir(), "excelize-")
	assert.NoError(t, err)
	truncated := append([]byte{0x50, 0x4b, 0x03, 0x04}, make([]byte, 20)...)
	_, err = tmp.Write(truncated)
	assert.NoError(t, err)
	assert.NoError(t, f.patchZip64LFH(tmp, int64(len(truncated))))
	assert.NoError(t, tmp.Close())

	// Test with filename length overflow the file size
	tmp, err = os.CreateTemp(t.TempDir(), "excelize-")
	assert.NoError(t, err)
	overflow := lfh(strings.Repeat("s", 100))[:40]
	_, err = tmp.Write(overflow)
	assert.NoError(t, err)
	assert.NoError(t, f.patchZip64LFH(tmp, int64(len(overflow))))
	assert.NoError(t, tmp.Close())
	assert.NoError(t, f.Close())
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
