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
// return errors. When createErr is nil it delegates Create to a real zip.Writer.
// If createWriter is set, Create returns that writer instead of calling inner.
type errZipWriter struct {
	inner        *zip.Writer
	createErr    error
	closeErr     error
	createWriter io.Writer
	onClose      func()
}

func (zw *errZipWriter) Create(name string) (io.Writer, error) {
	if zw.createErr != nil {
		return nil, zw.createErr
	}
	if zw.createWriter != nil {
		return zw.createWriter, nil
	}
	return zw.inner.Create(name)
}

func (zw *errZipWriter) AddFS(fsys fs.FS) error { return nil }

func (zw *errZipWriter) Close() error {
	if zw.closeErr != nil {
		return zw.closeErr
	}
	err := zw.inner.Close()
	if err == nil && zw.onClose != nil {
		zw.onClose()
	}
	return err
}

// limitedWriter returns an error after n bytes have been written.
type limitedWriter struct {
	w io.Writer
	n int
}

func (lw *limitedWriter) Write(p []byte) (int, error) {
	if lw.n <= 0 {
		return 0, errors.New("write limit exceeded")
	}
	if len(p) > lw.n {
		p = p[:lw.n]
	}
	n, err := lw.w.Write(p)
	lw.n -= n
	return n, err
}

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

	// Test WriteToBuffer with writeToZip error
	{
		f := NewFile()
		f.SetZipWriter(func(w io.Writer) ZipWriter {
			return &errZipWriter{inner: zip.NewWriter(w), createErr: errors.New("create error")}
		})
		_, err := f.WriteToBuffer()
		assert.EqualError(t, err, "create error")
	}
	// Test WriteToBuffer with zw.Close() error
	{
		f := NewFile()
		f.SetZipWriter(func(w io.Writer) ZipWriter {
			return &errZipWriter{inner: zip.NewWriter(w), closeErr: errors.New("close error")}
		})
		_, err := f.WriteToBuffer()
		assert.EqualError(t, err, "close error")
	}
	// Test writeDirectToWriter with invalid TmpDir (CreateTemp error)
	{
		f := NewFile()
		f.options = &Options{TmpDir: filepath.Join(os.TempDir(), "nonexistent-excelize-dir")}
		_, err := f.writeDirectToWriter(io.Discard)
		assert.Error(t, err)
	}
	// Test writeDirectToWriter with zw.Close() error
	{
		f := NewFile()
		f.SetZipWriter(func(w io.Writer) ZipWriter {
			return &errZipWriter{inner: zip.NewWriter(w), closeErr: errors.New("close error")}
		})
		_, err := f.writeDirectToWriter(io.Discard)
		assert.EqualError(t, err, "close error")
	}
	// Test writeDirectToWriter with writeToFile error
	{
		f := NewFile()
		f.SetZipWriter(func(w io.Writer) ZipWriter {
			return &errZipWriter{
				inner: zip.NewWriter(w),
				onClose: func() {
					f.zip64Entries = append(f.zip64Entries, defaultXMLPathSharedStrings)
					if file, ok := w.(*os.File); ok {
						file.Close()
					}
				},
			}
		})
		_, err := f.writeDirectToWriter(io.Discard)
		assert.Error(t, err)
	}
	// Test writeDirectToWriter with Seek error
	{
		f := NewFile()
		f.SetZipWriter(func(w io.Writer) ZipWriter {
			return &errZipWriter{
				inner: zip.NewWriter(w),
				onClose: func() {
					if file, ok := w.(*os.File); ok {
						file.Close()
					}
				},
			}
		})
		_, err := f.writeDirectToWriter(io.Discard)
		assert.Error(t, err)
	}
	// Test writeToZip with stream Create error
	{
		f := NewFile()
		f.streams = map[string]*StreamWriter{"s": {rawData: bufferedWriter{}}}
		zw := &errZipWriter{inner: zip.NewWriter(&bytes.Buffer{}), createErr: errors.New("stream create error")}
		assert.EqualError(t, f.writeToZip(zw), "stream create error")
	}
	// Test writeToZip with stream rawData.Reader() error
	{
		f := NewFile()
		tmp, err := os.CreateTemp(os.TempDir(), "excelize-")
		assert.NoError(t, err)
		name := tmp.Name()
		tmp.Close()
		os.Remove(name)
		closedFile, _ := os.Open(name)
		// Open failed, so create a real temp then close+remove to get a closed *os.File
		tmp2, err := os.CreateTemp(os.TempDir(), "excelize-")
		assert.NoError(t, err)
		_ = closedFile
		tmp2.Close()
		os.Remove(tmp2.Name())
		f.streams = map[string]*StreamWriter{
			"s": {rawData: bufferedWriter{tmp: tmp2}},
		}
		zw := zip.NewWriter(&bytes.Buffer{})
		assert.Error(t, f.writeToZip(zw))
	}
	// Test writeToZip with stream io.Copy error
	{
		f := NewFile()
		bw := bufferedWriter{}
		_, err := bw.WriteString("data")
		assert.NoError(t, err)
		f.streams = map[string]*StreamWriter{"s": {rawData: bw}}
		f.Pkg = sync.Map{}
		zw := &errZipWriter{
			inner:        zip.NewWriter(&bytes.Buffer{}),
			createWriter: &limitedWriter{w: io.Discard, n: 0},
		}
		assert.Error(t, f.writeToZip(zw))
	}
	// Test writeToFile with closed file (Stat error)
	{
		f := NewFile()
		f.zip64Entries = append(f.zip64Entries, defaultXMLPathSharedStrings)
		tmp, err := os.CreateTemp(os.TempDir(), "excelize-")
		assert.NoError(t, err)
		os.Remove(tmp.Name())
		tmp.Close()
		assert.Error(t, f.writeToFile(tmp))
	}
	// Test writeToFile with write-only file (ReadAt error)
	{
		f := NewFile()
		f.zip64Entries = append(f.zip64Entries, defaultXMLPathSharedStrings)
		tmp, err := os.CreateTemp(os.TempDir(), "excelize-")
		assert.NoError(t, err)
		name := tmp.Name()
		_, err = tmp.Write(make([]byte, 30))
		assert.NoError(t, err)
		tmp.Close()
		wo, err := os.OpenFile(name, os.O_WRONLY, 0)
		assert.NoError(t, err)
		assert.Error(t, f.writeToFile(wo))
		wo.Close()
		os.Remove(name)
	}
	// Test writeToFile with file too small for LFH header
	{
		f := NewFile()
		f.zip64Entries = append(f.zip64Entries, defaultXMLPathSharedStrings)
		tmp, err := os.CreateTemp(os.TempDir(), "excelize-")
		assert.NoError(t, err)
		data := make([]byte, 29)
		copy(data, []byte{0x50, 0x4b, 0x03, 0x04})
		_, err = tmp.Write(data)
		assert.NoError(t, err)
		assert.NoError(t, f.writeToFile(tmp))
		tmp.Close()
		os.Remove(tmp.Name())
	}
	// Test writeToFile with filenameLen extending past EOF
	{
		f := NewFile()
		f.zip64Entries = append(f.zip64Entries, defaultXMLPathSharedStrings)
		tmp, err := os.CreateTemp(os.TempDir(), "excelize-")
		assert.NoError(t, err)
		data := make([]byte, 30)
		copy(data[:4], []byte{0x50, 0x4b, 0x03, 0x04})
		binary.LittleEndian.PutUint16(data[26:28], 10)
		_, err = tmp.Write(data)
		assert.NoError(t, err)
		assert.NoError(t, f.writeToFile(tmp))
		tmp.Close()
		os.Remove(tmp.Name())
	}
	// Test writeToFile with concurrent file truncation (ReadAt fixed header error)
	{
		f := NewFile()
		f.zip64Entries = append(f.zip64Entries, defaultXMLPathSharedStrings)
		tmp, err := os.CreateTemp(os.TempDir(), "excelize-")
		assert.NoError(t, err)
		data := make([]byte, 50)
		copy(data[:4], []byte{0x50, 0x4b, 0x03, 0x04})
		binary.LittleEndian.PutUint16(data[26:28], 4)
		copy(data[30:34], []byte("test"))
		_, err = tmp.Write(data)
		assert.NoError(t, err)
		done := make(chan struct{})
		go func() {
			for {
				select {
				case <-done:
					return
				default:
					_ = tmp.Truncate(10)
				}
			}
		}()
		for range 2000 {
			_, _ = tmp.WriteAt(data, 0)
			_ = f.writeToFile(tmp)
		}
		close(done)
		tmp.Close()
		os.Remove(tmp.Name())
	}
	// Test writeToFile with concurrent file truncation (ReadAt filename error)
	{
		f := NewFile()
		f.zip64Entries = append(f.zip64Entries, defaultXMLPathSharedStrings)
		tmp, err := os.CreateTemp(os.TempDir(), "excelize-")
		assert.NoError(t, err)
		data := make([]byte, 100000)
		copy(data[90000:90004], []byte{0x50, 0x4b, 0x03, 0x04})
		binary.LittleEndian.PutUint16(data[90026:90028], 10)
		_, err = tmp.Write(data)
		assert.NoError(t, err)
		done := make(chan struct{})
		go func() {
			for {
				select {
				case <-done:
					return
				default:
					_ = tmp.Truncate(90035)
				}
			}
		}()
		for range 100 {
			_ = tmp.Truncate(100000)
			_ = f.writeToFile(tmp)
		}
		close(done)
		tmp.Close()
		os.Remove(tmp.Name())
	}
	// Test writeToFile with read-only file (WriteAt error)
	{
		f := NewFile()
		entryName := defaultXMLPathSharedStrings
		f.zip64Entries = append(f.zip64Entries, entryName)
		tmp, err := os.CreateTemp(os.TempDir(), "excelize-")
		assert.NoError(t, err)
		name := tmp.Name()
		header := make([]byte, 30)
		copy(header[:4], []byte{0x50, 0x4b, 0x03, 0x04})
		binary.LittleEndian.PutUint16(header[26:28], uint16(len(entryName)))
		_, err = tmp.Write(header)
		assert.NoError(t, err)
		_, err = tmp.WriteString(entryName)
		assert.NoError(t, err)
		tmp.Close()
		ro, err := os.Open(name)
		assert.NoError(t, err)
		assert.Error(t, f.writeToFile(ro))
		ro.Close()
		os.Remove(name)
	}
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
	for i := range 1000 {
		f.tempFiles.Store(strconv.Itoa(i), "/hopefully not existing")
	}
	f.tempFiles.Store("existing", tmpName)

	require.Error(t, f.Close())
	if _, err := os.Stat(tmpName); err == nil {
		t.Errorf("temp file %q still exist", tmpName)
		os.Remove(tmpName)
	}
}
