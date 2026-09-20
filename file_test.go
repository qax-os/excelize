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

func TestCompressionOption(t *testing.T) {
	// Test CompressionNone produces valid but larger output
	f := NewFile()
	f.options.Compression = CompressionNone
	f.SetCellValue("Sheet1", "A1", "hello")
	bufNone, err := f.WriteToBuffer()
	assert.NoError(t, err)
	f.Close()

	// Test CompressionBestSpeed produces valid output
	f2 := NewFile()
	f2.options.Compression = CompressionBestSpeed
	f2.SetCellValue("Sheet1", "A1", "hello")
	bufFast, err := f2.WriteToBuffer()
	assert.NoError(t, err)
	f2.Close()

	// Test CompressionDefault (baseline)
	f3 := NewFile()
	f3.SetCellValue("Sheet1", "A1", "hello")
	bufDefault, err := f3.WriteToBuffer()
	assert.NoError(t, err)
	f3.Close()

	// Uncompressed should be larger than default
	assert.Greater(t, bufNone.Len(), bufDefault.Len())
	// BestSpeed should be between the two (or equal to default for small files)
	assert.LessOrEqual(t, bufFast.Len(), bufNone.Len())

	// Verify all outputs are valid ZIP files that can be reopened
	for _, buf := range []*bytes.Buffer{bufNone, bufFast, bufDefault} {
		f4, err := OpenReader(buf)
		assert.NoError(t, err)
		val, err := f4.GetCellValue("Sheet1", "A1")
		assert.NoError(t, err)
		assert.Equal(t, "hello", val)
		f4.Close()
	}
}

func TestConfigureZipCompressionCustomWriter(t *testing.T) {
	// configureZipCompression is a no-op for non-*zip.Writer implementations
	f := NewFile()
	defer f.Close()
	f.options.Compression = CompressionNone
	// Directly call configureZipCompression with a non-*zip.Writer
	f.configureZipCompression(&errZipWriter{})
	// Also test unknown compression value
	f.options.Compression = Compression(99)
	f.configureZipCompression(zip.NewWriter(io.Discard))
}

func TestWriteToWithEncryption(t *testing.T) {
	// Test WriteTo with password triggers writeToWithEncryption
	f := NewFile()
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "encrypted"))
	var buf bytes.Buffer
	n, err := f.WriteTo(&buf, Options{Password: "test123"})
	assert.NoError(t, err)
	assert.Greater(t, n, int64(0))
	assert.Equal(t, int64(buf.Len()), n)
	assert.NoError(t, f.Close())

	// Verify the encrypted file can be opened
	f2, err := OpenReader(&buf, Options{Password: "test123"})
	assert.NoError(t, err)
	val, err := f2.GetCellValue("Sheet1", "A1")
	assert.NoError(t, err)
	assert.Equal(t, "encrypted", val)
	assert.NoError(t, f2.Close())
}

func TestWriteToWithEncryptionTmpDir(t *testing.T) {
	// Test that TmpDir option is respected
	tmpDir := t.TempDir()
	f := NewFile()
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "data"))
	var buf bytes.Buffer
	_, err := f.WriteTo(&buf, Options{Password: "pass", TmpDir: tmpDir})
	assert.NoError(t, err)
	assert.NoError(t, f.Close())
}

func TestWriteToWithEncryptionCreateTempError(t *testing.T) {
	// Test writeToWithEncryption with invalid TmpDir causes CreateTemp error
	f := NewFile()
	var buf bytes.Buffer
	_, err := f.WriteTo(&buf, Options{
		Password: "test",
		TmpDir:   "/nonexistent/path/that/should/not/exist",
	})
	assert.Error(t, err)
	assert.NoError(t, f.Close())
}

func TestWriteToWithEncryptionWriteToZipError(t *testing.T) {
	// Test writeToWithEncryption when writeToZip fails
	f := NewFile()
	f.SetZipWriter(func(w io.Writer) ZipWriter {
		return &errZipWriter{
			createFunc: func(name string) (io.Writer, error) {
				return nil, errors.New("create error in encryption path")
			},
		}
	})
	var buf bytes.Buffer
	_, err := f.WriteTo(&buf, Options{Password: "test"})
	assert.EqualError(t, err, "create error in encryption path")
	assert.NoError(t, f.Close())
}

func TestWriteToWithEncryptionCloseError(t *testing.T) {
	// Test writeToWithEncryption when zw.Close fails
	f := NewFile()
	f.SetZipWriter(func(w io.Writer) ZipWriter {
		return &errZipWriter{
			createFunc: func(name string) (io.Writer, error) {
				return &bytes.Buffer{}, nil
			},
			closeErr: errors.New("close error in encryption path"),
		}
	})
	var buf bytes.Buffer
	_, err := f.WriteTo(&buf, Options{Password: "test"})
	assert.EqualError(t, err, "close error in encryption path")
	assert.NoError(t, f.Close())
}

func TestWriteZip64LFHFile(t *testing.T) {
	// Test writeZip64LFHFile with no zip64 entries (no-op)
	f := NewFile()
	tmp, err := os.CreateTemp("", "excelize-test-lfh-*")
	assert.NoError(t, err)
	defer func() {
		tmp.Close()
		os.Remove(tmp.Name())
	}()
	f.zip64Entries = nil
	assert.NoError(t, f.writeZip64LFHFile(tmp))

	// Test writeZip64LFHFile with a valid ZIP local file header
	f.zip64Entries = []string{"test.xml"}
	// Write a fake local file header: PK\x03\x04 + 22 bytes header + 2-byte filename len + 2 bytes extra len + filename
	var hdr bytes.Buffer
	hdr.Write([]byte{0x50, 0x4b, 0x03, 0x04})          // signature
	hdr.Write(make([]byte, 22))                        // version through extra field len offset
	binary.Write(&hdr, binary.LittleEndian, uint16(8)) // filename length = 8
	hdr.Write(make([]byte, 2))                         // extra field length
	hdr.WriteString("test.xml")                        // filename
	_, err = tmp.Seek(0, 0)
	assert.NoError(t, err)
	assert.NoError(t, tmp.Truncate(0))
	_, err = tmp.Write(hdr.Bytes())
	assert.NoError(t, err)
	assert.NoError(t, f.writeZip64LFHFile(tmp))

	// Verify the version was updated to 45
	vBuf := make([]byte, 2)
	_, err = tmp.ReadAt(vBuf, 4)
	assert.NoError(t, err)
	assert.Equal(t, uint16(45), binary.LittleEndian.Uint16(vBuf))

	// Verify compressed size set to 0xFFFFFFFF
	sBuf := make([]byte, 4)
	_, err = tmp.ReadAt(sBuf, 18)
	assert.NoError(t, err)
	assert.Equal(t, uint32(0xFFFFFFFF), binary.LittleEndian.Uint32(sBuf))

	// Verify uncompressed size set to 0xFFFFFFFF
	_, err = tmp.ReadAt(sBuf, 22)
	assert.NoError(t, err)
	assert.Equal(t, uint32(0xFFFFFFFF), binary.LittleEndian.Uint32(sBuf))
}

func TestWriteZip64LFHFileSeekError(t *testing.T) {
	f := NewFile()
	f.zip64Entries = []string{"test.xml"}
	tmp, err := os.CreateTemp("", "excelize-test-lfh-seek-*")
	assert.NoError(t, err)
	tmp.Close() // close so Seek fails
	os.Remove(tmp.Name())
	err = f.writeZip64LFHFile(tmp)
	assert.Error(t, err)
}

func TestWriteZip64LFHFileTruncatedHeader(t *testing.T) {
	// Test with header too short (less than 30 bytes after signature)
	f := NewFile()
	f.zip64Entries = []string{"test.xml"}
	tmp, err := os.CreateTemp("", "excelize-test-lfh-trunc-*")
	assert.NoError(t, err)
	defer func() {
		tmp.Close()
		os.Remove(tmp.Name())
	}()
	// Write just the signature + a few bytes (not enough for full header)
	tmp.Write([]byte{0x50, 0x4b, 0x03, 0x04})
	tmp.Write(make([]byte, 10))
	assert.NoError(t, f.writeZip64LFHFile(tmp))
}

func TestWriteZip64LFHFileNonMatchingEntry(t *testing.T) {
	// Test with a valid header but filename doesn't match zip64Entries
	f := NewFile()
	f.zip64Entries = []string{"other.xml"}
	tmp, err := os.CreateTemp("", "excelize-test-lfh-nomatch-*")
	assert.NoError(t, err)
	defer func() {
		tmp.Close()
		os.Remove(tmp.Name())
	}()
	var hdr bytes.Buffer
	hdr.Write([]byte{0x50, 0x4b, 0x03, 0x04})
	hdr.Write(make([]byte, 22))
	binary.Write(&hdr, binary.LittleEndian, uint16(8))
	hdr.Write(make([]byte, 2))
	hdr.WriteString("test.xml")
	tmp.Write(hdr.Bytes())
	assert.NoError(t, f.writeZip64LFHFile(tmp))

	// Version should NOT be changed (still 0)
	vBuf := make([]byte, 2)
	_, err = tmp.ReadAt(vBuf, 4)
	assert.NoError(t, err)
	assert.Equal(t, uint16(0), binary.LittleEndian.Uint16(vBuf))
}

func TestWriteZip64LFHFileMultipleHeaders(t *testing.T) {
	// Test with multiple local file headers, some matching and some not
	f := NewFile()
	f.zip64Entries = []string{"match.xml"}
	tmp, err := os.CreateTemp("", "excelize-test-lfh-multi-*")
	assert.NoError(t, err)
	defer func() {
		tmp.Close()
		os.Remove(tmp.Name())
	}()

	// Write first header (non-matching)
	var hdr1 bytes.Buffer
	hdr1.Write([]byte{0x50, 0x4b, 0x03, 0x04}) // signature
	hdr1.Write(make([]byte, 22))               // filler
	binary.Write(&hdr1, binary.LittleEndian, uint16(10))
	hdr1.Write(make([]byte, 2))
	hdr1.WriteString("other.xml\x00") // 10 bytes
	hdr1.Write(make([]byte, 50))      // fake file data

	// Write second header (matching)
	var hdr2 bytes.Buffer
	hdr2.Write([]byte{0x50, 0x4b, 0x03, 0x04}) // signature
	hdr2.Write(make([]byte, 22))               // filler
	binary.Write(&hdr2, binary.LittleEndian, uint16(9))
	hdr2.Write(make([]byte, 2))
	hdr2.WriteString("match.xml") // 9 bytes

	tmp.Write(hdr1.Bytes())
	tmp.Write(hdr2.Bytes())
	assert.NoError(t, f.writeZip64LFHFile(tmp))

	// First header should NOT be modified
	vBuf := make([]byte, 2)
	_, err = tmp.ReadAt(vBuf, 4)
	assert.NoError(t, err)
	assert.Equal(t, uint16(0), binary.LittleEndian.Uint16(vBuf))

	// Second header should have version 45
	offset2 := int64(hdr1.Len())
	_, err = tmp.ReadAt(vBuf, offset2+4)
	assert.NoError(t, err)
	assert.Equal(t, uint16(45), binary.LittleEndian.Uint16(vBuf))
}

func TestWriteZip64LFHFileFilenameOverflow(t *testing.T) {
	// Test with filename length that extends beyond the read buffer
	f := NewFile()
	f.zip64Entries = []string{"test.xml"}
	tmp, err := os.CreateTemp("", "excelize-test-lfh-overflow-*")
	assert.NoError(t, err)
	defer func() {
		tmp.Close()
		os.Remove(tmp.Name())
	}()
	// Write header where filename length claims more bytes than available
	var hdr bytes.Buffer
	hdr.Write([]byte{0x50, 0x4b, 0x03, 0x04})
	hdr.Write(make([]byte, 22))
	binary.Write(&hdr, binary.LittleEndian, uint16(9999)) // very long filename
	hdr.Write(make([]byte, 2))
	hdr.WriteString("short") // actual data is short
	tmp.Write(hdr.Bytes())
	assert.NoError(t, f.writeZip64LFHFile(tmp))
}

// bigCountWriter wraps an io.Writer and lies about the number of bytes
// written, reporting math.MaxUint32+1 to trigger zip64 detection in writeToZip.
type bigCountWriter struct{ w io.Writer }

func (b *bigCountWriter) Write(p []byte) (int, error) {
	_, err := b.w.Write(p)
	if err != nil {
		return 0, err
	}
	return math.MaxUint32 + 1, nil
}

// zip64TriggerZipWriter wraps a real zip.Writer and makes every Create call
// return a bigCountWriter so that all entries appear to exceed 4GB.
type zip64TriggerZipWriter struct {
	real *zip.Writer
}

func (z *zip64TriggerZipWriter) Create(name string) (io.Writer, error) {
	w, err := z.real.Create(name)
	if err != nil {
		return w, err
	}
	return &bigCountWriter{w: w}, nil
}

func (z *zip64TriggerZipWriter) AddFS(fsys fs.FS) error { return z.real.AddFS(fsys) }
func (z *zip64TriggerZipWriter) Close() error           { return z.real.Close() }

// fileCloseZipWriter wraps a real zip.Writer and closes the underlying
// *os.File during Close, causing subsequent operations on the file to fail.
// When triggerZip64 is true, it also uses bigCountWriter to populate
// zip64Entries without allocating 4GB+ of memory.
type fileCloseZipWriter struct {
	real         *zip.Writer
	file         *os.File
	triggerZip64 bool
}

func (z *fileCloseZipWriter) Create(name string) (io.Writer, error) {
	w, err := z.real.Create(name)
	if err != nil || !z.triggerZip64 {
		return w, err
	}
	return &bigCountWriter{w: w}, nil
}

func (z *fileCloseZipWriter) AddFS(fsys fs.FS) error { return z.real.AddFS(fsys) }

func (z *fileCloseZipWriter) Close() error {
	if err := z.real.Close(); err != nil {
		return err
	}
	return z.file.Close()
}
func TestWriteToWithEncryptionZip64(t *testing.T) {
	// Use a custom ZipWriter that fakes >4GB writes to trigger the zip64 LFH
	// fixup path inside writeToWithEncryption without allocating 4GB+ of memory.
	f := NewFile()
	f.SetZipWriter(func(w io.Writer) ZipWriter {
		return &zip64TriggerZipWriter{real: zip.NewWriter(w)}
	})
	var buf bytes.Buffer
	_, err := f.WriteTo(&buf, Options{Password: "test"})
	assert.NoError(t, err)
	assert.NoError(t, f.Close())
}

func TestWriteToWithEncryptionZip64Error(t *testing.T) {
	// Close the underlying file during ZipWriter.Close so that
	// writeZip64LFHFile fails, exercising the error return at line 180-182.
	f := NewFile()
	f.SetZipWriter(func(w io.Writer) ZipWriter {
		return &fileCloseZipWriter{
			real:         zip.NewWriter(w),
			file:         w.(*os.File),
			triggerZip64: true,
		}
	})
	var buf bytes.Buffer
	_, err := f.WriteTo(&buf, Options{Password: "test"})
	assert.Error(t, err)
	assert.NoError(t, f.Close())
}

// fileRemoveZipWriter wraps a real zip.Writer and removes the underlying temp
// file during Close so that os.ReadFile(tmpPath) in writeToWithEncryption fails.
type fileRemoveZipWriter struct {
	real *zip.Writer
	file *os.File
}

func (z *fileRemoveZipWriter) Create(name string) (io.Writer, error) {
	return z.real.Create(name)
}

func (z *fileRemoveZipWriter) AddFS(fsys fs.FS) error { return z.real.AddFS(fsys) }

func (z *fileRemoveZipWriter) Close() error {
	if err := z.real.Close(); err != nil {
		return err
	}
	return os.Remove(z.file.Name())
}

func TestWriteToWithEncryptionReadFileError(t *testing.T) {
	// Remove the temp file during ZipWriter.Close so that os.ReadFile fails.
	f := NewFile()
	f.SetZipWriter(func(w io.Writer) ZipWriter {
		return &fileRemoveZipWriter{
			real: zip.NewWriter(w),
			file: w.(*os.File),
		}
	})
	var buf bytes.Buffer
	_, err := f.WriteTo(&buf, Options{Password: "test"})
	assert.Error(t, err)
	assert.NoError(t, f.Close())
}

// errReaderAt always returns the configured error from ReadAt.
type errReaderAt struct{ err error }

func (e *errReaderAt) ReadAt([]byte, int64) (int, error) { return 0, e.err }

func TestFixZip64LFHReadAtError(t *testing.T) {
	f := NewFile()
	f.zip64Entries = []string{"test.xml"}
	err := f.fixZip64LFH(&errReaderAt{err: errors.New("read error")}, nil)
	assert.EqualError(t, err, "read error")
}

// errWriterAt always returns an error from WriteAt.
type errWriterAt struct{ err error }

func (e *errWriterAt) WriteAt([]byte, int64) (int, error) { return 0, e.err }

func TestFixZip64LFHWriteAtError(t *testing.T) {
	f := NewFile()
	f.zip64Entries = []string{"test.xml"}

	// Build a buffer with a valid LFH header
	var hdr bytes.Buffer
	hdr.Write([]byte{0x50, 0x4b, 0x03, 0x04})
	hdr.Write(make([]byte, 22))
	binary.Write(&hdr, binary.LittleEndian, uint16(8))
	hdr.Write(make([]byte, 2))
	hdr.WriteString("test.xml")

	err := f.fixZip64LFH(bytes.NewReader(hdr.Bytes()), &errWriterAt{err: errors.New("write error")})
	assert.EqualError(t, err, "write error")
}

func TestFixZip64LFHLargeFile(t *testing.T) {
	// Data larger than 1MB forces the chunked read loop to iterate more than
	// once, exercising the offset overlap adjustment (offset -= 30).
	f := NewFile()
	f.zip64Entries = []string{"test.xml"}

	// Build a buffer with a valid LFH header followed by >1MB of padding
	var data bytes.Buffer
	data.Write([]byte{0x50, 0x4b, 0x03, 0x04})
	data.Write(make([]byte, 22))
	binary.Write(&data, binary.LittleEndian, uint16(8))
	data.Write(make([]byte, 2))
	data.WriteString("test.xml")
	data.Write(make([]byte, 1100000))

	// Use a copy for writing so patches are captured
	writeBuf := make([]byte, data.Len())
	copy(writeBuf, data.Bytes())

	err := f.fixZip64LFH(bytes.NewReader(writeBuf), &sliceWriterAt{buf: writeBuf})
	assert.NoError(t, err)

	// Verify the header was patched correctly
	assert.Equal(t, uint16(45), binary.LittleEndian.Uint16(writeBuf[4:6]))
	assert.Equal(t, uint32(0xFFFFFFFF), binary.LittleEndian.Uint32(writeBuf[18:22]))
}

// sliceWriterAt implements io.WriterAt backed by a byte slice.
type sliceWriterAt struct{ buf []byte }

func (s *sliceWriterAt) WriteAt(p []byte, off int64) (int, error) {
	copy(s.buf[off:], p)
	return len(p), nil
}

func TestWriteToBufferErrors(t *testing.T) {
	// writeToZip error path
	f := NewFile()
	f.SetZipWriter(func(w io.Writer) ZipWriter {
		return &errZipWriter{
			createFunc: func(string) (io.Writer, error) {
				return nil, errors.New("create error")
			},
		}
	})
	_, err := f.WriteToBuffer()
	assert.Error(t, err)
	assert.NoError(t, f.Close())

	// zw.Close error path
	f = NewFile()
	f.SetZipWriter(func(w io.Writer) ZipWriter {
		return &errZipWriter{closeErr: errors.New("close error")}
	})
	_, err = f.WriteToBuffer()
	assert.EqualError(t, err, "close error")
	assert.NoError(t, f.Close())
}

func TestWriteToBufferWithPassword(t *testing.T) {
	// Exercise the encryption path in WriteToBuffer
	f := NewFile(Options{Password: "pass"})
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "secret"))
	buf, err := f.WriteToBuffer()
	assert.NoError(t, err)
	assert.Greater(t, buf.Len(), 0)
	assert.NoError(t, f.Close())
}
