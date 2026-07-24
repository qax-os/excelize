package excelize

import (
	"bytes"
	"fmt"
	"io"
	"os"
	"strconv"
	"testing"
)

// writeExcelBench is adapted from github.com/mzimmerman/excelizetest to run
// inside this repo as an in-package benchmark (no external dependency).
func writeExcelBench(data [][]string, out io.Writer) error {
	file := NewFile()
	if len(data) == 0 {
		return nil
	}
	sw, err := file.NewStreamWriter("Sheet1")
	if err != nil {
		return err
	}
	lineInterface := make([]interface{}, len(data[0]))
	for excelLineNum, line := range data {
		lineInterface = lineInterface[:0]
		for x := range line {
			lineInterface = append(lineInterface, line[x])
		}
		cell, _ := CoordinatesToCellName(1, excelLineNum+1)
		if err = sw.SetRow(cell, lineInterface); err != nil {
			return err
		}
	}
	if err = sw.Flush(); err != nil {
		return err
	}
	_, err = file.WriteTo(out)
	return err
}

func benchmarkExcelize(rows, cols int, b *testing.B) {
	buf := bytes.Buffer{}
	for n := 0; n < b.N; n++ {
		b.StopTimer()
		buf.Reset()
		count := 0
		data := make([][]string, rows)
		for x := range data {
			data[x] = make([]string, cols)
			for y := range data[x] {
				data[x][y] = strconv.Itoa(count)
				count++
			}
		}
		b.StartTimer()
		if err := writeExcelBench(data, &buf); err != nil {
			b.Fatalf("error writing excel - %v", err)
		}
	}
}

func BenchmarkExcelize10x10(b *testing.B)       { benchmarkExcelize(10, 10, b) }
func BenchmarkExcelize100x100(b *testing.B)     { benchmarkExcelize(100, 100, b) }
func BenchmarkExcelize1000x1000(b *testing.B)   { benchmarkExcelize(1000, 1000, b) }
func BenchmarkExcelize10000x10000(b *testing.B) { benchmarkExcelize(10000, 10000, b) }
func BenchmarkExcelize1000x10(b *testing.B)     { benchmarkExcelize(1000, 10, b) }
func BenchmarkExcelize10000x10(b *testing.B)    { benchmarkExcelize(10000, 10, b) }
func BenchmarkExcelize100000x10(b *testing.B)   { benchmarkExcelize(100000, 10, b) }
func BenchmarkExcelize100000x100(b *testing.B)  { benchmarkExcelize(100000, 100, b) }
func BenchmarkExcelize10000x1000(b *testing.B)  { benchmarkExcelize(10000, 1000, b) }

// BenchmarkBioSizeSweep measures ns/op and B/op across a range of bufio.Writer
// buffer sizes for a large sheet (~75 MB XML) that exceeds StreamChunkSize.
// Run with: go test -bench=BenchmarkBioSizeSweep -benchmem -count=3 -run='^$'
func BenchmarkBioSizeSweep(b *testing.B) {
	sizes := []int{
		4 << 10,   // 4 KB
		8 << 10,   // 8 KB
		16 << 10,  // 16 KB
		32 << 10,  // 32 KB
		64 << 10,  // 64 KB
		128 << 10, // 128 KB
		256 << 10, // 256 KB
		512 << 10, // 512 KB
		1 << 20,   // 1 MB
		4 << 20,   // 4 MB
	}
	row := make([]interface{}, 100)
	for colID := range row {
		row[colID] = colID * 12345
	}
	for _, sz := range sizes {
		sz := sz
		b.Run(fmt.Sprintf("bio=%s", fmtSize(sz)), func(b *testing.B) {
			b.ReportAllocs()
			for n := 0; n < b.N; n++ {
				file := NewFile()
				sw, _ := file.NewStreamWriter("Sheet1")
				sw.rawData.bioSize = sz
				for rowID := 1; rowID <= 50000; rowID++ {
					cell, _ := CoordinatesToCellName(1, rowID)
					_ = sw.SetRow(cell, row)
				}
				_ = sw.Flush()
				_ = file.Close()
			}
		})
	}
}

// countingWriter wraps an os.File and records every Write call made to it.
type countingWriter struct {
	f     *os.File
	calls int
	bytes int64
}

// fmtSize returns a human-readable label for a byte size.
func fmtSize(n int) string {
	switch {
	case n >= 1<<20:
		return fmt.Sprintf("%dMiB", n>>20)
	case n >= 1<<10:
		return fmt.Sprintf("%dKiB", n>>10)
	default:
		return fmt.Sprintf("%dB", n)
	}
}

func (c *countingWriter) Write(p []byte) (int, error) {
	c.calls++
	c.bytes += int64(len(p))
	return c.f.Write(p)
}

// TestBioSizeIOProfile is not a benchmark — it runs once per bio size and
// prints: total bytes written to disk, number of write syscalls, and average
// write size. Run with: go test -v -run=TestBioSizeIOProfile -count=1
func TestBioSizeIOProfile(t *testing.T) {
	sizes := []int{
		4 << 10,   // 4 KB
		8 << 10,   // 8 KB
		16 << 10,  // 16 KB
		32 << 10,  // 32 KB
		64 << 10,  // 64 KB
		128 << 10, // 128 KB
		256 << 10, // 256 KB
		512 << 10, // 512 KB
		1 << 20,   // 1 MB
		4 << 20,   // 4 MB
	}
	row := make([]interface{}, 100)
	for i := range row {
		row[i] = i * 12345
	}

	t.Logf("%-10s  %12s  %8s  %10s", "bio size", "bytes to disk", "# writes", "avg write")
	t.Logf("%-10s  %12s  %8s  %10s", "--------", "-------------", "--------", "---------")
	for _, sz := range sizes {
		file := NewFile()
		sw, _ := file.NewStreamWriter("Sheet1")
		sw.rawData.bioSize = sz
		sw.rawData.flushSize = 1

		f, err := os.CreateTemp("", "excelize-profile-")
		if err != nil {
			t.Fatal(err)
		}
		cw := &countingWriter{f: f}
		sw.rawData.tmp = f
		for rowID := 1; rowID <= 50000; rowID++ {
			cell, _ := CoordinatesToCellName(1, rowID)
			_ = sw.SetRow(cell, row)
			if rowID == 1 && sw.rawData.bio != nil {
				sw.rawData.bio.Reset(cw)
				cw.calls = 0
				cw.bytes = 0
			}
		}
		_ = sw.Flush()

		avg := int64(0)
		if cw.calls > 0 {
			avg = cw.bytes / int64(cw.calls)
		}
		t.Logf("%-10s  %12s  %8d  %10s",
			fmtSize(sz), fmtSize(int(cw.bytes)), cw.calls, fmtSize(int(avg)))

		_ = file.Close()
		f.Close()
		os.Remove(f.Name())
	}
}

// BenchmarkStringCellClean and BenchmarkStringCellSpecial measure the
// writeEscaped fast path (no special chars) vs slow path (has <, >, &, etc.).
func BenchmarkStringCellClean(b *testing.B) {
	row := make([]interface{}, 50)
	for i := range row {
		row[i] = "normal cell content without special chars"
	}
	b.ReportAllocs()
	for n := 0; n < b.N; n++ {
		file := NewFile()
		sw, _ := file.NewStreamWriter("Sheet1")
		for rowID := 1; rowID <= 10000; rowID++ {
			cell, _ := CoordinatesToCellName(1, rowID)
			_ = sw.SetRow(cell, row)
		}
		_ = sw.Flush()
		_ = file.Close()
	}
}

func BenchmarkStringCellSpecial(b *testing.B) {
	row := make([]interface{}, 50)
	for i := range row {
		row[i] = "content with <special> & \"chars\""
	}
	b.ReportAllocs()
	for n := 0; n < b.N; n++ {
		file := NewFile()
		sw, _ := file.NewStreamWriter("Sheet1")
		for rowID := 1; rowID <= 10000; rowID++ {
			cell, _ := CoordinatesToCellName(1, rowID)
			_ = sw.SetRow(cell, row)
		}
		_ = sw.Flush()
		_ = file.Close()
	}
}
