package excelize

import (
	"bytes"
	"fmt"
	"strconv"
	"testing"
)

func BenchmarkCompressionLevels(b *testing.B) {
	const rows, cols = 10000, 50
	data := make([][]string, rows)
	for i := range data {
		data[i] = make([]string, cols)
		for j := range data[i] {
			data[i][j] = strconv.Itoa(i*cols + j)
		}
	}
	levels := []struct {
		name  string
		level Compression
	}{
		{"Default", CompressionDefault},
		{"None", CompressionNone},
		{"BestSpeed", CompressionBestSpeed},
	}
	for _, lvl := range levels {
		lvl := lvl
		b.Run(lvl.name, func(b *testing.B) {
			b.ReportAllocs()
			var buf bytes.Buffer
			for n := 0; n < b.N; n++ {
				buf.Reset()
				file := NewFile(Options{Compression: lvl.level})
				sw, _ := file.NewStreamWriter("Sheet1")
				row := make([]interface{}, cols)
				for i, line := range data {
					for j := range line {
						row[j] = line[j]
					}
					cell, _ := CoordinatesToCellName(1, i+1)
					_ = sw.SetRow(cell, row)
				}
				_ = sw.Flush()
				_, _ = file.WriteTo(&buf)
				b.SetBytes(int64(buf.Len()))
			}
			b.ReportMetric(float64(buf.Len()), "output-bytes")
		})
	}
}

func BenchmarkCompressionBySize(b *testing.B) {
	sizes := []struct{ rows, cols int }{
		{100, 10}, {1000, 50}, {10000, 50}, {50000, 20},
	}
	levels := []struct {
		name  string
		level Compression
	}{
		{"Default", CompressionDefault},
		{"None", CompressionNone},
		{"BestSpeed", CompressionBestSpeed},
	}
	for _, sz := range sizes {
		sz := sz
		for _, lvl := range levels {
			lvl := lvl
			name := fmt.Sprintf("%dx%d/%s", sz.rows, sz.cols, lvl.name)
			b.Run(name, func(b *testing.B) {
				b.ReportAllocs()
				var buf bytes.Buffer
				row := make([]interface{}, sz.cols)
				for j := range row {
					row[j] = "cell data value"
				}
				for n := 0; n < b.N; n++ {
					buf.Reset()
					file := NewFile(Options{Compression: lvl.level})
					sw, _ := file.NewStreamWriter("Sheet1")
					for i := 1; i <= sz.rows; i++ {
						cell, _ := CoordinatesToCellName(1, i)
						_ = sw.SetRow(cell, row)
					}
					_ = sw.Flush()
					_, _ = file.WriteTo(&buf)
				}
			})
		}
	}
}
