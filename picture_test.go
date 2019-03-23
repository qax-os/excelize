package excelize

import (
	"fmt"
	_ "image/png"
	"io/ioutil"
	"path/filepath"
	"testing"
)

func BenchmarkAddPictureFromBytes(b *testing.B) {
	f := NewFile()
	imgFile, err := ioutil.ReadFile(filepath.Join("test", "images", "excel.png"))
	if err != nil {
		b.Error("unable to load image for benchmark")
	}
	b.ResetTimer()
	for i := 1; i <= b.N; i++ {
		f.AddPictureFromBytes("Sheet1", fmt.Sprint("A", i), "", "excel", ".png", imgFile)
	}
}
