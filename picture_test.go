package excelize

import (
	"fmt"
	_ "image/png"
	"io/ioutil"
	"testing"
)

func BenchmarkAddPictureFromBytes(b *testing.B) {
	f := NewFile()
	imgFile, err := ioutil.ReadFile("logo.png")
	if err != nil {
		panic("unable to load image for benchmark")
	}
	b.ResetTimer()
	for i := 1; i <= b.N; i++ {
		f.AddPictureFromBytes("Sheet1", fmt.Sprint("A", i), "", "logo", ".png", imgFile)
	}
}
