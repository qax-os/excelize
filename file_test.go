package excelize

import (
	"testing"
)

func BenchmarkWrite(b *testing.B) {
	const s = "This is test data"
	for i := 0; i < b.N; i++ {
		f := NewFile()
		for row := 1; row <= 10000; row++ {
			for col := 1; col <= 20; col++ {
				val, err := CoordinatesToCellName(col, row)
				if err != nil {
					panic(err)
				}
				f.SetCellDefault("Sheet1", val, s)
			}
		}
		// Save xlsx file by the given path.
		err := f.SaveAs("./test.xlsx")
		if err != nil {
			panic(err)
		}
	}

}
