package excelize_test

import (
	"log"
	"os"
	"testing"

	"github.com/xuri/excelize/v2"
)

func Test_WriteStream(t *testing.T) {

	ef := excelize.NewStream(os.Stdout)
	ssw, err := ef.NewStreamWriter("Sheet1")
	if err != nil {
		log.Fatal(err)
	}

	for i := 0; i < 100; i++ {
		var cell, _ = excelize.CoordinatesToCellName(1, i+1)
		err = ssw.SetRow(cell, []interface{}{i, i + 1, i + 2, i + 3}, true)
		if err != nil {
			log.Fatal(err)
		}

	}
	ssw.Flush()
	ef.Close()
}
