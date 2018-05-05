package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func trimSliceSpace(s []string) []string {
	for {
		if len(s) > 0 && s[len(s)-1] == "" {
			s = s[:len(s)-1]
		} else {
			break
		}
	}
	return s
}

func TestRows(t *testing.T) {
	xlsx, err := OpenFile("./test/Book1.xlsx")
	assert.NoError(t, err)

	rows, err := xlsx.Rows("Sheet2")
	assert.NoError(t, err)

	rowStrs := make([][]string, 0)
	var i = 0
	for rows.Next() {
		i++
		columns := rows.Columns()
		//fmt.Println(i, columns)
		rowStrs = append(rowStrs, columns)
	}
	assert.NoError(t, rows.Error())

	dstRows := xlsx.GetRows("Sheet2")
	assert.EqualValues(t, len(dstRows), len(rowStrs))
	for i := 0; i < len(rowStrs); i++ {
		assert.EqualValues(t, trimSliceSpace(dstRows[i]), trimSliceSpace(rowStrs[i]))
	}
}
