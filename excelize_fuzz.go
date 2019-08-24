// +build gofuzz

package excelize

import (
	"bytes"
)

// Fuzz tests parsing
func Fuzz(fuzz []byte) int {
	_, err := OpenReader(bytes.NewReader(fuzz))
	if err != nil {
		return 0
	}
	return 1
}
