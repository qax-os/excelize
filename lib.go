package excelize

import (
	"archive/zip"
	"bytes"
	"encoding/gob"
	"io"
	"log"
	"math"
)

// ReadZipReader can be used to read an XLSX in memory without touching the
// filesystem.
func ReadZipReader(r *zip.Reader) (map[string]string, int, error) {
	fileList := make(map[string]string)
	worksheets := 0
	for _, v := range r.File {
		fileList[v.Name] = readFile(v)
		if len(v.Name) > 18 {
			if v.Name[0:19] == "xl/worksheets/sheet" {
				worksheets++
			}
		}
	}
	return fileList, worksheets, nil
}

// readXML provides function to read XML content as string.
func (f *File) readXML(name string) string {
	if content, ok := f.XLSX[name]; ok {
		return content
	}
	return ""
}

// saveFileList provides function to update given file content in file list of
// XLSX.
func (f *File) saveFileList(name, content string) {
	f.XLSX[name] = XMLHeader + content
}

// Read file content as string in a archive file.
func readFile(file *zip.File) string {
	rc, err := file.Open()
	if err != nil {
		log.Fatal(err)
	}
	buff := bytes.NewBuffer(nil)
	io.Copy(buff, rc)
	rc.Close()
	return string(buff.Bytes())
}

// ToAlphaString provides function to convert integer to Excel sheet column
// title. For example convert 36 to column title AK:
//
//     excelize.ToAlphaString(36)
//
func ToAlphaString(value int) string {
	if value < 0 {
		return ""
	}
	var ans string
	i := value + 1
	for i > 0 {
		ans = string((i-1)%26+65) + ans
		i = (i - 1) / 26
	}
	return ans
}

// TitleToNumber provides function to convert Excel sheet column title to int
// (this function doesn't do value check currently). For example convert AK
// and ak to column title 36:
//
//    excelize.TitleToNumber("AK")
//    excelize.TitleToNumber("ak")
//
func TitleToNumber(s string) int {
	weight := 0.0
	sum := 0
	for i := len(s) - 1; i >= 0; i-- {
		ch := int(s[i])
		if int(s[i]) >= int('a') && int(s[i]) <= int('z') {
			ch = int(s[i]) - 32
		}
		sum = sum + (ch-int('A')+1)*int(math.Pow(26, weight))
		weight++
	}
	return sum - 1
}

// letterOnlyMapF is used in conjunction with strings.Map to return only the
// characters A-Z and a-z in a string.
func letterOnlyMapF(rune rune) rune {
	switch {
	case 'A' <= rune && rune <= 'Z':
		return rune
	case 'a' <= rune && rune <= 'z':
		return rune - 32
	}
	return -1
}

// intOnlyMapF is used in conjunction with strings.Map to return only the
// numeric portions of a string.
func intOnlyMapF(rune rune) rune {
	if rune >= 48 && rune < 58 {
		return rune
	}
	return -1
}

// deepCopy provides method to creates a deep copy of whatever is passed to it
// and returns the copy in an interface. The returned value will need to be
// asserted to the correct type.
func deepCopy(dst, src interface{}) error {
	var buf bytes.Buffer
	if err := gob.NewEncoder(&buf).Encode(src); err != nil {
		return err
	}
	return gob.NewDecoder(bytes.NewBuffer(buf.Bytes())).Decode(dst)
}

// boolPtr returns a pointer to a bool with the given value.
func boolPtr(b bool) *bool { return &b }

// defaultTrue returns true if b is nil, or the pointed value.
func defaultTrue(b *bool) bool {
	if b == nil {
		return true
	}
	return *b
}
