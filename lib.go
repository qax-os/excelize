package excelize

import (
	"archive/zip"
	"bytes"
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
// title. For example convert 37 to column title AK:
//
//     excelize.ToAlphaString(37)
//
func ToAlphaString(value int) string {
	if value < 0 {
		return ""
	}
	var ans string
	i := value
	for i > 0 {
		ans = string((i-1)%26+65) + ans
		i = (i - 1) / 26
	}
	return ans
}

// titleToNumber provides function to convert Excel sheet column title to int.
func titleToNumber(s string) int {
	weight := 0.0
	sum := 0
	for i := len(s) - 1; i >= 0; i-- {
		sum = sum + (int(s[i])-int('A')+1)*int(math.Pow(26, weight))
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
