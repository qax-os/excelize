package excelize

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"io"
	"log"
	"math"
	"strings"
)

// ReadZip takes a pointer to a zip.ReadCloser and returns a
// xlsx.File struct populated with its contents. In most cases
// ReadZip is not used directly, but is called internally by OpenFile.
func ReadZip(f *zip.ReadCloser) (map[string]string, int, error) {
	defer f.Close()
	return ReadZipReader(&f.Reader)
}

// ReadZipReader can be used to read an XLSX in memory without
// touching the filesystem.
func ReadZipReader(r *zip.Reader) (map[string]string, int, error) {
	fileList := make(map[string]string)
	worksheets := 0
	for _, v := range r.File {
		fileList[v.Name] = readFile(v)
		if len(v.Name) > 18 {
			if v.Name[0:19] == `xl/worksheets/sheet` {
				var xlsx xlsxWorksheet
				xml.Unmarshal([]byte(strings.Replace(fileList[v.Name], `<drawing r:id=`, `<drawing rid=`, -1)), &xlsx)
				xlsx = checkRow(xlsx)
				output, _ := xml.Marshal(xlsx)
				fileList[v.Name] = replaceRelationshipsID(replaceWorkSheetsRelationshipsNameSpace(string(output)))
				worksheets++
			}
		}
	}
	return fileList, worksheets, nil
}

// Read XML content as string and replace drawing property in XML namespace of sheet
func (f *File) readXML(name string) string {
	if content, ok := f.XLSX[name]; ok {
		return strings.Replace(content, `<drawing r:id=`, `<drawing rid=`, -1)
	}
	return ``
}

// Update given file content in file list of XLSX
func (f *File) saveFileList(name string, content string) {
	f.XLSX[name] = XMLHeader + content
}

// Read file content as string in a archive file
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

// Convert integer to Excel sheet column title
func toAlphaString(value int) string {
	if value < 0 {
		return ``
	}
	var ans string
	i := value
	for i > 0 {
		ans = string((i-1)%26+65) + ans
		i = (i - 1) / 26
	}
	return ans
}

// Convert Excel sheet column title to int
func titleToNumber(s string) int {
	weight := 0.0
	sum := 0
	for i := len(s) - 1; i >= 0; i-- {
		sum = sum + (int(s[i])-int('A')+1)*int(math.Pow(26, weight))
		weight++
	}
	return sum - 1
}

// letterOnlyMapF is used in conjunction with strings.Map to return
// only the characters A-Z and a-z in a string
func letterOnlyMapF(rune rune) rune {
	switch {
	case 'A' <= rune && rune <= 'Z':
		return rune
	case 'a' <= rune && rune <= 'z':
		return rune - 32
	}
	return -1
}

// intOnlyMapF is used in conjunction with strings.Map to return only
// the numeric portions of a string.
func intOnlyMapF(rune rune) rune {
	if rune >= 48 && rune < 58 {
		return rune
	}
	return -1
}
