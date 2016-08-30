package excelize

import (
	"archive/zip"
	"bytes"
	"io"
	"log"
	"math"
	"os"
	"regexp"
	"strconv"
	"strings"
)

// ReadZip() takes a pointer to a zip.ReadCloser and returns a
// xlsx.File struct populated with its contents.  In most cases
// ReadZip is not used directly, but is called internally by OpenFile.
func ReadZip(f *zip.ReadCloser) ([]FileList, error) {
	defer f.Close()
	return ReadZipReader(&f.Reader)
}

// ReadZipReader() can be used to read an XLSX in memory without
// touching the filesystem.
func ReadZipReader(r *zip.Reader) ([]FileList, error) {
	var fileList []FileList
	for _, v := range r.File {
		singleFile := FileList{
			Key:   v.Name,
			Value: readFile(v),
		}
		fileList = append(fileList, singleFile)
	}
	return fileList, nil
}

// Read XML content as string and replace drawing property in XML namespace of sheet
func readXml(files []FileList, name string) string {
	for _, file := range files {
		if file.Key == name {
			return strings.Replace(file.Value, "<drawing r:id=", "<drawing rid=", -1)
		}
	}
	return ``
}

// Update given file content in file list of XLSX
func saveFileList(files []FileList, name string, content string) []FileList {
	for k, v := range files {
		if v.Key == name {
			files = files[:k+copy(files[k:], files[k+1:])]
			files = append(files, FileList{
				Key:   name,
				Value: XMLHeader + content,
			})
			return files
		}
	}
	files = append(files, FileList{
		Key:   name,
		Value: XMLHeader + content,
	})
	return files
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

// Check the file exists
func pathExist(_path string) bool {
	_, err := os.Stat(_path)
	if err != nil && os.IsNotExist(err) {
		return false
	}
	return true
}

// Split Excel sheet column title to string and integer, return XAxis
func getColIndex(axis string) string {
	r, err := regexp.Compile(`[^\D]`)
	if err != nil {
		log.Fatal(err)
	}
	return string(r.ReplaceAll([]byte(axis), []byte("")))
}

// Split Excel sheet column title to string and integer, return YAxis
func getRowIndex(axis string) int {
	r, err := regexp.Compile(`[\D]`)
	if err != nil {
		log.Fatal(err)
	}
	row, err := strconv.Atoi(string(r.ReplaceAll([]byte(axis), []byte(""))))
	if err != nil {
		log.Fatal(err)
	}
	return row
}
