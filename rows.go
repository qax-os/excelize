package excelize

import (
	"encoding/xml"
	"strings"
  "strconv"
)


// GetRows return all the rows in a sheet
func (f *File) GetRows(sheet string) ([]xlsxRow, error) {
	var xlsx xlsxWorksheet
	name := `xl/worksheets/` + strings.ToLower(sheet) + `.xml`
	err := xml.Unmarshal([]byte(f.readXML(name)), &xlsx)
	if ( err != nil ) {
		return nil, err
	}
	rows := xlsx.SheetData.Row

	return rows, nil

}


// readXMLSST read xmlSST simple function
func readXMLSST(f *File) (xlsxSST, error) {
	shardStrings := xlsxSST{}
	err := xml.Unmarshal([]byte(f.readXML(`xl/sharedStrings.xml`)), &shardStrings)
	return shardStrings, err
}

// GetValueFrom return a value from a column/row cell,
// this function is inteded to be used with for range on rows
// an argument with the xlsx opened file
func (self* xlsxC) GetValueFrom(f *File) (string, error) {
  switch self.T {
    case "s":
      xlsxSI := 0
      xlsxSI, _ = strconv.Atoi(self.V)
      d, err := readXMLSST(f)
      if ( err != nil ) {
        return "", err
      }
      return d.SI[xlsxSI].T, nil
    case "str":
      return self.V, nil
    default:
      return self.V, nil
  } // switch
}
