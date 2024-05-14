package excelize

import (
	"bytes"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"os"
	"sort"
	"strconv"
)

type Location struct {
	array     []int
	subEntity []int
}

var location Location

func (f *File) AddEntity(sheet, cell string, entityData []byte) error {
	err := f.checkOrCreateRichDataFiles()
	if err != nil {
		return err
	}
	err = f.writeMetadata()
	if err != nil {
		return err
	}
	err = f.writeSheetData(sheet, cell)
	if err != nil {
		return err
	}
	writeJSONToFile(string(entityData), "./console/mashalEntity.json")

	var entity Entity
	err = json.Unmarshal(entityData, &entity)
	if err != nil {
		return err
	}
	if err := writeJSONToFile(entity, "./console/unmashalEntity.json"); err != nil {
		//fmt.println("Error writing JSON to file:", err)
		// return
	}

	err = f.writeRdRichValueStructure(entity)
	if err != nil {
		return err
	}
	err = f.writeRdRichValue(entity)
	if err != nil {
		return err
	}
	err = f.writeSpbData(entity)
	if err != nil {
		return err
	}
	err = f.writeSpbStructure(entity)
	if err != nil {
		return err
	}

	// f.WriteProvider(entity)
	return nil
}

func (f *File) writeRdRichValueStructure(entity Entity) error {

	richValueStructure, err := f.richStructureReader()
	if err != nil {
		return err
	}
	newRichValueStructureKeys := []xlsxRichValueStructureKey{}
	if entity.Text != "" {
		newRichValueStructureKeys = append(newRichValueStructureKeys, xlsxRichValueStructureKey{N: "_DisplayString", T: "s"})
	}
	if entity.Layouts.Compact.Icon != "" {
		newRichValueStructureKeys = append(newRichValueStructureKeys, xlsxRichValueStructureKey{N: "_Icon", T: "s"})
	}
	if entity.Provider.Description != "" {
		newRichValueStructureKeys = append(newRichValueStructureKeys, xlsxRichValueStructureKey{N: "_Provider", T: "spb"})
	}
	properties := entity.Properties
	keys := make([]string, 0, len(properties))
	for key := range properties {
		keys = append(keys, key)
	}
	sort.Strings(keys)

	var array_count int = 0
	for _, key := range keys {
		propertyMap := properties[key].(map[string]interface{})
		if propertyMap["type"] == "String" {
			newRichValueStructureKeys = append(newRichValueStructureKeys, xlsxRichValueStructureKey{N: key, T: "s"}) // type needs to be determined and can vary from spb to r to s
		} else if propertyMap["type"] == "Double" {
			newRichValueStructureKeys = append(newRichValueStructureKeys, xlsxRichValueStructureKey{N: key})
		} else if propertyMap["type"] == "Boolean" {
			newRichValueStructureKeys = append(newRichValueStructureKeys, xlsxRichValueStructureKey{N: key, T: "b"})
		} else if propertyMap["type"] == "Array" {
			//append in _entity
			newRichValueStructureKeys = append(newRichValueStructureKeys, xlsxRichValueStructureKey{N: key, T: "r"})

			arrayStructure, err := f.createArrayStructure(propertyMap["elements"])
			if err != nil {
				return err
			}
			location.array = append(location.array, array_count)
			array_count++
			richValueStructure.S = append(richValueStructure.S, arrayStructure)

			//creating new rv for array

		}
	}
	// fmt.Println(newRichValueStructureKeys)

	newRichValueStructure := xlsxRichValueStructure{
		T: "_entity",
		K: newRichValueStructureKeys,
	}

	richValueStructure.S = append(richValueStructure.S, newRichValueStructure)
	richValueStructure.Count = strconv.Itoa(len(richValueStructure.S))
	// fmt.Println(richValueStructure)
	xmlData, err := xml.Marshal(richValueStructure)
	if err != nil {
		return err
	}
	xmlData = bytes.ReplaceAll(xmlData, []byte(`xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"`), []byte(`xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"`))
	f.saveFileList(defaultXMLRichDataRichValueStructure, xmlData)
	return nil

}

func (f *File) createArrayStructure(elements interface{}) (xlsxRichValueStructure, error) {

	newArrayRichValueStructure := []xlsxRichValueStructureKey{}
	newArrayRichValueStructure = append(newArrayRichValueStructure, xlsxRichValueStructureKey{N: "array", T: "a"})

	arrayRichStructure := xlsxRichValueStructure{
		T: "_array",
		K: newArrayRichValueStructure,
	}
	return arrayRichStructure, nil
}

func (f *File) writeRDRichArrayValue(data interface{}) (xlsxRichValue, error) {
	newRichValue := xlsxRichValue{}
	return newRichValue, nil
}

func (f *File) writeRdRichValue(entity Entity) error {
	richValueStructure, err := f.richStructureReader()
	// _ = richDataArray
	if err != nil {
		return err
	}

	richDataArray, err := f.richDataArrayReader()
	// richDataArray.Count = 0
	// richDataArray.Xmlns s= "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"

	if err != nil {
		return err
	}
	structureValue := len(richValueStructure.S) - 1
	richValue, err := f.richValueReader()
	if err != nil {
		return err
	}
	richDataRichValues := []string{}
	richDataRichValues = append(richDataRichValues, entity.Text)
	richDataRichValues = append(richDataRichValues, entity.Layouts.Compact.Icon)

	if entity.Provider.Description != "" {
		richDataSpbs, err := f.richDataSpbReader()
		if err != nil {
			return err
		}
		spbLen := len(richDataSpbs.SpbData.Spb)
		richDataRichValues = append(richDataRichValues, strconv.Itoa(spbLen))
	}

	writeJSONToFile(entity.Properties, "./console/entityproperties.json")
	var index int = 0
	var count int = 0
	var array_count int = 0

	properties := entity.Properties
	keys := make([]string, 0, len(properties))
	for key := range properties {
		keys = append(keys, key)
	}
	sort.Strings(keys)

	for _, key := range keys {
		propertyMap := properties[key].(map[string]interface{})
		writeJSONToFile(properties[key], fmt.Sprintf("./console/propertyValue%d.json", index))
		index++
		if propertyMap["type"] == "String" {
			richDataRichValues = append(richDataRichValues, propertyMap["basicValue"].(string))
		} else if propertyMap["type"] == "Double" {
			richDataRichValues = append(richDataRichValues, fmt.Sprintf("%v", propertyMap["basicValue"]))
		} else if propertyMap["type"] == "Boolean" {
			if propertyMap["basicValue"] == true {
				richDataRichValues = append(richDataRichValues, "1")
			} else {
				richDataRichValues = append(richDataRichValues, "0")
			}
		} else if propertyMap["type"] == "Array" {
			richDataRichValues = append(richDataRichValues, strconv.Itoa(location.array[count]))

			//creating new rv for array
			richDataRichArrayValues := []string{}
			richDataRichArrayValues = append(richDataRichArrayValues, strconv.Itoa(count))
			newRichArrayValue := xlsxRichValue{
				S: array_count,
				V: richDataRichArrayValues,
			}

			richValue.Rv = append(richValue.Rv, newRichArrayValue)

			//creating rdarrayfile

			f.checkOrCreateXML(defaultXMLRichDataArray, []byte(xml.Header+templateRDArray))

			writeJSONToFile(propertyMap["elements"], "./console/elements.json")
			elements, ok := propertyMap["elements"].([]interface{})
			if !ok {
				fmt.Println("Error: elements is not a slice")
				return err
			}

			maps := elements[0].([]interface{})
			cols := len(maps)
			rows := len(elements)

			values_array := []xlsxRichArrayValue{}

			for _, element_row := range elements {
				newRow, ok := element_row.([]interface{})
				if !ok {
					fmt.Println("Error: elements is not a interface")
					return err
				}
				for _, key := range newRow {
					newMap := key.(map[string]interface{})

					fmt.Println("\n\n\n\newMap")
					fmt.Println(newMap)
					xlsxRichArrayValue := xlsxRichArrayValue{
						Text: "basicValue",
						T:    "s",
					}
					values_array = append(values_array, xlsxRichArrayValue)

				}
			}

			array_data := xlsxRichValuesArray{
				R: strconv.Itoa(rows),
				C: cols,
				V: values_array,
			}
			richDataArray.A = append(richDataArray.A, array_data)
			richDataArray.Count++
			array_count++
			count++
			//for array
			arrayData, err := xml.Marshal(richDataArray)
			if err != nil {
				return err
			}
			f.saveFileList(defaultXMLRichDataArray, arrayData)

		}
	}

	newRichValue := xlsxRichValue{
		S: structureValue,
		V: richDataRichValues,
	}
	richValue.Rv = append(richValue.Rv, newRichValue)
	richValue.Count = len(richValue.Rv)
	// fmt.Println(richValue)
	xmlData, err := xml.Marshal(richValue)
	if err != nil {
		return err
	}
	f.saveFileList(defaultXMLRdRichValuePart, xmlData)
	return nil
}

func (f *File) writeMetadata() error {
	richValue, err := f.richValueReader()
	if err != nil {
		return err
	}
	rvbIdx := len(richValue.Rv)
	metadata, err := f.metadataReader()
	if err != nil {
		return err
	}
	richMetadataExists := checkRichMetadataExists(*metadata)
	if !richMetadataExists {
		newMetadataType := xlsxMetadataType{
			MetadataName:        "XLRICHVALUE",
			MinSupportedVersion: 120000,
			Copy:                1,
			PasteAll:            1,
			PasteValues:         1,
			Merge:               1,
			SplitFirst:          1,
			RowColShift:         1,
			ClearFormats:        1,
			ClearComments:       1,
			Assign:              1,
			Coerce:              1,
		}
		metadata.MetadataTypes.MetadataType = append(metadata.MetadataTypes.MetadataType, newMetadataType)
		metadata.MetadataTypes.Count++
	}
	var maxValuemetadataValue int
	if metadata.ValueMetadata != nil {
		maxValuemetadataValue = metadata.ValueMetadata.Bk[len(metadata.ValueMetadata.Bk)-1].Rc[0].V
	} else {
		maxValuemetadataValue = 0
		metadata.ValueMetadata = &xlsxMetadataBlocks{}
	}
	vmBlock := xlsxMetadataBlock{Rc: []xlsxMetadataRecord{
		{
			T: 1,
			V: maxValuemetadataValue,
		},
	}}
	metadata.ValueMetadata.Bk = append(metadata.ValueMetadata.Bk, vmBlock)
	metadata.ValueMetadata.Count = len(metadata.ValueMetadata.Bk)

	fmBlock := xlsxFutureMetadataBlock{
		ExtLst: ExtLst{
			Ext: Ext{
				URI: ExtURIFutureMetadata,
				Rvb: Rvb{
					I: rvbIdx,
				},
			},
		},
	}
	if len(metadata.FutureMetadata) == 0 {
		metadata.FutureMetadata = append(metadata.FutureMetadata, xlsxFutureMetadata{})
	}
	metadata.FutureMetadata[0].Bk = append(metadata.FutureMetadata[0].Bk, fmBlock)
	metadata.FutureMetadata[0].Count = len(metadata.FutureMetadata[0].Bk)
	metadata.FutureMetadata[0].Name = "XLRICHVALUE"
	metadata.XmlnsXlrd = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
	metadata.Xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	// fmt.Println(metadata)
	xmlData, err := xml.Marshal(metadata)
	if err != nil {
		return err
	}
	xmlData = bytes.ReplaceAll(xmlData, []byte(`></xlrd:rvb>`), []byte(`/>`))
	f.saveFileList(defaultXMLMetadata, xmlData)
	return nil
}

func checkRichMetadataExists(metadata xlsxMetadata) bool {
	count := metadata.MetadataTypes.Count
	if count > 0 {
		for _, metadataType := range metadata.MetadataTypes.MetadataType {
			if metadataType.MetadataName == "XLRICHVALUE" {
				return true
			}
		}
	}
	return false
}

func (f *File) writeSheetData(sheet, cell string) error {
	ws, err := f.workSheetReader(sheet)
	if err != nil {
		return err
	}
	ws.mu.Lock()
	defer ws.mu.Unlock()
	c, col, row, err := ws.prepareCell(cell)
	if err != nil {
		return err
	}
	c.S = ws.prepareCellStyle(col, row, c.S)
	metadata, err := f.metadataReader()
	if err != nil {
		return err
	}
	futureMetadataLen := len(metadata.FutureMetadata)
	var vmValue int
	if futureMetadataLen != 0 {
		vmValue = len(metadata.FutureMetadata[0].Bk)
	} else {
		vmValue = 1
	}
	vmValueUint := uint(vmValue)
	c.Vm = &vmValueUint
	c.V = "#VALUE!"
	c.T = "e"
	// fmt.Println(ws.SheetData)
	return nil
}

func (f *File) writeSpbData(entity Entity) error {
	if entity.Provider.Description != "" {
		f.checkOrCreateXML(defaultXMLRichDataSupportingPropertyBag, []byte(xml.Header+templateSpbData))
		richDataSpbs, err := f.richDataSpbReader()
		if err != nil {
			return err
		}
		richDataSpbStructure, err := f.richDataSpbStructureReader()
		if err != nil {
			return err
		}

		providerSpb := xlsxRichDataSpb{
			S: len(richDataSpbStructure.S),
			V: []string{entity.Provider.Description},
		}

		richDataSpbs.SpbData.Spb = append(richDataSpbs.SpbData.Spb, providerSpb)
		richDataSpbs.SpbData.Count++
		xmlData, err := xml.Marshal(richDataSpbs)
		if err != nil {
			return err
		}
		xmlData = bytes.ReplaceAll(xmlData, []byte(`xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2" xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"`), []byte(`xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"`))
		f.saveFileList(defaultXMLRichDataSupportingPropertyBag, xmlData)
	}
	return nil
}

func (f *File) writeSpbStructure(entity Entity) error {

	if entity.Provider.Description != "" {
		f.checkOrCreateXML(defaultXMLRichDataSupportingPropertyBagStructure, []byte(xml.Header+templateSpbStructure))
		richDataSpbStructure, err := f.richDataSpbStructureReader()
		if err != nil {
			return err
		}

		providerSpbStructureKey := xlsxRichDataSpbStructureKey{
			N: "name",
			T: "s",
		}

		providerSpbStructure := xlsxRichDataSpbStructure{}
		providerSpbStructure.K = append(providerSpbStructure.K, providerSpbStructureKey)

		richDataSpbStructure.S = append(richDataSpbStructure.S, providerSpbStructure)
		richDataSpbStructure.Count++
		xmlData, err := xml.Marshal(richDataSpbStructure)
		if err != nil {
			return err
		}
		xmlData = bytes.ReplaceAll(xmlData, []byte(`xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2" xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"`), []byte(`xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"`))
		f.saveFileList(defaultXMLRichDataSupportingPropertyBagStructure, xmlData)

	}
	return nil
}

func writeJSONToFile(data interface{}, filename string) error {
	// Open a file for writing (create if not exists, truncate if exists)
	file, err := os.Create(filename)
	if err != nil {
		return err
	}
	defer file.Close()

	// Encode the data to JSON and write it to the file
	encoder := json.NewEncoder(file)
	if err := encoder.Encode(data); err != nil {
		return err
	}

	return nil
}
