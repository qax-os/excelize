package excelize

import (
	"bytes"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"sort"
	"strconv"
	"strings"
	"time"
)

type Location struct {
	maxRdRichValueStructureIndex int
	maxRdRichValueIndex          int
	spbStructureIndex            int
	arrayCount                   int
	spbDataIndex                 int
}

var location = Location{
	maxRdRichValueStructureIndex: 0,
	maxRdRichValueIndex:          0,
	spbStructureIndex:            0,
	arrayCount:                   0,
	spbDataIndex:                 0,
}

func (f *File) AddEntity(sheet, cell string, entityData []byte) error {
	err := f.checkOrCreateRichDataFiles()
	if err != nil {
		return err
	}
	err = f.writeSheetData(sheet, cell)
	if err != nil {
		return err
	}

	var entity Entity
	err = json.Unmarshal(entityData, &entity)
	if err != nil {
		return err
	}

	err = f.writeRdRichValueStructure(entity)
	if err != nil {
		return err
	}
	err = f.writeRdRichValue(entity)
	if err != nil {
		return err
	}
	err = f.writeMetadata()
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
	// err = f.writeRichStyles(entity)
	// if err != nil {
	// 	return err
	// }
	return nil

}

func (f *File) writeRdRichValueStructure(entity Entity) error {

	richValueStructure, err := f.richStructureReader()
	if err != nil {
		return err
	}
	newRichValueStructureKeys := []xlsxRichValueStructureKey{}
	if entity.Layouts.Card.Title.Property != "" || entity.Layouts.Card.SubTitle.Property != "" {
		newRichValueStructureKeys = append(newRichValueStructureKeys, xlsxRichValueStructureKey{N: "_Display", T: "spb"})
	}
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

	for _, key := range keys {
		propertyMap := properties[key].(map[string]interface{})
		if propertyMap["type"] == "String" {
			newRichValueStructureKeys = append(newRichValueStructureKeys, xlsxRichValueStructureKey{N: key, T: "s"})
		} else if propertyMap["type"] == "Double" {
			newRichValueStructureKeys = append(newRichValueStructureKeys, xlsxRichValueStructureKey{N: key})
		} else if propertyMap["type"] == "FormattedNumber" {
			newRichValueStructureKeys = append(newRichValueStructureKeys, xlsxRichValueStructureKey{N: key, T: "s"})
		} else if propertyMap["type"] == "Boolean" {
			newRichValueStructureKeys = append(newRichValueStructureKeys, xlsxRichValueStructureKey{N: key, T: "b"})
		} else if propertyMap["type"] == "Array" {
			//append in _entity
			newRichValueStructureKeys = append(newRichValueStructureKeys, xlsxRichValueStructureKey{N: key, T: "r"})

			arrayStructure, err := f.createArrayStructure()
			if err != nil {
				return err
			}
			richValueStructure.S = append(richValueStructure.S, arrayStructure)

			//creating new rv for array

		} else if propertyMap["type"] == "Entity" {
			newRichValueStructureKeys = append(newRichValueStructureKeys, xlsxRichValueStructureKey{N: key, T: "r"})
			subEntityJson, err := json.Marshal(propertyMap)
			if err != nil {
				fmt.Println(err)
			}
			var subEntity Entity
			err = json.Unmarshal(subEntityJson, &subEntity)
			if err != nil {
				fmt.Println(err)
			}

			err = f.writeSubEntityRdRichValueStructure(subEntity, richValueStructure)
			if err != nil {
				return err
			}

		}
	}

	newRichValueStructure := xlsxRichValueStructure{
		T: "_entity",
		K: newRichValueStructureKeys,
	}

	richValueStructure.S = append(richValueStructure.S, newRichValueStructure)
	richValueStructure.Count = strconv.Itoa(len(richValueStructure.S))
	xmlData, err := xml.Marshal(richValueStructure)
	if err != nil {
		return err
	}
	xmlData = bytes.ReplaceAll(xmlData, []byte(`xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"`), []byte(`xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"`))
	f.saveFileList(defaultXMLRichDataRichValueStructure, xmlData)
	return nil

}

func (f *File) writeSubEntityRdRichValueStructure(subEntity Entity, richValueStructures *xlsxRichValueStructures) error {
	properties := subEntity.Properties
	keys := make([]string, 0, len(properties))
	for key := range properties {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	newRichValueStructureKeys := []xlsxRichValueStructureKey{}

	if subEntity.Text != "" {
		newRichValueStructureKeys = append(newRichValueStructureKeys, xlsxRichValueStructureKey{N: "_DisplayString", T: "s"})
	}

	for _, key := range keys {
		propertyMap := properties[key].(map[string]interface{})
		if propertyMap["type"] == "String" {
			newRichValueStructureKeys = append(newRichValueStructureKeys, xlsxRichValueStructureKey{N: key, T: "s"})
		}
	}

	richValueStructure := xlsxRichValueStructure{
		T: "_entity",
		K: newRichValueStructureKeys,
	}
	richValueStructures.S = append(richValueStructures.S, richValueStructure)

	return nil
}

func (f *File) writeSubentityRdRichValue(subEntity Entity, richValueData *xlsxRichValueData) error {

	properties := subEntity.Properties
	keys := make([]string, 0, len(properties))
	for key := range properties {
		keys = append(keys, key)
	}
	sort.Strings(keys)

	richDataRichValues := []string{}
	richDataRichValues = append(richDataRichValues, subEntity.Text)
	for _, key := range keys {
		propertyMap := properties[key].(map[string]interface{})
		if propertyMap["type"] == "String" {
			richDataRichValues = append(richDataRichValues, propertyMap["basicValue"].(string))
		}
	}

	newRichValue := xlsxRichValue{
		S: location.maxRdRichValueStructureIndex,
		V: richDataRichValues,
	}
	location.maxRdRichValueStructureIndex++

	richValueData.Rv = append(richValueData.Rv, newRichValue)
	location.maxRdRichValueIndex++

	return nil
}

func (f *File) createArrayStructure() (xlsxRichValueStructure, error) {

	newArrayRichValueStructure := []xlsxRichValueStructureKey{}
	newArrayRichValueStructure = append(newArrayRichValueStructure, xlsxRichValueStructureKey{N: "array", T: "a"})

	arrayRichStructure := xlsxRichValueStructure{
		T: "_array",
		K: newArrayRichValueStructure,
	}
	return arrayRichStructure, nil
}

func (f *File) createArrayFbStructure(elements interface{}) (xlsxRichValueStructure, error) {
	var arrayFbStructure xlsxRichValueStructure

	arrayElements := elements.([]interface{})
	for _, element_row := range arrayElements {
		newRow := element_row.([]interface{})
		for _, key := range newRow {
			newMap := key.(map[string]interface{})

			if newMap["type"].(string) == "FormattedNumber" {

				newArrayFbStructure := []xlsxRichValueStructureKey{}
				newArrayFbStructure = append(newArrayFbStructure, xlsxRichValueStructureKey{N: "_Format", T: "spb"})

				arrayFbStructure = xlsxRichValueStructure{
					T: "_formattednumber",
					K: newArrayFbStructure,
				}

			}
		}
	}

	return arrayFbStructure, nil
}

func (f *File) writeRdRichValue(entity Entity) error {
	richValue, err := f.richValueReader()
	if err != nil {
		return err
	}
	richDataRichValues := []string{}
	if entity.Layouts.Card.Title.Property != "" || entity.Layouts.Card.SubTitle.Property != "" {
		richDataRichValues = append(richDataRichValues, strconv.Itoa(location.spbDataIndex))
		location.spbDataIndex++
	}
	richDataRichValues = append(richDataRichValues, entity.Text)
	if entity.Layouts.Compact.Icon != "" {
		richDataRichValues = append(richDataRichValues, entity.Layouts.Compact.Icon)
	}
	if entity.Provider.Description != "" {

		richDataRichValues = append(richDataRichValues, strconv.Itoa(location.spbDataIndex))
		location.spbDataIndex++
	}

	var index int = 0

	properties := entity.Properties
	keys := make([]string, 0, len(properties))
	for key := range properties {
		keys = append(keys, key)
	}
	sort.Strings(keys)

	for _, key := range keys {
		propertyMap := properties[key].(map[string]interface{})
		index++
		if propertyMap["type"] == "String" {
			richDataRichValues = append(richDataRichValues, propertyMap["basicValue"].(string))
		} else if propertyMap["type"] == "FormattedNumber" {
			formattedValue := propertyMap["basicValue"].(float64)
			var formattedString string
			if propertyMap["numberFormat"] == "yyyy-mm-dd" {
				base := time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC)
				formattedDateValue := int(formattedValue)
				target := base.AddDate(0, 0, formattedDateValue-2)
				formattedString = target.Format("2006-01-02")
			} else {
				numFormat := propertyMap["numberFormat"]
				firstChar := string(numFormat.(string)[0])
				parts := strings.Split(numFormat.(string), ".")
				decimalPlaces := 0
				if len(parts) > 1 {
					decimalPlaces = strings.Count(parts[1], "0")
				}
				formattedString = fmt.Sprintf("%s%.*f", firstChar, decimalPlaces, formattedValue)
			}
			richDataRichValues = append(richDataRichValues, formattedString)

		} else if propertyMap["type"] == "Double" {
			richDataRichValues = append(richDataRichValues, fmt.Sprintf("%v", propertyMap["basicValue"]))
		} else if propertyMap["type"] == "Boolean" {
			if propertyMap["basicValue"] == true {
				richDataRichValues = append(richDataRichValues, "1")
			} else {
				richDataRichValues = append(richDataRichValues, "0")
			}
		} else if propertyMap["type"] == "Array" {
			richDataRichValues = append(richDataRichValues, strconv.Itoa(location.maxRdRichValueIndex))

			//creating new rv for array
			richDataRichArrayValues := []string{}
			richDataRichArrayValues = append(richDataRichArrayValues, strconv.Itoa(location.arrayCount))
			location.arrayCount++
			newRichArrayValue := xlsxRichValue{
				S: location.maxRdRichValueStructureIndex,
				V: richDataRichArrayValues,
			}
			location.maxRdRichValueStructureIndex++

			richValue.Rv = append(richValue.Rv, newRichArrayValue)
			location.maxRdRichValueIndex++

			//creating rdarrayfile

			f.checkOrCreateXML(defaultXMLRichDataArray, []byte(xml.Header+templateRDArray))

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

					if newMap["type"].(string) == "String" {
						basicValue := newMap["basicValue"].(string)

						xlsxRichArrayValue := xlsxRichArrayValue{
							Text: basicValue,
							T:    "s",
						}
						values_array = append(values_array, xlsxRichArrayValue)
					} else if newMap["type"].(string) == "FormattedNumber" {

						formattedValue := newMap["basicValue"].(float64)
						var formattedString string
						if newMap["numberFormat"] == "yyyy-mm-dd" {
							base := time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC)
							formattedDateValue := int(formattedValue)
							target := base.AddDate(0, 0, formattedDateValue-2)
							formattedString = target.Format("2006-01-02")
						} else {
							numFormat := newMap["numberFormat"]
							firstChar := string(numFormat.(string)[0])
							parts := strings.Split(numFormat.(string), ".")
							decimalPlaces := 0
							if len(parts) > 1 {
								decimalPlaces = strings.Count(parts[1], "0")
							}
							formattedString = fmt.Sprintf("%s%.*f", firstChar, decimalPlaces, formattedValue)
						}
						xlsxRichArrayValue := xlsxRichArrayValue{
							Text: formattedString,
							T:    "s",
						}
						values_array = append(values_array, xlsxRichArrayValue)

					}

				}
			}

			richDataArray, err := f.richDataArrayReader()
			if err != nil {
				return err
			}

			array_data := xlsxRichValuesArray{
				R: strconv.Itoa(rows),
				C: cols,
				V: values_array,
			}
			richDataArray.A = append(richDataArray.A, array_data)
			richDataArray.Count++

			arrayData, err := xml.Marshal(richDataArray)
			if err != nil {
				return err
			}
			f.saveFileList(defaultXMLRichDataArray, arrayData)

		} else if propertyMap["type"] == "Entity" {
			richDataRichValues = append(richDataRichValues, strconv.Itoa(location.maxRdRichValueIndex))

			subEntityJson, err := json.Marshal(propertyMap)
			if err != nil {
				fmt.Println(err)
			}
			var subEntity Entity
			err = json.Unmarshal(subEntityJson, &subEntity)
			if err != nil {
				fmt.Println(err)
			}

			f.writeSubentityRdRichValue(subEntity, richValue)
		}
	}

	newRichValue := xlsxRichValue{
		S: location.maxRdRichValueStructureIndex,
		V: richDataRichValues,
	}
	location.maxRdRichValueStructureIndex++

	richValue.Rv = append(richValue.Rv, newRichValue)
	richValue.Count = len(richValue.Rv)
	xmlData, err := xml.Marshal(richValue)
	if err != nil {
		return err
	}
	f.saveFileList(defaultXMLRdRichValuePart, xmlData)
	return nil
}

func (f *File) writeMetadata() error {
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
					I: location.maxRdRichValueIndex,
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
	return nil
}

func (f *File) writeSpbData(entity Entity) error {

	properties := entity.Properties
	keys := make([]string, 0, len(properties))
	for key := range properties {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	f.checkOrCreateXML(defaultXMLRichDataSupportingPropertyBag, []byte(xml.Header+templateSpbData))

	richDataSpbs, err := f.richDataSpbReader()
	if err != nil {
		return err
	}

	if entity.Provider.Description != "" || entity.Layouts.Card.SubTitle.Property != "" || entity.Layouts.Card.Title.Property != "" {
		f.checkOrCreateXML(defaultXMLRichDataSupportingPropertyBag, []byte(xml.Header+templateSpbData))
		if entity.Layouts.Card.Title.Property != "" || entity.Layouts.Card.SubTitle.Property != "" {
			titlesSpb := xlsxRichDataSpb{
				S: location.spbStructureIndex,
			}
			if entity.Layouts.Card.Title.Property != "" {
				titlesSpb.V = append(titlesSpb.V, entity.Layouts.Card.Title.Property)
			}
			if entity.Layouts.Card.SubTitle.Property != "" {
				titlesSpb.V = append(titlesSpb.V, entity.Layouts.Card.SubTitle.Property)
			}

			richDataSpbs.SpbData.Spb = append(richDataSpbs.SpbData.Spb, titlesSpb)
			location.spbStructureIndex++
			richDataSpbs.SpbData.Count++
		}

		if entity.Provider.Description != "" {

			providerSpb := xlsxRichDataSpb{
				S: location.spbStructureIndex,
				V: []string{entity.Provider.Description},
			}
			location.spbStructureIndex++

			richDataSpbs.SpbData.Spb = append(richDataSpbs.SpbData.Spb, providerSpb)
			richDataSpbs.SpbData.Count++
		}
	}

	xmlData, err := xml.Marshal(richDataSpbs)
	if err != nil {
		return err
	}
	xmlData = bytes.ReplaceAll(xmlData, []byte(`xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2" xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"`), []byte(`xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"`))
	f.saveFileList(defaultXMLRichDataSupportingPropertyBag, xmlData)

	return nil
}

func (f *File) writeSpbStructure(entity Entity) error {

	properties := entity.Properties
	keys := make([]string, 0, len(properties))
	for key := range properties {
		keys = append(keys, key)
	}
	sort.Strings(keys)

	f.checkOrCreateXML(defaultXMLRichDataSupportingPropertyBagStructure, []byte(xml.Header+templateSpbStructure))
	richDataSpbStructure, err := f.richDataSpbStructureReader()
	if err != nil {
		return err
	}

	if entity.Provider.Description != "" || entity.Layouts.Card.SubTitle.Property != "" || entity.Layouts.Card.Title.Property != "" {

		spbStructure := xlsxRichDataSpbStructure{}

		if entity.Layouts.Card.Title.Property != "" || entity.Layouts.Card.SubTitle.Property != "" {

			if entity.Layouts.Card.Title.Property != "" {
				titleSpbStructureKey := xlsxRichDataSpbStructureKey{
					N: "TitleProperty",
					T: "s",
				}
				spbStructure.K = append(spbStructure.K, titleSpbStructureKey)
			}

			if entity.Layouts.Card.SubTitle.Property != "" {
				subtitleSpbStructureKey := xlsxRichDataSpbStructureKey{
					N: "SubTitleProperty",
					T: "s",
				}
				spbStructure.K = append(spbStructure.K, subtitleSpbStructureKey)
			}

			richDataSpbStructure.S = append(richDataSpbStructure.S, spbStructure)
			richDataSpbStructure.Count++
		}

		if entity.Provider.Description != "" {

			providerSpbStructureKey := xlsxRichDataSpbStructureKey{
				N: "name",
				T: "s",
			}

			providerSpbStructure := xlsxRichDataSpbStructure{}
			providerSpbStructure.K = append(providerSpbStructure.K, providerSpbStructureKey)

			richDataSpbStructure.S = append(richDataSpbStructure.S, providerSpbStructure)
			richDataSpbStructure.Count++
		}

	}

	xmlData, err := xml.Marshal(richDataSpbStructure)
	if err != nil {
		return err
	}
	xmlData = bytes.ReplaceAll(xmlData, []byte(`xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2" xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"`), []byte(`xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"`))
	f.saveFileList(defaultXMLRichDataSupportingPropertyBagStructure, xmlData)
	return nil
}

func (f *File) writeRichStyles(entity Entity) error {
	f.checkOrCreateXML(defaultXMLRichDataRichStyles, []byte(xml.Header+templateRichStyles))
	rdRichStyles, err := f.richDataStyleReader()
	if err != nil {
		return err
	}

	properties := entity.Properties
	keys := make([]string, 0, len(properties))
	for key := range properties {
		keys = append(keys, key)
	}
	sort.Strings(keys)

	for _, key := range keys {
		propertyMap := properties[key].(map[string]interface{})
		if propertyMap["type"] == "FormattedNumber" {
			newRpv := Rpv{
				I:    "0",
				Text: propertyMap["numberFormat"].(string),
			}
			newRichSty := RSty{}
			newRichSty.Rpv = newRpv
			rdRichStyles.RichStyles.RSty = append(rdRichStyles.RichStyles.RSty, newRichSty)
		}
	}
	fmt.Println(rdRichStyles)
	xmlData, err := xml.Marshal(rdRichStyles)
	if err != nil {
		return err
	}
	xmlData = bytes.ReplaceAll(xmlData, []byte(`xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2" xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"`), []byte(`xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"`))
	f.saveFileList(defaultXMLRichDataRichStyles, xmlData)
	return nil
}
