package excelize

import (
	"encoding/json"
	"errors"
	"fmt"
	"strconv"
)

func (f *File) processSpbType(entityMap map[string]interface{}, cellRichStructure xlsxRichValueStructureKey, cellRichDataValue string) error {
	richDataSpbs, err := f.richDataSpbReader()
	if err != nil {
		return err
	}

	richDataSpbStructure, err := f.richDataSpbStructureReader()
	if err != nil {
		return err
	}
	spbIndex, err := strconv.Atoi(cellRichDataValue)
	if err != nil {
		return err
	}

	if spbIndex < 0 || spbIndex >= len(richDataSpbs.SpbData.Spb) {
		return fmt.Errorf("index out of range: %d", spbIndex)
	}

	if cellRichStructure.N == "_Provider" {
		// if provider is not in spb then needs handling
		entityMap["_Provider"] = richDataSpbs.SpbData.Spb[spbIndex].V[0]
	} else if cellRichStructure.N == "_Display" {
		displayData := richDataSpbs.SpbData.Spb[spbIndex]
		for spbDataValueIndex, spbDataValue := range displayData.V {
			entityMap[richDataSpbStructure.S[displayData.S].K[spbDataValueIndex].N] = spbDataValue
		}
	}
	return nil
}

func (f *File) processRichDataArrayType(entityMap map[string]interface{}, cellRichStructure xlsxRichValueStructureKey, subRichStructure xlsxRichValueStructure, richValue *xlsxRichValueData) error {
	richDataArray, err := f.richDataArrayReader()
	if err != nil {
		return err
	}

	for subRichStructureIdx := range subRichStructure.K {
		colCount := richDataArray.A[subRichStructureIdx].C
		rows := make([][]interface{}, 0)
		row := make([]interface{}, 0, colCount)
		for richDataArrayValueIdx, richDataArrayValue := range richDataArray.A[subRichStructureIdx].V {
			if richDataArrayValue.T == "s" {
				row = append(row, richDataArrayValue.Text)
			} else if richDataArrayValue.T == "r" {
				arrayValueRichValueIdx, err := strconv.Atoi(richDataArrayValue.Text)
				if err != nil {
					return err
				}
				if arrayValueRichValueIdx < 0 || arrayValueRichValueIdx >= len(richValue.Rv) {
					return fmt.Errorf("index out of range: %d", arrayValueRichValueIdx)
				}
				arrayValueRichValue := richValue.Rv[arrayValueRichValueIdx]
				if arrayValueRichValue.Fb != 0 {
					unformattedValue := arrayValueRichValue.Fb
					row = append(row, unformattedValue)
				}
			}
			if (richDataArrayValueIdx+1)%colCount == 0 {
				rows = append(rows, row)
				row = make([]interface{}, 0, colCount)
			}
		}
		if len(row) > 0 {
			rows = append(rows, row)
		}
		entityMap[cellRichStructure.N] = rows
	}
	return nil
}

func (f *File) processRichType(entityMap map[string]interface{}, cellRichStructure xlsxRichValueStructureKey, cellRichDataValue string, richValue *xlsxRichValueData) error {

	cellRichDataValueInt, err := strconv.Atoi(cellRichDataValue)
	if err != nil {
		return err
	}

	if cellRichDataValueInt < 0 || cellRichDataValueInt >= len(richValue.Rv) {
		return fmt.Errorf("index out of range: %d", cellRichDataValueInt)
	}

	subRichData := richValue.Rv[cellRichDataValueInt]
	if subRichData.Fb != 0 {
		entityMap[cellRichStructure.N] = subRichData.Fb
	} else {
		richValueStructure, err := f.richStructureReader()
		if err != nil {
			return err
		}
		subRichStructure := richValueStructure.S[subRichData.S]
		if subRichStructure.T == "_entity" {
			// works only if all the nested entity values are unformatted strings
			subRichEntityMap := make(map[string]interface{})
			for subRichDataValueIdx, subRichDatavalue := range subRichData.V {
				subRichDataStructure := subRichStructure.K[subRichDataValueIdx].N
				subRichEntityMap[subRichDataStructure] = subRichDatavalue
			}
			entityMap[cellRichStructure.N] = subRichEntityMap
		} else if subRichStructure.T == "_array" {
			err := f.processRichDataArrayType(entityMap, cellRichStructure, subRichStructure, richValue)
			if err != nil {
				return err
			}
		}
	}
	return nil
}
func (f *File) ReadEntity(sheet, cell string) ([]byte, error) {

	cellType, err := f.GetCellType(sheet, cell)
	if err != nil {
		return nil, err
	}
	if cellType != 3 {
		return nil, errors.New("Cell is not of type entity")
	}

	metadata, err := f.metadataReader()
	if err != nil {
		return nil, err
	}

	ws, _ := f.workSheetReader(sheet)
	ws.mu.Lock()
	defer ws.mu.Unlock()
	for _, row := range ws.SheetData.Row {
		for _, c := range row.C {
			if c.R == cell {
				entity, err := f.readCellEntity(c, metadata)
				if err != nil {
					return nil, err
				}
				entityJSON, err := json.Marshal(entity)
				if err != nil {
					return nil, err
				}
				return entityJSON, nil

			}
		}
	}
	return nil, nil
}

func (f *File) readCellEntity(c xlsxC, metadata *xlsxMetadata) (Entity, error) {
	entity := Entity{}
	entity.Type = "entity"

	entityMap := make(map[string]interface{})
	stringValueMap := make(map[string]string)

	cellMetadataIdx := *c.Vm - 1
	richValueIdx := metadata.FutureMetadata[0].Bk[cellMetadataIdx].ExtLst.Ext.Rvb.I
	richValue, err := f.richValueReader()
	if err != nil {
		return entity, err
	}
	if richValueIdx >= len(richValue.Rv) {
		return entity, err
	}

	cellRichData := richValue.Rv[richValueIdx]

	richValueStructure, err := f.richStructureReader()
	if err != nil {
		return entity, err
	}

	for cellRichDataIdx, cellRichDataValue := range cellRichData.V {
		cellRichStructure := richValueStructure.S[cellRichData.S].K[cellRichDataIdx]

		if cellRichStructure.T == "" {
			entityMap[cellRichStructure.N] = cellRichDataValue
		} else if cellRichStructure.T == "b" {
			boolValue := cellRichDataValue == "1"
			entityMap[cellRichStructure.N] = boolValue
		} else if cellRichStructure.T == "s" {
			processStringType(entityMap, stringValueMap, cellRichStructure, cellRichDataValue)

		} else if cellRichStructure.T == "r" {
			err := f.processRichType(entityMap, cellRichStructure, cellRichDataValue, richValue)
			if err != nil {
				return entity, err
			}

		} else if cellRichStructure.T == "spb" {
			err := f.processSpbType(entityMap, cellRichStructure, cellRichDataValue)
			if err != nil {
				return entity, err
			}
		}
	}

	entity.Text = entityMap["_DisplayString"].(string)
	delete(entityMap, "_DisplayString")

	entity.Layouts.Compact.Icon = entityMap["_Icon"].(string)
	delete(entityMap, "_Icon")

	if titleProp, ok := entityMap["TitleProperty"].(string); ok {
		entity.Layouts.Card.Title.Property = titleProp
		delete(entityMap, "TitleProperty")
	}
	if subTitleProp, ok := entityMap["SubTitleProperty"].(string); ok {
		entity.Layouts.Card.SubTitle.Property = subTitleProp
		delete(entityMap, "SubTitleProperty")
	}
	if providerDesc, ok := entityMap["_Provider"].(string); ok {
		entity.Provider.Description = providerDesc
		delete(entityMap, "_Provider")
	}
	entity.Properties = entityMap

	return entity, nil
}

func processStringType(entityMap map[string]interface{}, stringValueMap map[string]string, cellRichStructure xlsxRichValueStructureKey, cellRichDataValue string) {
	if cellRichStructure.N[0] == '_' {
		if cellRichStructure.N == "_DisplayString" {
			entityMap["_DisplayString"] = cellRichDataValue
		} else if cellRichStructure.N == "_Icon" {
			entityMap["_Icon"] = cellRichDataValue
		}
	} else {
		entityMap[cellRichStructure.N] = cellRichDataValue
	}
}
