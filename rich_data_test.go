package excelize

import (
	"bytes"
	"fmt"
	"io"
	"testing"
)

func TestRichDataInsert(t *testing.T) {
	f := NewFile()
	f.Pkg.Store(defaultXMLMetadata, []byte(`<metadata><valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>`))
	f.Pkg.Store(defaultXMLRdRichValuePart, []byte(`<rvData count="1"><rv s="0"><v>0</v><v>5</v></rv></rvData>`))
	f.Pkg.Store(defaultXMLRdRichValueRel, []byte(`<richValueRels><rel r:id="rId1"/></richValueRels>`))
	f.Pkg.Store(defaultXMLRdRichValueRelRels, []byte(fmt.Sprintf(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="%s" Target="../media/image1.png"/></Relationships>`, SourceRelationshipImage)))
	f.Sheet.Store("xl/worksheets/sheet1.xml", &xlsxWorksheet{
		SheetData: xlsxSheetData{Row: []xlsxRow{
			{R: 1, C: []xlsxC{{R: "A1", T: "e", V: formulaErrorVALUE, Vm: uintPtr(1)}}},
		}},
	})
}

func (f *File) TestRichValueReader() (*xlsxRichValueData, error) {
	var richValue xlsxRichValueData
	if err := f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readXML(defaultXMLRdRichValuePart)))).
		Decode(&richValue); err != nil && err != io.EOF {
		return &richValue, err
	}
	return &richValue, nil
}

func (f *File) TestRichValueArrayDataReader() (*xlsxRichValueArrayData, error) {
	var richValueArray xlsxRichValueArrayData
	if err := f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readXML(defaultXMLRichDataArray)))).
		Decode(&richValueArray); err != nil && err != io.EOF {
		return &richValueArray, err
	}
	return &richValueArray, nil
}

func (f *File) TestRichValueStructureReader() (*xlsxRichValueStructures, error) {
	var richValueStructures xlsxRichValueStructures
	if err := f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readXML(defaultXMLRichDataRichValueStructure)))).
		Decode(&richValueStructures); err != nil && err != io.EOF {
		return &richValueStructures, err
	}
	return &richValueStructures, nil
}

func (f *File) TestRichDataSpb() (*XlsxRichDataSupportingPropertyBags, error) {
	var richDataSpb XlsxRichDataSupportingPropertyBags
	if err := f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readXML(defaultXMLRichDataSupportingPropertyBag)))).
		Decode(&richDataSpb); err != nil && err != io.EOF {
		return &richDataSpb, err
	}
	return &richDataSpb, nil
}

func (f *File) TestRichDataSpbStructure() (*xlsxRichDataSpbStructures, error) {
	var richDataSpbStructure xlsxRichDataSpbStructures
	if err := f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readXML(defaultXMLRichDataSupportingPropertyBagStructure)))).
		Decode(&richDataSpbStructure); err != nil && err != io.EOF {
		return &richDataSpbStructure, err
	}
	return &richDataSpbStructure, nil
}

func (f *File) TestRichDataStyles() (*RichStyleSheet, error) {
	var richDataStyle RichStyleSheet
	if err := f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readXML(defaultXMLRichDataRichStyles)))).
		Decode(&richDataStyle); err != nil && err != io.EOF {
		return &richDataStyle, err
	}
	return &richDataStyle, nil
}

func (f *File) TestRichValueTypes() (*RvTypesInfo, error) {
	var richDataTypes RvTypesInfo
	if err := f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readXML(defaultXMLRichDataRichValueTypes)))).
		Decode(&richDataTypes); err != nil && err != io.EOF {
		return &richDataTypes, err
	}
	return &richDataTypes, nil
}
