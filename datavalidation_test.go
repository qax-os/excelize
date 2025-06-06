// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.
//
// Package excelize providing a set of functions that allow you to write to and
// read from XLAM / XLSM / XLSX / XLTM / XLTX files. Supports reading and
// writing spreadsheet documents generated by Microsoft Excel™ 2007 and later.
// Supports complex components by high compatibility, and provided streaming
// API for generating or reading data from a worksheet with huge amounts of
// data. This library needs Go version 1.23 or later.

package excelize

import (
	"fmt"
	"math"
	"path/filepath"
	"strings"
	"sync"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestDataValidation(t *testing.T) {
	resultFile := filepath.Join("test", "TestDataValidation.xlsx")

	f := NewFile()

	dv := NewDataValidation(true)
	dv.Sqref = "A1:B2"
	assert.NoError(t, dv.SetRange(10, 20, DataValidationTypeWhole, DataValidationOperatorBetween))
	dv.SetError(DataValidationErrorStyleStop, "error title", "error body")
	dv.SetError(DataValidationErrorStyleWarning, "error title", "error body")
	dv.SetError(DataValidationErrorStyleInformation, "error title", "error body")
	assert.NoError(t, f.AddDataValidation("Sheet1", dv))

	dataValidations, err := f.GetDataValidations("Sheet1")
	assert.NoError(t, err)
	assert.Len(t, dataValidations, 1)

	assert.NoError(t, f.SaveAs(resultFile))

	dv = NewDataValidation(true)
	dv.Sqref = "A3:B4"
	assert.NoError(t, dv.SetRange(10, 20, DataValidationTypeWhole, DataValidationOperatorGreaterThan))
	dv.SetInput("input title", "input body")
	assert.NoError(t, f.AddDataValidation("Sheet1", dv))

	dataValidations, err = f.GetDataValidations("Sheet1")
	assert.NoError(t, err)
	assert.Len(t, dataValidations, 2)

	assert.NoError(t, f.SaveAs(resultFile))

	_, err = f.NewSheet("Sheet2")
	assert.NoError(t, err)
	assert.NoError(t, f.SetSheetRow("Sheet2", "A2", &[]interface{}{"B2", 1}))
	assert.NoError(t, f.SetSheetRow("Sheet2", "A3", &[]interface{}{"B3", 3}))
	dv = NewDataValidation(true)
	dv.Sqref = "A1:B1"
	assert.NoError(t, dv.SetRange("INDIRECT($A$2)", "INDIRECT($A$3)", DataValidationTypeWhole, DataValidationOperatorBetween))
	dv.SetError(DataValidationErrorStyleStop, "error title", "error body")
	assert.NoError(t, f.AddDataValidation("Sheet2", dv))
	dataValidations, err = f.GetDataValidations("Sheet1")
	assert.NoError(t, err)
	assert.Len(t, dataValidations, 2)
	dataValidations, err = f.GetDataValidations("Sheet2")
	assert.NoError(t, err)
	assert.Len(t, dataValidations, 1)

	dv = NewDataValidation(true)
	dv.Sqref = "A5:B6"
	for _, listValid := range [][]string{
		{"1", "2", "3"},
		{"=A1"},
		{strings.Repeat("&", MaxFieldLength)},
		{strings.Repeat("\u4E00", MaxFieldLength)},
		{strings.Repeat("\U0001F600", 100), strings.Repeat("\u4E01", 50), "<&>"},
		{`A<`, `B>`, `C"`, "D\t", `E'`, `F`},
	} {
		dv.Formula1 = ""
		assert.NoError(t, dv.SetDropList(listValid),
			"SetDropList failed for valid input %v", listValid)
		assert.NotEmpty(t, dv.Formula1,
			"Formula1 should not be empty for valid input %v", listValid)
	}
	assert.Equal(t, `"A&lt;,B&gt;,C"",D	,E',F"`, dv.Formula1)
	assert.NoError(t, f.AddDataValidation("Sheet1", dv))

	dataValidations, err = f.GetDataValidations("Sheet1")
	assert.NoError(t, err)
	assert.Len(t, dataValidations, 3)

	// Test get data validation on no exists worksheet
	_, err = f.GetDataValidations("SheetN")
	assert.EqualError(t, err, "sheet SheetN does not exist")
	// Test get data validation with invalid sheet name
	_, err = f.GetDataValidations("Sheet:1")
	assert.EqualError(t, err, ErrSheetNameInvalid.Error())

	assert.NoError(t, f.SaveAs(resultFile))

	// Test get data validation on a worksheet without data validation settings
	f = NewFile()
	dataValidations, err = f.GetDataValidations("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, []*DataValidation(nil), dataValidations)

	// Test get data validations which storage in the extension lists
	f = NewFile()
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).ExtLst = &xlsxExtLst{Ext: fmt.Sprintf(`<ext uri="%s" xmlns:x14="%s"><x14:dataValidations><x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>Sheet1!$B$1:$B$5</xm:f></x14:formula1><xm:sqref>A7:B8</xm:sqref></x14:dataValidation></x14:dataValidations></ext>`, ExtURIDataValidations, NameSpaceSpreadSheetX14.Value)}
	dataValidations, err = f.GetDataValidations("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, []*DataValidation{
		{
			AllowBlank: true,
			Type:       "list",
			Formula1:   "Sheet1!$B$1:$B$5",
			Sqref:      "A7:B8",
		},
	}, dataValidations)

	// Test get data validations with invalid extension list characters
	ws.(*xlsxWorksheet).ExtLst = &xlsxExtLst{Ext: fmt.Sprintf(`<ext uri="%s" xmlns:x14="%s"><x14:dataValidations></x14:dataValidation></x14:dataValidations></ext>`, ExtURIDataValidations, NameSpaceSpreadSheetX14.Value)}
	_, err = f.GetDataValidations("Sheet1")
	assert.EqualError(t, err, "XML syntax error on line 1: element <dataValidations> closed by </dataValidation>")

	// Test get validations without validations
	assert.Nil(t, getDataValidations(nil))
	assert.Nil(t, getDataValidations(&xlsxDataValidations{DataValidation: []*xlsxDataValidation{nil}}))
}

func TestDataValidationError(t *testing.T) {
	resultFile := filepath.Join("test", "TestDataValidationError.xlsx")

	f := NewFile()
	assert.NoError(t, f.SetCellStr("Sheet1", "E1", "E1"))
	assert.NoError(t, f.SetCellStr("Sheet1", "E2", "E2"))
	assert.NoError(t, f.SetCellStr("Sheet1", "E3", "E3"))

	dv := NewDataValidation(true)
	dv.SetSqref("A7:B8")
	dv.SetSqref("A7:B8")
	dv.SetSqrefDropList("$E$1:$E$3")

	assert.NoError(t, f.AddDataValidation("Sheet1", dv))

	dv = NewDataValidation(true)
	err := dv.SetDropList(make([]string, 258))
	if dv.Formula1 != "" {
		t.Errorf("data validation error. Formula1 must be empty!")
		return
	}
	assert.EqualError(t, err, ErrDataValidationFormulaLength.Error())
	assert.EqualError(t, dv.SetRange(nil, 20, DataValidationTypeWhole, DataValidationOperatorBetween), ErrParameterInvalid.Error())
	assert.EqualError(t, dv.SetRange(10, nil, DataValidationTypeWhole, DataValidationOperatorBetween), ErrParameterInvalid.Error())
	assert.NoError(t, dv.SetRange(10, 20, DataValidationTypeWhole, DataValidationOperatorGreaterThan))
	dv.SetSqref("A9:B10")

	assert.NoError(t, f.AddDataValidation("Sheet1", dv))

	// Test width invalid data validation formula
	prevFormula1 := dv.Formula1
	for _, keys := range [][]string{
		make([]string, 257),
		{strings.Repeat("s", 256)},
		{strings.Repeat("\u4E00", 256)},
		{strings.Repeat("\U0001F600", 128)},
		{strings.Repeat("\U0001F600", 127), "s"},
	} {
		err = dv.SetDropList(keys)
		assert.Equal(t, prevFormula1, dv.Formula1,
			"Formula1 should be unchanged for invalid input %v", keys)
		assert.EqualError(t, err, ErrDataValidationFormulaLength.Error())
	}
	assert.NoError(t, f.AddDataValidation("Sheet1", dv))
	assert.NoError(t, dv.SetRange(
		-math.MaxFloat32, math.MaxFloat32,
		DataValidationTypeWhole, DataValidationOperatorGreaterThan))
	assert.EqualError(t, dv.SetRange(
		-math.MaxFloat64, math.MaxFloat32,
		DataValidationTypeWhole, DataValidationOperatorGreaterThan), ErrDataValidationRange.Error())
	assert.EqualError(t, dv.SetRange(
		math.SmallestNonzeroFloat64, math.MaxFloat64,
		DataValidationTypeWhole, DataValidationOperatorGreaterThan), ErrDataValidationRange.Error())
	assert.NoError(t, f.SaveAs(resultFile))

	// Test add data validation on no exists worksheet
	f = NewFile()
	assert.EqualError(t, f.AddDataValidation("SheetN", nil), "sheet SheetN does not exist")

	// Test add data validation with invalid sheet name
	f = NewFile()
	assert.EqualError(t, f.AddDataValidation("Sheet:1", nil), ErrSheetNameInvalid.Error())
}

func TestDeleteDataValidation(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.DeleteDataValidation("Sheet1", "A1:B2"))

	dv := NewDataValidation(true)
	dv.Sqref = "A1:B2"
	assert.NoError(t, dv.SetRange(10, 20, DataValidationTypeWhole, DataValidationOperatorBetween))
	dv.SetInput("input title", "input body")
	assert.NoError(t, f.AddDataValidation("Sheet1", dv))
	assert.NoError(t, f.DeleteDataValidation("Sheet1", "A1:B2"))

	dv.Sqref = "A1"
	assert.NoError(t, f.AddDataValidation("Sheet1", dv))
	assert.NoError(t, f.DeleteDataValidation("Sheet1", "B1"))
	assert.NoError(t, f.DeleteDataValidation("Sheet1", "A1"))

	dv.Sqref = "C2:C5"
	assert.NoError(t, f.AddDataValidation("Sheet1", dv))
	assert.NoError(t, f.DeleteDataValidation("Sheet1", "C4"))

	dv = NewDataValidation(true)
	dv.Sqref = "D2:D2 D3 D4"
	assert.NoError(t, dv.SetRange(10, 20, DataValidationTypeWhole, DataValidationOperatorBetween))
	dv.SetInput("input title", "input body")
	assert.NoError(t, f.AddDataValidation("Sheet1", dv))
	assert.NoError(t, f.DeleteDataValidation("Sheet1", "D3"))

	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestDeleteDataValidation.xlsx")))

	dv.Sqref = "A"
	assert.NoError(t, f.AddDataValidation("Sheet1", dv))
	assert.EqualError(t, f.DeleteDataValidation("Sheet1", "A1"), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())

	assert.EqualError(t, f.DeleteDataValidation("Sheet1", "A1:A"), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())
	ws, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok)
	ws.(*xlsxWorksheet).DataValidations.DataValidation[0].Sqref = "A1:A"
	assert.EqualError(t, f.DeleteDataValidation("Sheet1", "A1:B2"), newCellNameToCoordinatesError("A", newInvalidCellNameError("A")).Error())

	// Test delete data validation on no exists worksheet
	assert.EqualError(t, f.DeleteDataValidation("SheetN", "A1:B2"), "sheet SheetN does not exist")
	// Test delete all data validation with invalid sheet name
	assert.EqualError(t, f.DeleteDataValidation("Sheet:1"), ErrSheetNameInvalid.Error())
	// Test delete all data validations in the worksheet
	assert.NoError(t, f.DeleteDataValidation("Sheet1"))
	assert.Nil(t, ws.(*xlsxWorksheet).DataValidations)

	t.Run("delete_data_validation_from_extLst", func(t *testing.T) {
		f := NewFile()
		f.Sheet.Delete("xl/worksheets/sheet1.xml")
		f.Pkg.Store("xl/worksheets/sheet1.xml", fmt.Appendf(nil,
			`<worksheet xmlns="%s"><sheetData/><extLst><ext xmlns:x14="%s" uri="%s"><x14:dataValidations xmlns:xm="%s" count="2"><x14:dataValidation allowBlank="true" showErrorMessage="true" showInputMessage="true" sqref="" type="list"><xm:sqref>A1:A2</xm:sqref><x14:formula1><xm:f>Sheet1!$A$2:$A$4</xm:f></x14:formula1></x14:dataValidation><x14:dataValidation allowBlank="true" showErrorMessage="true" showInputMessage="true" sqref="" type="list"><xm:sqref>B1:B2</xm:sqref><x14:formula1><xm:f>Sheet1!$B$2:$B$3</xm:f></x14:formula1></x14:dataValidation></x14:dataValidations></ext></extLst></worksheet>`,
			NameSpaceSpreadSheet.Value, NameSpaceSpreadSheetExcel2006Main.Value,
			ExtURIDataValidations, NameSpaceSpreadSheetExcel2006Main.Value))
		f.checked = sync.Map{}
		assert.NoError(t, f.DeleteDataValidation("Sheet1", "A1:A2"))
		dvs, err := f.GetDataValidations("Sheet1")
		assert.NoError(t, err)
		assert.Len(t, dvs, 1)
		assert.Equal(t, "B1:B2", dvs[0].Sqref)

		assert.NoError(t, f.DeleteDataValidation("Sheet1", "B1:B2"))
		dvs, err = f.GetDataValidations("Sheet1")
		assert.NoError(t, err)
		assert.Empty(t, dvs)
	})

	t.Run("delete_data_validation_failed_from_extLst", func(t *testing.T) {
		f := NewFile()
		assert.EqualError(t, f.deleteX14DataValidation(&xlsxWorksheet{
			ExtLst: &xlsxExtLst{Ext: "<extLst><x14:dataValidations></x14:dataValidation></x14:dataValidations></ext></extLst>"},
		}, nil), "XML syntax error on line 1: element <dataValidations> closed by </dataValidation>")
	})
}
