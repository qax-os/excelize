package excelize

import (
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestWorkbookProps(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetWorkbookProps(nil))
	wb, err := f.workbookReader()
	assert.NoError(t, err)
	wb.WorkbookPr = nil
	expected := WorkbookPropsOptions{
		Date1904:      boolPtr(true),
		FilterPrivacy: boolPtr(true),
		CodeName:      stringPtr("code"),
	}
	assert.NoError(t, f.SetWorkbookProps(&expected))
	opts, err := f.GetWorkbookProps()
	assert.NoError(t, err)
	assert.Equal(t, expected, opts)
	wb.WorkbookPr = nil
	opts, err = f.GetWorkbookProps()
	assert.NoError(t, err)
	assert.Equal(t, WorkbookPropsOptions{}, opts)
	// Test set workbook properties with unsupported charset workbook
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	assert.EqualError(t, f.SetWorkbookProps(&expected), "XML syntax error on line 1: invalid UTF-8")
	// Test get workbook properties with unsupported charset workbook
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	_, err = f.GetWorkbookProps()
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
}

func effectiveCalcProps() CalcPropsOptions {
	return CalcPropsOptions{
		CalcID:                uintPtr(calcPrDefaultCalcID),
		CalcMode:              stringPtr(calcPrDefaultCalcMode),
		FullCalcOnLoad:        boolPtr(calcPrDefaultFullCalcOnLoad),
		RefMode:               stringPtr(calcPrDefaultRefMode),
		Iterate:               boolPtr(calcPrDefaultIterate),
		IterateCount:          uintPtr(calcPrDefaultIterateCount),
		IterateDelta:          float64Ptr(calcPrDefaultIterateDelta),
		FullPrecision:         boolPtr(calcPrDefaultFullPrecision),
		CalcCompleted:         boolPtr(calcPrDefaultCalcCompleted),
		CalcOnSave:            boolPtr(calcPrDefaultCalcOnSave),
		ConcurrentCalc:        boolPtr(calcPrDefaultConcurrentCalc),
		ConcurrentManualCount: uintPtr(calcPrDefaultConcurrentManualCount),
		ForceFullCalc:         boolPtr(calcPrDefaultForceFullCalc),
	}
}

func workbookXML(t *testing.T, f *File) string {
	t.Helper()
	f.workBookWriter()
	v, ok := f.Pkg.Load(f.getWorkbookPath())
	assert.True(t, ok)
	return string(v.([]byte))
}

func TestCalcProps(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.SetCalcProps(nil))
	wb, err := f.workbookReader()
	assert.NoError(t, err)
	wb.CalcPr = nil
	assert.NoError(t, f.SetCalcProps(&CalcPropsOptions{
		FullCalcOnLoad:        boolPtr(true),
		CalcID:                uintPtr(122211),
		ConcurrentManualCount: uintPtr(5),
		IterateCount:          uintPtr(10),
		ConcurrentCalc:        boolPtr(true),
	}))
	expected := effectiveCalcProps()
	expected.FullCalcOnLoad = boolPtr(true)
	expected.CalcID = uintPtr(122211)
	expected.ConcurrentManualCount = uintPtr(5)
	expected.IterateCount = uintPtr(10)
	opts, err := f.GetCalcProps()
	assert.NoError(t, err)
	assert.Equal(t, expected, opts)

	wb.CalcPr = nil
	opts, err = f.GetCalcProps()
	assert.NoError(t, err)
	assert.Equal(t, effectiveCalcProps(), opts)
	// Test set calculation properties with unsupported optional value
	assert.Equal(t, newInvalidOptionalValue("CalcMode", "AUTO", supportedCalcMode), f.SetCalcProps(&CalcPropsOptions{CalcMode: stringPtr("AUTO")}))
	assert.Equal(t, newInvalidOptionalValue("RefMode", "a1", supportedRefMode), f.SetCalcProps(&CalcPropsOptions{RefMode: stringPtr("a1")}))
	// Test set calculation properties with unsupported charset workbook
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	assert.EqualError(t, f.SetCalcProps(&CalcPropsOptions{CalcID: uintPtr(1)}), "XML syntax error on line 1: invalid UTF-8")
	// Test get calculation properties with unsupported charset workbook
	f.WorkBook = nil
	f.Pkg.Store(defaultXMLPathWorkbook, MacintoshCyrillicCharset)
	_, err = f.GetCalcProps()
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
}

func TestCalcPropsOmitsSpecDefaults(t *testing.T) {
	for _, tc := range []struct {
		name    string
		opts    CalcPropsOptions
		attr    string
		written bool
	}{
		{"FullPrecision default", CalcPropsOptions{FullPrecision: boolPtr(true)}, "fullPrecision", false},
		{"FullPrecision changed", CalcPropsOptions{FullPrecision: boolPtr(false)}, "fullPrecision", true},
		{"CalcCompleted default", CalcPropsOptions{CalcCompleted: boolPtr(true)}, "calcCompleted", false},
		{"CalcCompleted changed", CalcPropsOptions{CalcCompleted: boolPtr(false)}, "calcCompleted", true},
		{"CalcOnSave default", CalcPropsOptions{CalcOnSave: boolPtr(true)}, "calcOnSave", false},
		{"CalcOnSave changed", CalcPropsOptions{CalcOnSave: boolPtr(false)}, "calcOnSave", true},
		{"ConcurrentCalc default", CalcPropsOptions{ConcurrentCalc: boolPtr(true)}, "concurrentCalc", false},
		{"ConcurrentCalc changed", CalcPropsOptions{ConcurrentCalc: boolPtr(false)}, "concurrentCalc", true},
		{"ForceFullCalc default", CalcPropsOptions{ForceFullCalc: boolPtr(false)}, "forceFullCalc", false},
		{"ForceFullCalc changed", CalcPropsOptions{ForceFullCalc: boolPtr(true)}, "forceFullCalc", true},
		{"FullCalcOnLoad default", CalcPropsOptions{FullCalcOnLoad: boolPtr(false)}, "fullCalcOnLoad", false},
		{"FullCalcOnLoad changed", CalcPropsOptions{FullCalcOnLoad: boolPtr(true)}, "fullCalcOnLoad", true},
		{"Iterate default", CalcPropsOptions{Iterate: boolPtr(false)}, "iterate", false},
		{"Iterate changed", CalcPropsOptions{Iterate: boolPtr(true)}, "iterate", true},
		{"IterateCount default", CalcPropsOptions{IterateCount: uintPtr(100)}, "iterateCount", false},
		{"IterateCount changed", CalcPropsOptions{IterateCount: uintPtr(10)}, "iterateCount", true},
		{"IterateDelta default", CalcPropsOptions{IterateDelta: float64Ptr(0.001)}, "iterateDelta", false},
		{"IterateDelta changed", CalcPropsOptions{IterateDelta: float64Ptr(0.5)}, "iterateDelta", true},
		{"CalcMode default", CalcPropsOptions{CalcMode: stringPtr("auto")}, "calcMode", false},
		{"CalcMode changed", CalcPropsOptions{CalcMode: stringPtr("manual")}, "calcMode", true},
		{"RefMode default", CalcPropsOptions{RefMode: stringPtr("A1")}, "refMode", false},
		{"RefMode changed", CalcPropsOptions{RefMode: stringPtr("R1C1")}, "refMode", true},
	} {
		t.Run(tc.name, func(t *testing.T) {
			f := NewFile()
			defer func() { assert.NoError(t, f.Close()) }()
			assert.NoError(t, f.SetCalcProps(&tc.opts))
			assert.Equal(t, tc.written, strings.Contains(workbookXML(t, f), tc.attr+`="`))
		})
	}
}

func TestCalcPropsRoundTrip(t *testing.T) {
	for _, tc := range []struct {
		name string
		opts CalcPropsOptions
		want func(*CalcPropsOptions)
	}{
		{"FullPrecision false", CalcPropsOptions{FullPrecision: boolPtr(false)}, func(o *CalcPropsOptions) { o.FullPrecision = boolPtr(false) }},
		{"FullPrecision true", CalcPropsOptions{FullPrecision: boolPtr(true)}, func(o *CalcPropsOptions) {}},
		{"CalcCompleted false", CalcPropsOptions{CalcCompleted: boolPtr(false)}, func(o *CalcPropsOptions) { o.CalcCompleted = boolPtr(false) }},
		{"CalcOnSave false", CalcPropsOptions{CalcOnSave: boolPtr(false)}, func(o *CalcPropsOptions) { o.CalcOnSave = boolPtr(false) }},
		{"ConcurrentCalc false", CalcPropsOptions{ConcurrentCalc: boolPtr(false)}, func(o *CalcPropsOptions) { o.ConcurrentCalc = boolPtr(false) }},
		{"ForceFullCalc true", CalcPropsOptions{ForceFullCalc: boolPtr(true)}, func(o *CalcPropsOptions) { o.ForceFullCalc = boolPtr(true) }},
		{"FullCalcOnLoad true", CalcPropsOptions{FullCalcOnLoad: boolPtr(true)}, func(o *CalcPropsOptions) { o.FullCalcOnLoad = boolPtr(true) }},
		{"Iterate true", CalcPropsOptions{Iterate: boolPtr(true)}, func(o *CalcPropsOptions) { o.Iterate = boolPtr(true) }},
		{"IterateCount zero", CalcPropsOptions{IterateCount: uintPtr(0)}, func(o *CalcPropsOptions) { o.IterateCount = uintPtr(0) }},
		{"IterateDelta zero", CalcPropsOptions{IterateDelta: float64Ptr(0)}, func(o *CalcPropsOptions) { o.IterateDelta = float64Ptr(0) }},
		{"CalcMode manual", CalcPropsOptions{CalcMode: stringPtr("manual")}, func(o *CalcPropsOptions) { o.CalcMode = stringPtr("manual") }},
		{"RefMode R1C1", CalcPropsOptions{RefMode: stringPtr("R1C1")}, func(o *CalcPropsOptions) { o.RefMode = stringPtr("R1C1") }},
	} {
		t.Run(tc.name, func(t *testing.T) {
			f := NewFile()
			assert.NoError(t, f.SetCalcProps(&tc.opts))
			buf, err := f.WriteToBuffer()
			assert.NoError(t, err)
			assert.NoError(t, f.Close())

			reopened, err := OpenReader(buf)
			assert.NoError(t, err)
			defer func() { assert.NoError(t, reopened.Close()) }()
			got, err := reopened.GetCalcProps()
			assert.NoError(t, err)

			want := effectiveCalcProps()
			want.CalcID = uintPtr(122211)
			tc.want(&want)
			assert.Equal(t, want, got)
		})
	}
}

func TestDeleteWorkbookRels(t *testing.T) {
	f := NewFile()
	// Test delete pivot table without worksheet relationships
	f.Relationships.Delete("xl/_rels/workbook.xml.rels")
	f.Pkg.Delete("xl/_rels/workbook.xml.rels")
	rID, err := f.deleteWorkbookRels("", "")
	assert.Empty(t, rID)
	assert.NoError(t, err)
}
