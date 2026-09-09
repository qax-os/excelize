package excelize

import (
	"archive/zip"
	"bytes"
	"io"
	"strings"
	"testing"
)

// Regression test: worksheet normalization (checkSheet -> checkSheetR0) must
// not panic on attacker-controlled cell references whose column name
// overflows the int64 accumulator in ColumnNameToNumber.
//
// A 14-letter column name ("VGWQHXLSDVIKWV", bijective base-26 value 3*2^64)
// previously wrapped the accumulator to 0, passing the `col > MaxColumns`
// gate with a nil error. CellNameToCoordinates then returned col=0, and the
// normalization path indexed the row's cell slice with col-1 = -1, panicking
// out of the public OpenReader/GetCellValue API.
//
// Expected behavior: the out-of-domain cell is rejected during parsing;
// OpenReader either rejects the file or opens it, and GetCellValue returns
// normally without panicking.
func TestCheckSheetColumnNameOverflowPanic(t *testing.T) {
	base := NewFile()
	defer func() { _ = base.Close() }()
	buf, err := base.WriteToBuffer()
	if err != nil {
		t.Fatalf("build base workbook: %v", err)
	}

	craftedSheet := strings.Replace(templateSheet,
		"<sheetData/>",
		`<sheetData><row r="0"><c r="VGWQHXLSDVIKWV1" t="inlineStr"><is><t>pwn</t></is></c></row></sheetData>`,
		1)

	out := new(bytes.Buffer)
	zw := zip.NewWriter(out)
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatalf("reopen base workbook: %v", err)
	}
	for _, zf := range zr.File {
		rc, err := zf.Open()
		if err != nil {
			t.Fatalf("open %s: %v", zf.Name, err)
		}
		data, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			t.Fatalf("read %s: %v", zf.Name, err)
		}
		if zf.Name == "xl/worksheets/sheet1.xml" {
			data = []byte(craftedSheet)
		}
		w, err := zw.Create(zf.Name)
		if err != nil {
			t.Fatalf("create %s: %v", zf.Name, err)
		}
		if _, err := w.Write(data); err != nil {
			t.Fatalf("write %s: %v", zf.Name, err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatalf("close zip writer: %v", err)
	}

	f, err := OpenReader(bytes.NewReader(out.Bytes()))
	if err != nil {
		// Rejected at open time is an acceptable fix behavior.
		return
	}
	defer func() { _ = f.Close() }()

	defer func() {
		if r := recover(); r != nil {
			t.Fatalf("worksheet normalization panicked on attacker-controlled cell reference: %v", r)
		}
	}()
	_, _ = f.GetCellValue("Sheet1", "A1")

	// The coordinate gate itself must reject the overflowing name in every
	// case: the in-loop bound check catches the overflow before the int64
	// accumulator can wrap (including wraps that would land back inside the
	// valid domain).
	if col, err := ColumnNameToNumber("VGWQHXLSDVIKWV"); err == nil {
		t.Fatalf("ColumnNameToNumber accepted an overflowing column name: col=%d, err=nil", col)
	}
}
