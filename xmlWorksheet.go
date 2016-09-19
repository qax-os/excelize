// Some code of this file reference tealeg/xlsx

package excelize

import (
	"encoding/xml"
)

// xlsxWorksheet directly maps the worksheet element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxWorksheet struct {
	XMLName       xml.Name          `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main worksheet"`
	SheetPr       xlsxSheetPr       `xml:"sheetPr"`
	Dimension     xlsxDimension     `xml:"dimension"`
	SheetViews    xlsxSheetViews    `xml:"sheetViews"`
	SheetFormatPr xlsxSheetFormatPr `xml:"sheetFormatPr"`
	Cols          *xlsxCols         `xml:"cols,omitempty"`
	SheetData     xlsxSheetData     `xml:"sheetData"`
	MergeCells    *xlsxMergeCells   `xml:"mergeCells,omitempty"`
	PrintOptions  xlsxPrintOptions  `xml:"printOptions"`
	PageMargins   xlsxPageMargins   `xml:"pageMargins"`
	PageSetUp     xlsxPageSetUp     `xml:"pageSetup"`
	HeaderFooter  xlsxHeaderFooter  `xml:"headerFooter"`
	Drawing       xlsxDrawing       `xml:"drawing"`
}

// xlsxDrawing change r:id to rid in the namespace
type xlsxDrawing struct {
	RID string `xml:"rid,attr"`
}

// xlsxHeaderFooter directly maps the headerFooter element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxHeaderFooter struct {
	DifferentFirst   bool            `xml:"differentFirst,attr"`
	DifferentOddEven bool            `xml:"differentOddEven,attr"`
	OddHeader        []xlsxOddHeader `xml:"oddHeader"`
	OddFooter        []xlsxOddFooter `xml:"oddFooter"`
}

// xlsxOddHeader directly maps the oddHeader element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxOddHeader struct {
	Content string `xml:",chardata"`
}

// xlsxOddFooter directly maps the oddFooter element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxOddFooter struct {
	Content string `xml:",chardata"`
}

// xlsxPageSetUp directly maps the pageSetup element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxPageSetUp struct {
	PaperSize          string  `xml:"paperSize,attr,omitempty"`
	Scale              int     `xml:"scale,attr"`
	FirstPageNumber    int     `xml:"firstPageNumber,attr"`
	FitToWidth         int     `xml:"fitToWidth,attr"`
	FitToHeight        int     `xml:"fitToHeight,attr"`
	PageOrder          string  `xml:"pageOrder,attr,omitempty"`
	Orientation        string  `xml:"orientation,attr,omitempty"`
	UsePrinterDefaults bool    `xml:"usePrinterDefaults,attr"`
	BlackAndWhite      bool    `xml:"blackAndWhite,attr"`
	Draft              bool    `xml:"draft,attr"`
	CellComments       string  `xml:"cellComments,attr,omitempty"`
	UseFirstPageNumber bool    `xml:"useFirstPageNumber,attr"`
	HorizontalDPI      float32 `xml:"horizontalDpi,attr"`
	VerticalDPI        float32 `xml:"verticalDpi,attr"`
	Copies             int     `xml:"copies,attr"`
}

// xlsxPrintOptions directly maps the printOptions element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxPrintOptions struct {
	Headings           bool `xml:"headings,attr"`
	GridLines          bool `xml:"gridLines,attr"`
	GridLinesSet       bool `xml:"gridLinesSet,attr"`
	HorizontalCentered bool `xml:"horizontalCentered,attr"`
	VerticalCentered   bool `xml:"verticalCentered,attr"`
}

// xlsxPageMargins directly maps the pageMargins element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxPageMargins struct {
	Left   float64 `xml:"left,attr"`
	Right  float64 `xml:"right,attr"`
	Top    float64 `xml:"top,attr"`
	Bottom float64 `xml:"bottom,attr"`
	Header float64 `xml:"header,attr"`
	Footer float64 `xml:"footer,attr"`
}

// xlsxSheetFormatPr directly maps the sheetFormatPr element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSheetFormatPr struct {
	DefaultColWidth  float64 `xml:"defaultColWidth,attr,omitempty"`
	DefaultRowHeight float64 `xml:"defaultRowHeight,attr"`
	CustomHeight     float64 `xml:"customHeight,attr,omitempty"`
	ZeroHeight       float64 `xml:"zeroHeight,attr,omitempty"`
	OutlineLevelCol  uint8   `xml:"outlineLevelCol,attr,omitempty"`
	OutlineLevelRow  uint8   `xml:"outlineLevelRow,attr,omitempty"`
}

// xlsxSheetViews directly maps the sheetViews element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSheetViews struct {
	SheetView []xlsxSheetView `xml:"sheetView"`
}

// xlsxSheetView directly maps the sheetView element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSheetView struct {
	// WindowProtection        bool            `xml:"windowProtection,attr"`
	// ShowFormulas            bool            `xml:"showFormulas,attr"`
	ShowGridLines string `xml:"showGridLines,attr,omitempty"`
	// ShowRowColHeaders       bool            `xml:"showRowColHeaders,attr"`
	// ShowZeros               bool            `xml:"showZeros,attr"`
	// RightToLeft             bool            `xml:"rightToLeft,attr"`
	TabSelected bool `xml:"tabSelected,attr"`
	// ShowOutlineSymbols      bool            `xml:"showOutlineSymbols,attr"`
	// DefaultGridColor        bool            `xml:"defaultGridColor,attr"`
	// View                    string          `xml:"view,attr"`
	TopLeftCell string `xml:"topLeftCell,attr,omitempty"`
	// ColorId                 int             `xml:"colorId,attr"`
	// ZoomScale               float64         `xml:"zoomScale,attr"`
	// ZoomScaleNormal         float64         `xml:"zoomScaleNormal,attr"`
	// ZoomScalePageLayoutView float64         `xml:"zoomScalePageLayoutView,attr"`
	WorkbookViewID int             `xml:"workbookViewId,attr"`
	Selection      []xlsxSelection `xml:"selection"`
	Pane           *xlsxPane       `xml:"pane,omitempty"`
}

// xlsxSelection directly maps the selection element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSelection struct {
	Pane         string `xml:"pane,attr,omitempty"`
	ActiveCell   string `xml:"activeCell,attr,omitempty"`
	ActiveCellID int    `xml:"activeCellId,attr"`
	SQRef        string `xml:"sqref,attr"`
}

// xlsxSelection directly maps the selection element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxPane struct {
	XSplit      float64 `xml:"xSplit,attr"`
	YSplit      float64 `xml:"ySplit,attr"`
	TopLeftCell string  `xml:"topLeftCell,attr,omitempty"`
	ActivePane  string  `xml:"activePane,attr,omitempty"`
	State       string  `xml:"state,attr,omitempty"` // Either "split" or "frozen"
}

// xlsxSheetPr directly maps the sheetPr element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSheetPr struct {
	FilterMode  bool              `xml:"filterMode,attr"`
	PageSetUpPr []xlsxPageSetUpPr `xml:"pageSetUpPr"`
}

// xlsxPageSetUpPr directly maps the pageSetupPr element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxPageSetUpPr struct {
	FitToPage bool `xml:"fitToPage,attr"`
}

// xlsxCols directly maps the cols element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxCols struct {
	Col []xlsxCol `xml:"col"`
}

// xlsxCol directly maps the col element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxCol struct {
	Collapsed    bool    `xml:"collapsed,attr"`
	Hidden       bool    `xml:"hidden,attr"`
	Max          int     `xml:"max,attr"`
	Min          int     `xml:"min,attr"`
	Style        int     `xml:"style,attr"`
	Width        float64 `xml:"width,attr"`
	CustomWidth  int     `xml:"customWidth,attr,omitempty"`
	OutlineLevel uint8   `xml:"outlineLevel,attr,omitempty"`
}

// xlsxDimension directly maps the dimension element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxDimension struct {
	Ref string `xml:"ref,attr"`
}

// xlsxSheetData directly maps the sheetData element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSheetData struct {
	XMLName xml.Name  `xml:"sheetData"`
	Row     []xlsxRow `xml:"row"`
}

// xlsxRow directly maps the row element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxRow struct {
	R            int     `xml:"r,attr"`
	Spans        string  `xml:"spans,attr,omitempty"`
	Hidden       bool    `xml:"hidden,attr,omitempty"`
	C            []xlsxC `xml:"c"`
	Ht           string  `xml:"ht,attr,omitempty"`
	CustomHeight bool    `xml:"customHeight,attr,omitempty"`
	OutlineLevel uint8   `xml:"outlineLevel,attr,omitempty"`
}

type xlsxMergeCell struct {
	Ref string `xml:"ref,attr"` // ref: horiz "A1:C1", vert "B3:B6", both  "D3:G4"
}

type xlsxMergeCells struct {
	XMLName xml.Name        //`xml:"mergeCells,omitempty"`
	Count   int             `xml:"count,attr,omitempty"`
	Cells   []xlsxMergeCell `xml:"mergeCell,omitempty"`
}

// xlsxC directly maps the c element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxC struct {
	R string `xml:"r,attr"`           // Cell ID, e.g. A1
	S int    `xml:"s,attr,omitempty"` // Style reference.
	// Str string `xml:"str,attr,omitempty"` // Style reference.
	T string `xml:"t,attr,omitempty"` // Type.
	F *xlsxF `xml:"f,omitempty"`      // Formula
	V string `xml:"v,omitempty"`      // Value
}

// xlsxF directly maps the f element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxF struct {
	Content string `xml:",chardata"`
	T       string `xml:"t,attr,omitempty"`   // Formula type
	Ref     string `xml:"ref,attr,omitempty"` // Shared formula ref
	Si      int    `xml:"si,attr,omitempty"`  // Shared formula index
}
