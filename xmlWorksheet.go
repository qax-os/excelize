// Some code of this file reference tealeg/xlsx.

package excelize

import "encoding/xml"

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
	Hyperlinks    xlsxHyperlinks    `xml:"hyperlinks"`
	MergeCells    *xlsxMergeCells   `xml:"mergeCells,omitempty"`
	PrintOptions  xlsxPrintOptions  `xml:"printOptions"`
	PageMargins   xlsxPageMargins   `xml:"pageMargins"`
	PageSetUp     xlsxPageSetUp     `xml:"pageSetup"`
	HeaderFooter  xlsxHeaderFooter  `xml:"headerFooter"`
	Drawing       xlsxDrawing       `xml:"drawing"`
	LegacyDrawing xlsxLegacyDrawing `xml:"legacyDrawing"`
	Picture       xlsxPicture       `xml:"picture"`
	TableParts    xlsxTableParts    `xml:"tableParts"`
}

// xlsxDrawing change r:id to rid in the namespace.
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
//
// A single sheet view definition. When more than one sheet view is
// defined in the file, it means that when opening the workbook, each
// sheet view corresponds to a separate window within the spreadsheet
// application, where each window is showing the particular sheet
// containing the same workbookViewId value, the last sheetView
// definition is loaded, and the others are discarded. When multiple
// windows are viewing the same sheet, multiple sheetView elements
// (with corresponding workbookView entries) are saved.
type xlsxSheetView struct {
	WindowProtection        bool            `xml:"windowProtection,attr,omitempty"`
	ShowFormulas            bool            `xml:"showFormulas,attr,omitempty"`
	ShowGridLines           string          `xml:"showGridLines,attr,omitempty"`
	ShowRowColHeaders       bool            `xml:"showRowColHeaders,attr,omitempty"`
	ShowZeros               bool            `xml:"showZeros,attr,omitempty"`
	RightToLeft             bool            `xml:"rightToLeft,attr,omitempty"`
	TabSelected             bool            `xml:"tabSelected,attr,omitempty"`
	ShowOutlineSymbols      bool            `xml:"showOutlineSymbols,attr,omitempty"`
	DefaultGridColor        bool            `xml:"defaultGridColor,attr"`
	View                    string          `xml:"view,attr,omitempty"`
	TopLeftCell             string          `xml:"topLeftCell,attr,omitempty"`
	ColorId                 int             `xml:"colorId,attr,omitempty"`
	ZoomScale               float64         `xml:"zoomScale,attr,omitempty"`
	ZoomScaleNormal         float64         `xml:"zoomScaleNormal,attr,omitempty"`
	ZoomScalePageLayoutView float64         `xml:"zoomScalePageLayoutView,attr,omitempty"`
	WorkbookViewID          int             `xml:"workbookViewId,attr"`
	Pane                    *xlsxPane       `xml:"pane,omitempty"`
	Selection               []xlsxSelection `xml:"selection"`
}

// xlsxSelection directly maps the selection element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxSelection struct {
	Pane         string `xml:"pane,attr,omitempty"`
	ActiveCell   string `xml:"activeCell,attr,omitempty"`
	ActiveCellID int    `xml:"activeCellId,attr"`
	SQRef        string `xml:"sqref,attr,omitempty"`
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
	XMLName                           xml.Name        `xml:"sheetPr"`
	FilterMode                        bool            `xml:"filterMode,attr,omitempty"`
	CodeName                          string          `xml:"codeName,attr,omitempty"`
	EnableFormatConditionsCalculation int             `xml:"enableFormatConditionsCalculation,attr,omitempty"`
	TabColor                          xlsxTabColor    `xml:"tabColor,omitempty"`
	PageSetUpPr                       xlsxPageSetUpPr `xml:"pageSetUpPr"`
}

// xlsxPageSetUpPr directly maps the pageSetupPr element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxPageSetUpPr struct {
	FitToPage bool `xml:"fitToPage,attr"` // Flag indicating whether the Fit to Page print option is enabled.
}

// xlsxTabColor directly maps the tabColor element in the namespace
// currently I have not checked it for completeness - it does as much
// as I need.
type xlsxTabColor struct {
	Theme int   `xml:"theme,attr,omitempty"` //  (Theme Color) A zero-based index into the <clrScheme> collection (§20.1.6.2), referencing a particular <sysClr> or <srgbClr> value expressed in the Theme part.
	Tint  uint8 `xml:"tint,attr,omitempty"`  // Specifies the tint value applied to the color.
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
	Si      string `xml:"si,attr,omitempty"`  // Shared formula index
}

// xlsxHyperlinks directly maps the hyperlinks element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main
type xlsxHyperlinks struct {
	Hyperlink []xlsxHyperlink `xml:"hyperlink"`
}

// xlsxHyperlink directly maps the hyperlink element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main
type xlsxHyperlink struct {
	Ref      string `xml:"ref,attr"`
	Location string `xml:"location,attr,omitempty"`
	Display  string `xml:"display,attr,omitempty"`
	RID      string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
}

// xlsxTableParts directly maps the tableParts element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// The table element has several attributes applied to identify the table
// and the data range it covers. The table id attribute needs to be unique
// across all table parts, the same goes for the name and displayName. The
// displayName has the further restriction that it must be unique across
// all defined names in the workbook. Later on we will see that you can
// define names for many elements, such as cells or formulas. The name
// value is used for the object model in Microsoft Office Excel. The
// displayName is used for references in formulas. The ref attribute is
// used to identify the cell range that the table covers. This includes
// not only the table data, but also the table header containing column
// names.
// To add columns to your table you add new tableColumn elements to the
// tableColumns container. Similar to the shared string table the
// collection keeps a count attribute identifying the number of columns.
// Besides the table definition in the table part there is also the need
// to identify which tables are displayed in the worksheet. The worksheet
// part has a separate element tableParts to store this information. Each
// table part is referenced through the relationship ID and again a count
// of the number of table parts is maintained. The following markup sample
// is taken from the documents accompanying this book. The sheet data
// element has been removed to reduce the size of the sample. To reference
// the table, just add the tableParts element, of course after having
// created and stored the table part.
// Example:
//
//    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
//        ...
//        <tableParts count="1">
// 		      <tablePart r:id="rId1" />
//        </tableParts>
//    </worksheet>
//
type xlsxTableParts struct {
	Count      int             `xml:"count,attr"`
	TableParts []xlsxTablePart `xml:"tablePart"`
}

// xlsxTablePart directly maps the tablePart element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main
type xlsxTablePart struct {
	RID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
}

// xlsxPicture directly maps the picture element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// Background sheet image.
// Example:
//
//    <picture r:id="rId1"/>
//
type xlsxPicture struct {
	RID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"` // Relationship Id pointing to the image part.
}

// xlsxLegacyDrawing directly maps the legacyDrawing element in the namespace
// http://schemas.openxmlformats.org/spreadsheetml/2006/main -
// A comment is a rich text note that is attached to, and associated with,
// a cell, separate from other cell content. Comment content is stored
// separate from the cell, and is displayed in a drawing object (like a
// text box) that is separate from, but associated with, a cell. Comments
// are used as reminders, such as noting how a complex formula works, or
// to provide feedback to other users. Comments can also be used to explain
// assumptions made in a formula or to call out something special about the cell.
type xlsxLegacyDrawing struct {
	RID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
}
