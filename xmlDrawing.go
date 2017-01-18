package excelize

// Source relationship and namespace.
const (
	SourceRelationship          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	SourceRelationshipImage     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
	SourceRelationshipDrawingML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
	SourceRelationshipWorkSheet = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
	NameSpaceDrawingML          = "http://schemas.openxmlformats.org/drawingml/2006/main"
	NameSpaceSpreadSheetDrawing = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
)

// xlsxCNvPr directly maps the cNvPr (Non-Visual Drawing Properties).
// This element specifies non-visual canvas properties. This allows for
// additional information that does not affect the appearance of the
// picture to be stored.
type xlsxCNvPr struct {
	ID    int    `xml:"id,attr"`
	Name  string `xml:"name,attr"`
	Descr string `xml:"descr,attr"`
	Title string `xml:"title,attr,omitempty"`
}

// xlsxPicLocks directly maps the picLocks (Picture Locks). This element
// specifies all locking properties for a graphic frame. These properties
// inform the generating application about specific properties that have
// been previously locked and thus should not be changed.
type xlsxPicLocks struct {
	NoAdjustHandles    bool `xml:"noAdjustHandles,attr,omitempty"`
	NoChangeArrowheads bool `xml:"noChangeArrowheads,attr,omitempty"`
	NoChangeAspect     bool `xml:"noChangeAspect,attr"`
	NoChangeShapeType  bool `xml:"noChangeShapeType,attr,omitempty"`
	NoCrop             bool `xml:"noCrop,attr,omitempty"`
	NoEditPoints       bool `xml:"noEditPoints,attr,omitempty"`
	NoGrp              bool `xml:"noGrp,attr,omitempty"`
	NoMove             bool `xml:"noMove,attr,omitempty"`
	NoResize           bool `xml:"noResize,attr,omitempty"`
	NoRot              bool `xml:"noRot,attr,omitempty"`
	NoSelect           bool `xml:"noSelect,attr,omitempty"`
}

// xlsxBlip directly maps the blip element in the namespace
// http://purl.oclc.or g/ooxml/officeDoc ument/relationships -
// This element specifies the existence of an image (binary large image or
// picture) and contains a reference to the image data.
type xlsxBlip struct {
	Embed  string `xml:"r:embed,attr"`
	Cstate string `xml:"cstate,attr,omitempty"`
	R      string `xml:"xmlns:r,attr"`
}

// xlsxStretch directly maps the stretch element. This element specifies
// that a BLIP should be stretched to fill the target rectangle. The other
// option is a tile where a BLIP is tiled to fill the available area.
type xlsxStretch struct {
	FillRect string `xml:"a:fillRect"`
}

// xlsxOff directly maps the colOff and rowOff element. This element is used
// to specify the column offset within a cell.
type xlsxOff struct {
	X int `xml:"x,attr"`
	Y int `xml:"y,attr"`
}

// xlsxExt directly maps the ext element.
type xlsxExt struct {
	Cx int `xml:"cx,attr"`
	Cy int `xml:"cy,attr"`
}

// xlsxPrstGeom directly maps the prstGeom (Preset geometry). This element specifies
// when a preset geometric shape should be used instead of a custom geometric shape.
// The generating application should be able to render all preset geometries enumerated
// in the ST_ShapeType list.
type xlsxPrstGeom struct {
	Prst string `xml:"prst,attr"`
}

// xlsxXfrm directly maps the xfrm (2D Transform for Graphic Frame). This element
// specifies the transform to be applied to the corresponding graphic frame. This
// transformation is applied to the graphic frame just as it would be for a shape
// or group shape.
type xlsxXfrm struct {
	Off xlsxOff `xml:"a:off"`
	Ext xlsxExt `xml:"a:ext"`
}

// xlsxCNvPicPr directly maps the cNvPicPr (Non-Visual Picture Drawing Properties).
// This element specifies the non-visual properties for the picture canvas. These
// properties are to be used by the generating application to determine how certain
// properties are to be changed for the picture object in question.
type xlsxCNvPicPr struct {
	PicLocks xlsxPicLocks `xml:"a:picLocks"`
}

// directly maps the nvPicPr (Non-Visual Properties for a Picture). This element specifies
// all non-visual properties for a picture. This element is a container for the non-visual
// identification properties, shape properties and application properties that are to be
// associated with a picture. This allows for additional information that does not affect
// the appearance of the picture to be stored.
type xlsxNvPicPr struct {
	CNvPr    xlsxCNvPr    `xml:"xdr:cNvPr"`
	CNvPicPr xlsxCNvPicPr `xml:"xdr:cNvPicPr"`
}

// xlsxBlipFill directly maps the blipFill (Picture Fill). This element specifies the kind
// of picture fill that the picture object has. Because a picture has a picture fill already
// by default, it is possible to have two fills specified for a picture object.
type xlsxBlipFill struct {
	Blip    xlsxBlip    `xml:"a:blip"`
	Stretch xlsxStretch `xml:"a:stretch"`
}

// xlsxSpPr directly maps the spPr (Shape Properties). This element specifies the visual shape
// properties that can be applied to a picture. These are the same properties that are allowed
// to describe the visual properties of a shape but are used here to describe the visual
// appearance of a picture within a document.
type xlsxSpPr struct {
	Xfrm     xlsxXfrm     `xml:"a:xfrm"`
	PrstGeom xlsxPrstGeom `xml:"a:prstGeom"`
}

// xlsxPic elements encompass the definition of pictures within the DrawingML framework. While
// pictures are in many ways very similar to shapes they have specific properties that are unique
// in order to optimize for picture- specific scenarios.
type xlsxPic struct {
	NvPicPr  xlsxNvPicPr  `xml:"xdr:nvPicPr"`
	BlipFill xlsxBlipFill `xml:"xdr:blipFill"`
	SpPr     xlsxSpPr     `xml:"xdr:spPr"`
}

// xlsxFrom specifies the starting anchor.
type xlsxFrom struct {
	Col    int `xml:"xdr:col"`
	ColOff int `xml:"xdr:colOff"`
	Row    int `xml:"xdr:row"`
	RowOff int `xml:"xdr:rowOff"`
}

// xlsxTo directly specifies the ending anchor.
type xlsxTo struct {
	Col    int `xml:"xdr:col"`
	ColOff int `xml:"xdr:colOff"`
	Row    int `xml:"xdr:row"`
	RowOff int `xml:"xdr:rowOff"`
}

// xlsxClientData directly maps the clientData element. An empty element which specifies (via
// attributes) certain properties related to printing and selection of the drawing object. The
// fLocksWithSheet attribute (either true or false) determines whether to disable selection when
// the sheet is protected, and fPrintsWithSheet attribute (either true or false) determines whether
// the object is printed when the sheet is printed.
type xlsxClientData struct {
	FLocksWithSheet  bool `xml:"fLocksWithSheet,attr"`
	FPrintsWithSheet bool `xml:"fPrintsWithSheet,attr"`
}

// xlsxTwoCellAnchor directly maps the twoCellAnchor (Two Cell Anchor Shape Size). This element
// specifies a two cell anchor placeholder for a group, a shape, or a drawing element. It moves
// with cells and its extents are in EMU units.
type xlsxTwoCellAnchor struct {
	EditAs       string          `xml:"editAs,attr,omitempty"`
	From         *xlsxFrom       `xml:"xdr:from"`
	To           *xlsxTo         `xml:"xdr:to"`
	Pic          *xlsxPic        `xml:"xdr:pic,omitempty"`
	GraphicFrame string          `xml:",innerxml"`
	ClientData   *xlsxClientData `xml:"xdr:clientData"`
}

// xlsxWsDr directly maps the root element for a part of this content type shall wsDr.
type xlsxWsDr struct {
	TwoCellAnchor []*xlsxTwoCellAnchor `xml:"xdr:twoCellAnchor"`
	Xdr           string               `xml:"xmlns:xdr,attr"`
	A             string               `xml:"xmlns:a,attr"`
}

// encodeWsDr directly maps the element xdr:wsDr.
type encodeWsDr struct {
	WsDr xlsxWsDr `xml:"xdr:wsDr"`
}
