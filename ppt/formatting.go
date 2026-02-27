package ppt

// GeoVertex represents a vertex point in a freeform shape path.
type GeoVertex struct {
	X, Y int32
}

// GeoSegment represents a path segment command.
// SegType: 0=lineTo, 1=curveTo, 2=moveTo, 3=close, 4=end, 5=escape
type GeoSegment struct {
	SegType uint16
	Count   uint16 // number of points consumed
}

// ShapeFormatting represents the complete formatting information for a shape.
type ShapeFormatting struct {
	ShapeType  uint16 // Shape type ID (1=rectangle, 202=textbox, etc.)
	Left       int32  // EMU
	Top        int32  // EMU
	Width      int32  // EMU
	Height     int32  // EMU
	IsText     bool   // Whether this is a text shape
	IsImage    bool   // Whether this is an image shape
	ImageIdx   int    // Image index (-1 = no image)
	FillColor  string // 6-digit hex RGB, empty = no solid fill
	LineColor  string // 6-digit hex RGB, empty = no line
	NoFill     bool   // Shape has no fill
	NoLine     bool   // Shape has no line
	LineWidth  int32  // Line width in EMU (0 = default)
	LineDash   int32  // Line dash style (-1 = not set, 0=solid, 1=dash, etc.)
	Rotation   int32  // Rotation in 1/64000 of a degree (fixedPoint)
	FlipH      bool   // Horizontal flip
	FlipV      bool   // Vertical flip
	// Text body properties
	TextMarginLeft   int32 // Left text margin in EMU (-1 = default)
	TextMarginTop    int32 // Top text margin in EMU (-1 = default)
	TextMarginRight  int32 // Right text margin in EMU (-1 = default)
	TextMarginBottom int32 // Bottom text margin in EMU (-1 = default)
	TextAnchor       int32 // Vertical anchor: -1=default, 0=top, 1=middle, 2=bottom, 3=topCentered, 4=middleCentered, 5=bottomCentered
	TextWordWrap     int32 // Word wrap: -1=default, 0=none, 1=square
	FillOpacity      int32 // Fill opacity: 0-65536 (65536=fully opaque, -1=not set)
	FillColorRaw     uint32 // Raw fill color value (for scheme color detection)
	LineColorRaw     uint32 // Raw line color value (for scheme color detection)
	// Image cropping (1/65536 = 100%)
	CropFromTop    int32 // Crop from top edge
	CropFromBottom int32 // Crop from bottom edge
	CropFromLeft   int32 // Crop from left edge
	CropFromRight  int32 // Crop from right edge
	// Freeform geometry (type=0)
	GeoVertices  []GeoVertex  // Path vertices for freeform shapes
	GeoSegments  []GeoSegment // Path segment commands for freeform shapes
	GeoLeft      int32        // Geometry coordinate space left
	GeoTop       int32        // Geometry coordinate space top
	GeoRight     int32        // Geometry coordinate space right
	GeoBottom    int32        // Geometry coordinate space bottom
	// Line arrow properties (per MS-ODRAW lineStartArrowhead / lineEndArrowhead)
	LineStartArrowHead   int32 // 0=none, 1=triangle, 2=stealth, 3=diamond, 4=oval, 5=open
	LineEndArrowHead     int32 // 0=none, 1=triangle, 2=stealth, 3=diamond, 4=oval, 5=open
	LineStartArrowWidth  int32 // 0=narrow, 1=medium, 2=wide (-1=not set)
	LineStartArrowLength int32 // 0=short, 1=medium, 2=long (-1=not set)
	LineEndArrowWidth    int32 // 0=narrow, 1=medium, 2=wide (-1=not set)
	LineEndArrowLength   int32 // 0=short, 1=medium, 2=long (-1=not set)
	// TextType from TextHeaderAtom: 0=title, 1=body, 4=other, -1=unknown
	TextType         int
	Paragraphs       []SlideParagraph
}

// SlideParagraph represents a paragraph within a shape.
type SlideParagraph struct {
	Alignment   uint8  // 0=left, 1=center, 2=right, 3=justify
	IndentLevel uint8  // 0-4
	SpaceBefore int32  // centipoints or percentage (negative = centipoints)
	SpaceAfter  int32  // centipoints or percentage
	LineSpacing int32  // percentage*100 (positive) or centipoints (negative)
	HasBullet   bool
	BulletChar  string
	BulletColor string // 6-digit hex RGB, empty = inherit
	BulletSize  int16  // percentage of text size (0 = not set)
	BulletFont  string // bullet font name, empty = inherit from text
	LeftMargin  int32  // master units (1/576 inch)
	Indent      int32  // master units (first line indent, can be negative)
	Runs        []SlideTextRun
}

// SlideTextRun represents a run of text with character formatting.
type SlideTextRun struct {
	Text      string
	FontName  string
	FontSize  uint16 // centipoints (hundredths of a point)
	Bold      bool
	Italic    bool
	Underline bool
	Color     string // 6-digit hex RGB, empty string means no color
	ColorRaw  uint32 // Raw color value (for scheme color detection)
}

// SlideBackground represents the background fill of a slide.
type SlideBackground struct {
	HasBackground bool   // Whether a background fill was found
	FillColor     string // 6-digit hex RGB for solid fill, empty = no solid fill
	ImageIdx      int    // Image index for blip fill (-1 = no image)
	fillColorRaw  uint32 // Raw fill color value (for scheme color detection)
}


// MasterTextStyle holds default text properties for a single indent level.
type MasterTextStyle struct {
	FontSize uint16 // centipoints (hundredths of a point), 0 = not set
	FontName string // font name, empty = not set
	Bold     bool
	Italic   bool
	Color    string // 6-digit hex RGB, empty = not set
	ColorRaw uint32 // Raw color value (for scheme color detection)
}

// MasterSlide represents a parsed slide master from the PPT file.
type MasterSlide struct {
	Background  SlideBackground
	Shapes      []ShapeFormatting
	ColorScheme []string // 8 RGB hex strings (scheme indices 0-7)
	// DefaultTextStyles holds default text properties per indent level (0-4).
	// Index 0 = level 0 (no indent), index 4 = level 4.
	DefaultTextStyles [5]MasterTextStyle
	// TextTypeStyles holds default text properties per text type (from TextHeaderAtom).
	// Key is the text type (0=title, 1=body, 4=other, etc.), value is styles per indent level.
	TextTypeStyles map[int][5]MasterTextStyle
}

// ResolveSchemeColor resolves a PPT color value that may be a scheme color reference.
// Per [MS-PPT] ColorIndexStruct, the high byte (index field) determines interpretation:
//   0x00-0x07: scheme color index (0=bg, 1=text, 2=shadow, 3=title, 4=fill, 5=accent1, 6=accent2, 7=accent3)
//   0x08: scheme color with index in low byte (legacy format)
//   0xFE: direct sRGB color specified by the RGB bytes
//   0xFF: undefined
// For 0x00 we skip resolution since 0x00000000 is commonly used as "no explicit color".
func ResolveSchemeColor(colorHex string, rawVal uint32, scheme []string) string {
	if len(scheme) < 8 || colorHex == "" {
		return colorHex
	}
	flag := rawVal >> 24 // high byte = index field
	switch {
	case flag >= 0x01 && flag <= 0x07:
		// Direct scheme color index (1=text, 2=shadow, 3=title, 4=fill, 5=accent1, 6=accent2, 7=accent3)
		idx := int(flag)
		if idx < len(scheme) {
			return scheme[idx]
		}
	case flag == 0x08:
		// Legacy scheme color reference with index in low byte
		idx := int(rawVal & 0xFF)
		if idx >= 0 && idx < len(scheme) {
			return scheme[idx]
		}
	case flag == 0xFE:
		// Direct sRGB color OR scheme[0] reference.
		// Per [MS-PPT], 0xFE means "direct sRGB" — the low 3 bytes are the RGB color.
		// Special case: 0xFE000000 (all-zero RGB) is commonly used as a scheme[0]
		// (background/lt1) reference rather than literal black.
		//
		// Non-zero low bytes always represent the direct sRGB color value.
		if rawVal&0x00FFFFFF == 0 {
			// 0xFE000000: scheme[0] reference
			if len(scheme) > 0 {
				return scheme[0]
			}
		}
		// Non-zero low bytes: direct sRGB color
		return colorHex
	}
	return colorHex
}
