package doc

// CharacterFormatting holds character-level formatting properties.
type CharacterFormatting struct {
	FontName       string
	FontSize       uint16 // half-points
	Bold           bool
	Italic         bool
	Underline      uint8  // underline type
	Color          string // 6-digit hex RGB, empty means no color
	IstdChar       uint16 // character style index
	PicLocation    int32  // sprmCPicLocation: offset into Data stream for inline images (-1 = not set)
	HasPicLocation bool   // true if sprmCPicLocation was explicitly set
}

// ParagraphFormatting holds paragraph-level formatting properties.
type ParagraphFormatting struct {
	Alignment    uint8  // 0=left, 1=center, 2=right, 3=both
	AlignmentSet bool   // true if Alignment was explicitly set by PAPX (distinguishes "left" from "not set")
	IndentLeft   int32  // twip
	IndentRight  int32  // twip
	IndentFirst  int32  // twip
	SpaceBefore  uint16 // twip
	SpaceAfter   uint16 // twip
	LineSpacing  int32  // line spacing value
	LineRule     uint8  // line spacing rule
}

// TextRun represents a run of text with uniform character formatting.
type TextRun struct {
	Text     string
	Props    CharacterFormatting
	ImageRef int // BSE image index (0-based) for inline image, -1 if not an image
}

// Paragraph represents a single paragraph with its formatting and content.
type Paragraph struct {
	Props           ParagraphFormatting
	Runs            []TextRun
	HeadingLevel    uint8 // 0=not a heading, 1-9=heading level
	IsListItem      bool
	ListType        uint8  // 0=unordered, 1=ordered
	ListLevel       uint8  // 0-8
	ListIlfo        uint16 // list override index (ilfo) from PAPX, identifies the list instance
	ListNfc         uint8  // number format code (0=decimal, 1=upperRoman, 2=lowerRoman, 3=upperLetter, 4=lowerLetter, 23=bullet)
	ListLvlText     string // level text template from DOC (e.g. "(%1)" or "%1.")
	InTable         bool
	TableRowEnd     bool
	CellWidths      []int32 // table cell widths in twips (only on row-end paragraphs)
	PageBreakBefore bool
	HasPageBreak    bool  // text contains 0x0C
	IsSectionBreak  bool
	SectionType     uint8 // 0=continuous, 1=new page, 2=even page, 3=odd page
	IsTOC           bool  // true if this paragraph is a TOC entry
	TOCLevel        uint8 // TOC level (1-9)
	DrawnImages     []int    // BSE image indices (0-based) for drawn objects in this paragraph
	IsTableCellEnd  bool     // true if this paragraph ends a table cell (0x07 separator)
	TextBoxText     string   // text from a text box shape anchored to this paragraph
}

// HeaderFooterEntry holds a single header or footer story with its type and content.
type HeaderFooterEntry struct {
	Type     string // "default", "even", "first"
	Text     string // cleaned visible text
	RawText  string // raw text with field codes
	Images   []int  // BSE image indices for drawn objects
}

// FormattedContent holds the structured formatted document content.
type FormattedContent struct {
	Paragraphs    []Paragraph
	Headers       []string // header text for each section (cleaned, visible text only) - DEPRECATED
	Footers       []string // footer text for each section (cleaned, visible text only) - DEPRECATED
	HeadersRaw    []string // raw header text with field codes (0x13/0x14/0x15) - DEPRECATED
	FootersRaw    []string // raw footer text with field codes - DEPRECATED
	HeaderImages  [][]int  // BSE image indices for each header (drawn objects) - DEPRECATED
	FooterImages  [][]int  // BSE image indices for each footer (drawn objects) - DEPRECATED
	HeaderEntries []HeaderFooterEntry // structured header entries
	FooterEntries []HeaderFooterEntry // structured footer entries
}
