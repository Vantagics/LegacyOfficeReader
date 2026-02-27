package doc

import (
	"fmt"

	"github.com/shakinm/xlsReader/common"
)

// Document represents a parsed DOC file.
type Document struct {
	text             string
	images           []common.Image
	formattedContent *FormattedContent // can be nil if format parsing failed
	lid              uint16            // language ID from FIB
	codepage         uint16            // resolved codepage
	fonts            []string          // font names from SttbfFfn
	styles           []styleDef        // style definitions from STSH
	charRuns         []charFormatRun   // for debugging
	paraRuns         []paraFormatRun   // for debugging
	fcPlcfSed        uint32            // for debugging
	lcbPlcfSed       uint32            // for debugging
	picLocToBSE      map[int32]int     // PicLocation offset -> BSE index (0-based)
}

// GetText returns the full plain text content of the document.
func (d *Document) GetText() string {
	return d.text
}

// GetImages returns all embedded images extracted from the document.
// It always returns a non-nil slice; if there are no images, it returns an empty slice.
func (d *Document) GetImages() []common.Image {
	if d.images != nil {
		return d.images
	}
	return []common.Image{}
}

// GetFormattedContent returns the structured formatted document content.
// Returns nil if format parsing was not available or failed.
func (d *Document) GetFormattedContent() *FormattedContent {
	return d.formattedContent
}

// GetLid returns the language ID from the FIB.
func (d *Document) GetLid() uint16 {
	return d.lid
}

// GetCodepage returns the resolved codepage.
func (d *Document) GetCodepage() uint16 {
	return d.codepage
}

// GetFonts returns the font name table.
func (d *Document) GetFonts() []string {
	return d.fonts
}

// GetStyles returns the style names.
func (d *Document) GetStyles() []string {
	names := make([]string, len(d.styles))
	for i, s := range d.styles {
		names[i] = s.name
	}
	return names
}

// GetStyleSTIs returns the built-in style indices.
func (d *Document) GetStyleSTIs() []uint16 {
	stis := make([]uint16, len(d.styles))
	for i, s := range d.styles {
		stis[i] = s.sti
	}
	return stis
}

// DebugRanges prints debug info about char/para run ranges.
func (d *Document) DebugRanges() {
	fmt.Printf("CharRuns: %d, ParaRuns: %d\n", len(d.charRuns), len(d.paraRuns))
	// Show para runs with non-zero istd or heading-related properties
	for i := 0; i < len(d.paraRuns); i++ {
		pr := d.paraRuns[i]
		if pr.istd != 0 || pr.outLvl != 9 || pr.inTable || pr.ilfo != 0 {
			fmt.Printf("  ParaRun %d: cp[%d-%d] istd=%d outLvl=%d inTable=%v ilfo=%d align=%d\n",
				i, pr.cpStart, pr.cpEnd, pr.istd, pr.outLvl, pr.inTable, pr.ilfo, pr.props.Alignment)
		}
	}
	// Show char runs with non-zero size
	for i := 0; i < len(d.charRuns); i++ {
		cr := d.charRuns[i]
		if cr.props.FontSize != 0 || cr.props.Bold {
			fmt.Printf("  CharRun %d: cp[%d-%d] font=%q size=%d bold=%v color=%q\n",
				i, cr.cpStart, cr.cpEnd, cr.props.FontName, cr.props.FontSize, cr.props.Bold, cr.props.Color)
		}
	}
}

// DebugSections prints section break info for debugging.
func (d *Document) DebugSections() {
	fmt.Printf("PlcfSed: fc=%d, lcb=%d\n", d.fcPlcfSed, d.lcbPlcfSed)
}

// DebugStyleProps prints paragraph and character properties for each style.
func (d *Document) DebugStyleProps() {
	for i, s := range d.styles {
		if s.name == "" && i >= 10 {
			continue
		}
		fmt.Printf("Style[%d] %q (sti=%d type=%d base=%d):\n", i, s.name, s.sti, s.styleType, s.istdBase)
		if s.paraProps != nil {
			pp := s.paraProps
			fmt.Printf("  paraProps: align=%d alignSet=%v indent=%d/%d/%d spacing=%d/%d line=%d/%d\n",
				pp.Alignment, pp.AlignmentSet, pp.IndentLeft, pp.IndentRight, pp.IndentFirst,
				pp.SpaceBefore, pp.SpaceAfter, pp.LineSpacing, pp.LineRule)
		} else {
			fmt.Printf("  paraProps: nil\n")
		}
		if s.charProps != nil {
			cp := s.charProps
			fmt.Printf("  charProps: font=%q size=%d bold=%v italic=%v color=%q\n",
				cp.FontName, cp.FontSize, cp.Bold, cp.Italic, cp.Color)
		} else {
			fmt.Printf("  charProps: nil\n")
		}
	}
}

// DebugParaRunDetails prints detailed PAPX run info for debugging alignment.
func (d *Document) DebugParaRunDetails() {
	fmt.Printf("Total paraRuns: %d\n", len(d.paraRuns))
	for i, pr := range d.paraRuns {
		fmt.Printf("ParaRun[%d]: cp[%d-%d] istd=%d align=%d alignSet=%v inTable=%v\n",
			i, pr.cpStart, pr.cpEnd, pr.istd, pr.props.Alignment, pr.props.AlignmentSet, pr.inTable)
	}
}

// DebugPieces prints piece table info for debugging.
func (d *Document) DebugPieces() {
	// Not directly accessible from Document - need to expose from doc.go
	fmt.Println("(Piece table debug not available from Document)")
}
