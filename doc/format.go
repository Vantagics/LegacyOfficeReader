package doc

import (
	"strings"
)

// cleanRunText strips DOC special characters from text run content.
// Removes: 0x07 (cell mark), 0x08 (drawn object), 0x13/0x14/0x15 (field codes),
// 0x02-0x05 (footnote/annotation refs), 0x0C (page break - detected separately).
// Keeps: 0x01 (image placeholder).
func cleanRunText(raw string) string {
	var result []rune
	depth := 0
	inInstruction := false

	for _, r := range raw {
		switch r {
		case 0x13:
			depth++
			inInstruction = true
		case 0x14:
			inInstruction = false
		case 0x15:
			depth--
			if depth <= 0 {
				depth = 0
				inInstruction = false
			}
		case 0x02, 0x03, 0x04, 0x05:
			continue
		case 0x07, 0x08:
			continue
		case 0x0C:
			continue // removed from display text; HasPageBreak detected separately
		default:
			if depth > 0 && inInstruction {
				continue
			}
			result = append(result, r)
		}
	}
	return string(result)
}

// buildFormattedContent combines all parsed format data into a structured FormattedContent.
func buildFormattedContent(
	rawText string,
	charRuns []charFormatRun,
	paraRuns []paraFormatRun,
	styles []styleDef,
	lists []listDef,
	listOverrides []listOverride,
	sections []sectionBreak,
	shapeMappings []shapeImageMapping,
	picLocToBSE map[int32]int,
) *FormattedContent {
	// Normalize line endings
	rawText = strings.ReplaceAll(rawText, "\r\n", "\r")

	// In DOC format, paragraphs are separated by \r (0x0D) and table cells/rows
	// are separated by \x07 (cell mark). Both act as paragraph terminators.
	// We split on both to properly handle table content.
	// We use a custom split that tracks which separator was used.
	type paraInfo struct {
		text      string
		separator rune // '\r' or '\x07'
	}
	var paraTexts []paraInfo
	runes := []rune(rawText)
	start := 0
	for i, r := range runes {
		if r == '\r' || r == 0x07 {
			paraTexts = append(paraTexts, paraInfo{text: string(runes[start:i]), separator: r})
			start = i + 1
		}
	}
	// Handle remaining text after last separator
	if start < len(runes) {
		paraTexts = append(paraTexts, paraInfo{text: string(runes[start:]), separator: 0})
	}

	// Remove trailing empty paragraph if the text ends with a separator
	if len(paraTexts) > 0 && paraTexts[len(paraTexts)-1].text == "" && paraTexts[len(paraTexts)-1].separator == 0 {
		paraTexts = paraTexts[:len(paraTexts)-1]
	}

	// Build a CP-indexed lookup for shape mappings
	// shapeMappings maps PlcSpaMom anchor CPs to BSE indices
	// We need to find which paragraph each shape belongs to
	// The anchor CP is the paragraph that the shape is anchored to
	// But the visual position is at the 0x08 character, not the anchor
	// So we track 0x08 positions and match them to shapes by order
	shapesByOrder := make([]int, len(shapeMappings)) // BSE indices in order
	for i, sm := range shapeMappings {
		shapesByOrder[i] = sm.bseIndex
	}

	fc := &FormattedContent{
		Paragraphs: make([]Paragraph, 0, len(paraTexts)),
	}

	// Track global character position as we iterate through paragraphs.
	// Each paragraph occupies cpStart to cpStart+len(text), then +1 for the separator.
	cpPos := uint32(0)
	shapeIdx := 0 // index into shapesByOrder for matching 0x08 chars

	for _, pInfo := range paraTexts {
		pText := pInfo.text
		pLen := uint32(len([]rune(pText)))
		cpStart := cpPos
		cpEnd := cpPos + pLen

		p := Paragraph{}

		// If this paragraph was separated by 0x07, it's a table cell/row mark
		if pInfo.separator == 0x07 {
			p.InTable = true
			p.IsTableCellEnd = true
		}

		// Find matching paraFormatRun by character position overlap
		pr := findParaRun(paraRuns, cpStart, cpEnd)

		if pr != nil {
			p.Props = pr.props
			// Inherit paragraph properties from style if not set by direct PAPX
			if int(pr.istd) < len(styles) {
				mergeStyleParaProps(&p.Props, pr.istd, styles)
			}
			// Use inTable from PAPX if set (overrides separator-based detection)
			if pr.inTable {
				p.InTable = true
			}
			// Only mark as row-end if PAPX says so AND the paragraph text is empty.
			if pr.tableRowEnd && pLen == 0 {
				p.TableRowEnd = true
			}
			p.PageBreakBefore = pr.pageBreakBefore
			if pr.tableRowEnd && len(pr.cellWidths) > 0 {
				p.CellWidths = pr.cellWidths
				if pLen == 0 {
					p.TableRowEnd = true
				}
			}

			if pInfo.separator == 0x07 && !p.TableRowEnd {
				if len(pr.cellWidths) > 0 && pLen == 0 {
					p.TableRowEnd = true
					p.InTable = true
					p.CellWidths = pr.cellWidths
				}
			}

			// Heading identification (priority order)
			p.HeadingLevel = identifyHeading(pr, styles)

			// TOC detection: STI 19-27 are TOC 1-9 styles
			if int(pr.istd) < len(styles) {
				sti := styles[pr.istd].sti
				if sti >= 19 && sti <= 27 {
					p.IsTOC = true
					p.TOCLevel = uint8(sti - 18)
				}
			}

			// List marking
			if pr.ilfo > 0 {
				p.IsListItem = true
				p.ListLevel = pr.ilvl
				p.ListIlfo = pr.ilfo
				p.ListType, p.ListNfc, p.ListLvlText = resolveListInfo(pr.ilfo, lists, listOverrides)
			}
		} else {
			// No PAPX run found for this paragraph.
			// Apply Normal style (istd=0) properties as fallback.
			if len(styles) > 0 {
				mergeStyleParaProps(&p.Props, 0, styles)
			}
		}

		// Page break detection: check if paragraph text contains 0x0C
		if strings.ContainsRune(pText, 0x0C) {
			p.HasPageBreak = true
		}

		// Drawn object detection: check if any 0x08 characters in this paragraph
		// correspond to shapes. Match 0x08 chars to PlcSpaMom entries by order.
		// Shapes with bseIndex=-1 are non-image shapes (text boxes, etc.) and
		// are skipped in the drawn images list but still consume their 0x08 slot.
		for _, r := range []rune(pText) {
			if r == 0x08 && shapeIdx < len(shapesByOrder) {
				if shapesByOrder[shapeIdx] >= 0 {
					p.DrawnImages = append(p.DrawnImages, shapesByOrder[shapeIdx])
				}
				shapeIdx++
			}
		}

		// Section break detection: check if this paragraph's end CP matches a section break
		for _, sb := range sections {
			// Section break CP is the position of the section mark character
			// which is at the end of the paragraph (cpEnd position)
			if sb.cpEnd >= cpStart && sb.cpEnd <= cpEnd+1 {
				p.IsSectionBreak = true
				// Map bkc values: 0=continuous, 1=new column, 2=new page, 3=even page, 4=odd page
				switch sb.bkc {
				case 0:
					p.SectionType = 0 // continuous
				case 2:
					p.SectionType = 1 // new page
				case 3:
					p.SectionType = 2 // even page
				case 4:
					p.SectionType = 3 // odd page
				default:
					p.SectionType = 1 // default to new page
				}
				break
			}
		}

		// Build TextRuns by splitting paragraph text according to charFormatRun ranges
		p.Runs = buildTextRuns(pText, cpStart, charRuns, pr, styles, picLocToBSE)

		fc.Paragraphs = append(fc.Paragraphs, p)

		// Advance past this paragraph text + 1 for the separator
		cpPos = cpEnd + 1
	}

	// Post-processing: fix table structure.
	// Some row-end paragraphs may not have inTable=true from PAPX.
	// If a 0x07-terminated paragraph (InTable from separator) is between
	// InTable paragraphs, ensure it's marked as InTable.
	// Also detect row-end paragraphs: in a sequence of InTable paragraphs,
	// find the ones that have cellWidths (sprmTDefTable) and mark them as row-end.
	// For rows without explicit cellWidths, use the column count from the nearest
	// row that has cellWidths.
	fixTableStructure(fc)

	return fc
}

// fixTableStructure post-processes paragraphs to fix table row detection.
// In DOC format, table rows are sequences of 0x07-terminated paragraphs.
// Row-end paragraphs should have sprmPFTtp=1 and sprmTDefTable, but some DOC files
// don't set these consistently (e.g., the header row may lack sprmPFTtp).
// This function uses explicit row-end markers to determine column count,
// then applies that to find missing row-end markers.
func fixTableStructure(fc *FormattedContent) {
	if len(fc.Paragraphs) == 0 {
		return
	}

	// Step 1: Propagate InTable to gaps between InTable paragraphs
	for i := 1; i < len(fc.Paragraphs)-1; i++ {
		if !fc.Paragraphs[i].InTable && fc.Paragraphs[i-1].InTable && fc.Paragraphs[i+1].InTable {
			fc.Paragraphs[i].InTable = true
		}
	}

	// Step 2: Determine column count from explicit row-end markers.
	// Find two consecutive row-end markers and count cells between them.
	// Or find a row-end with cellWidths and use that count.
	colCount := 0

	// Method A: Count cells between two consecutive explicit row-ends
	prevRowEnd := -1
	for i := range fc.Paragraphs {
		if !fc.Paragraphs[i].InTable {
			prevRowEnd = -1
			continue
		}
		if fc.Paragraphs[i].TableRowEnd {
			if prevRowEnd >= 0 {
				count := 0
				for j := prevRowEnd + 1; j < i; j++ {
					if fc.Paragraphs[j].InTable && !fc.Paragraphs[j].TableRowEnd {
						count++
					}
				}
				if count > 0 {
					colCount = count
					break
				}
			}
			prevRowEnd = i
		}
	}

	// Method B: Fallback - count non-empty cells before first empty cell
	if colCount == 0 {
		for i := range fc.Paragraphs {
			if !fc.Paragraphs[i].InTable {
				continue
			}
			nonEmptyCount := 0
			for j := i; j < len(fc.Paragraphs) && fc.Paragraphs[j].InTable; j++ {
				hasText := false
				for _, r := range fc.Paragraphs[j].Runs {
					if r.Text != "" {
						hasText = true
						break
					}
				}
				if hasText {
					nonEmptyCount++
				} else {
					colCount = nonEmptyCount
					break
				}
			}
			break
		}
	}

	if colCount == 0 {
		return
	}

	// Step 3: Walk through table paragraphs and mark missing row-end markers.
	// Only mark empty paragraphs as row-ends (real row-end markers are always empty).
	cellCount := 0
	for i := range fc.Paragraphs {
		if !fc.Paragraphs[i].InTable {
			cellCount = 0
			continue
		}
		if fc.Paragraphs[i].TableRowEnd {
			cellCount = 0
			continue
		}
		cellCount++
		if cellCount == colCount {
			// The next InTable paragraph should be the row-end mark.
			// Only mark it if it has no text content (real row-end markers are empty).
			if i+1 < len(fc.Paragraphs) && fc.Paragraphs[i+1].InTable && !fc.Paragraphs[i+1].TableRowEnd {
				hasText := false
				for _, r := range fc.Paragraphs[i+1].Runs {
					if r.Text != "" {
						hasText = true
						break
					}
				}
				if !hasText {
					fc.Paragraphs[i+1].TableRowEnd = true
					// Copy cellWidths from the nearest explicit row-end if available
					for j := i + 2; j < len(fc.Paragraphs); j++ {
						if fc.Paragraphs[j].TableRowEnd && len(fc.Paragraphs[j].CellWidths) > 0 {
							fc.Paragraphs[i+1].CellWidths = fc.Paragraphs[j].CellWidths
							break
						}
					}
				}
			}
			cellCount = 0
		}
	}
}

// findParaRun finds the paraFormatRun that overlaps with the given character range.
// Uses binary search since paraRuns are sorted by cpStart.
func findParaRun(paraRuns []paraFormatRun, cpStart, cpEnd uint32) *paraFormatRun {
	// For zero-length paragraphs (cpStart == cpEnd), find the run that contains cpStart
	if cpEnd <= cpStart {
		cpEnd = cpStart + 1
	}
	lo, hi := 0, len(paraRuns)
	for lo < hi {
		mid := lo + (hi-lo)/2
		if paraRuns[mid].cpEnd <= cpStart {
			lo = mid + 1
		} else {
			hi = mid
		}
	}
	if lo < len(paraRuns) {
		pr := &paraRuns[lo]
		if pr.cpStart < cpEnd {
			return pr
		}
	}
	return nil
}

// identifyHeading determines the heading level for a paragraph.
// Priority: outLvl (0-8) > sti (built-in style index 1-9) > style name matching.
func identifyHeading(pr *paraFormatRun, styles []styleDef) uint8 {
	// 1. If outLvl is 0-8 (not 9 = body text), use outLvl + 1
	if pr.outLvl <= 8 {
		return pr.outLvl + 1
	}

	// 2. Check the style's built-in style index (sti)
	// sti 1-9 are the built-in heading styles (Heading 1 through Heading 9)
	if int(pr.istd) < len(styles) {
		sti := styles[pr.istd].sti
		if sti >= 1 && sti <= 9 {
			return uint8(sti)
		}
	}

	// 3. Check style name for "heading N" or Chinese "标题 N" (case-insensitive)
	if int(pr.istd) < len(styles) {
		name := strings.ToLower(styles[pr.istd].name)
		// English heading names
		if strings.HasPrefix(name, "heading ") {
			rest := name[len("heading "):]
			if len(rest) == 1 && rest[0] >= '1' && rest[0] <= '9' {
				return rest[0] - '0'
			}
		}
		// Chinese heading names: "标题 N" (U+6807 U+9898)
		chineseHeading := "\u6807\u9898"
		if strings.HasPrefix(name, chineseHeading) {
			rest := strings.TrimPrefix(name, chineseHeading)
			rest = strings.TrimSpace(rest)
			if len(rest) == 1 && rest[0] >= '1' && rest[0] <= '9' {
				return rest[0] - '0'
			}
		}
	}

	return 0
}

// resolveListType determines the list type (0=unordered, 1=ordered) for a list item.
func resolveListType(ilfo uint16, lists []listDef, listOverrides []listOverride) uint8 {
	lt, _, _ := resolveListInfo(ilfo, lists, listOverrides)
	return lt
}

// resolveListInfo returns the list type (0=unordered, 1=ordered), the nfc
// (number format code), and the lvlText template for the given ilfo.
func resolveListInfo(ilfo uint16, lists []listDef, listOverrides []listOverride) (uint8, uint8, string) {
	if ilfo == 0 {
		return 0, 23, ""
	}

	idx := int(ilfo) - 1
	if idx < 0 || idx >= len(listOverrides) {
		return 0, 23, "" // default unordered/bullet
	}

	targetID := listOverrides[idx].listID
	for _, ld := range lists {
		if ld.listID == targetID {
			if ld.ordered {
				return 1, ld.nfc, ld.lvlText
			}
			return 0, ld.nfc, ld.lvlText
		}
	}

	return 0, 23, "" // default unordered if list def not found
}

// buildTextRuns splits paragraph text into TextRuns based on charFormatRun ranges.
// It applies style inheritance for properties not explicitly set in the CHPX.
// Field codes (0x13/0x14/0x15) are handled at the paragraph level since they
// can span multiple runs.
// picLocToBSE maps sprmCPicLocation offsets to BSE image indices (0-based).
func buildTextRuns(pText string, cpStart uint32, charRuns []charFormatRun, pr *paraFormatRun, styles []styleDef, picLocToBSE map[int32]int) []TextRun {
	// Resolve default character formatting from the paragraph's style
	var defaultProps CharacterFormatting
	if pr != nil && int(pr.istd) < len(styles) {
		defaultProps = resolveStyleCharProps(pr.istd, styles)
	}

	if len(charRuns) == 0 || len(pText) == 0 {
		cleaned := cleanRunText(pText)
		if cleaned == "" && pText == "" {
			return []TextRun{{Text: "", Props: defaultProps, ImageRef: -1}}
		}
		return []TextRun{{Text: cleaned, Props: defaultProps, ImageRef: -1}}
	}

	pRunes := []rune(pText)
	pLen := uint32(len(pRunes))
	cpEnd := cpStart + pLen

	// First pass: build a per-rune field depth map so we can strip field codes
	// that span across multiple runs. Field codes: 0x13=begin, 0x14=separator, 0x15=end.
	// Text between 0x13 and 0x14 is field instruction (hidden).
	// Text between 0x14 and 0x15 is field result (visible).
	fieldDepth := make([]int, pLen)
	inInstruction := make([]bool, pLen)
	depth := 0
	isInstr := false
	for i, r := range pRunes {
		switch r {
		case 0x13:
			depth++
			isInstr = true
		case 0x14:
			isInstr = false
		case 0x15:
			depth--
			if depth <= 0 {
				depth = 0
				isInstr = false
			}
		}
		fieldDepth[i] = depth
		inInstruction[i] = isInstr
	}

	// Binary search to find the first charRun that could overlap this paragraph
	startIdx := 0
	lo, hi := 0, len(charRuns)
	for lo < hi {
		mid := lo + (hi-lo)/2
		if charRuns[mid].cpEnd <= cpStart {
			lo = mid + 1
		} else {
			hi = mid
		}
	}
	startIdx = lo

	// Collect raw runs (before field stripping)
	type rawRun struct {
		localStart uint32
		localEnd   uint32
		props      CharacterFormatting
	}
	var rawRuns []rawRun
	covered := uint32(0)

	for i := startIdx; i < len(charRuns); i++ {
		cr := charRuns[i]
		if cr.cpStart >= cpEnd {
			break
		}
		overlapStart := cr.cpStart
		if overlapStart < cpStart {
			overlapStart = cpStart
		}
		overlapEnd := cr.cpEnd
		if overlapEnd > cpEnd {
			overlapEnd = cpEnd
		}
		localStart := overlapStart - cpStart
		localEnd := overlapEnd - cpStart

		// Skip charRuns that overlap already-covered text (prevents duplicate text)
		if localStart < covered {
			localStart = covered
		}
		if localStart >= localEnd {
			continue
		}

		if localStart > covered {
			rawRuns = append(rawRuns, rawRun{covered, localStart, defaultProps})
		}
		rawRuns = append(rawRuns, rawRun{localStart, localEnd, mergeCharProps(defaultProps, cr.props)})
		covered = localEnd
	}
	if covered < pLen {
		rawRuns = append(rawRuns, rawRun{covered, pLen, defaultProps})
	}

	// Second pass: for each raw run, emit only visible characters
	var runs []TextRun
	for _, rr := range rawRuns {
		var visible []rune
		imageRef := -1 // BSE index for inline image, -1 if not an image
		for i := rr.localStart; i < rr.localEnd; i++ {
			r := pRunes[i]
			// Skip field delimiter characters themselves
			if r == 0x13 || r == 0x14 || r == 0x15 {
				continue
			}
			// Skip field instruction text
			if fieldDepth[i] > 0 && inInstruction[i] {
				continue
			}
			// Skip other special characters
			if r == 0x02 || r == 0x03 || r == 0x04 || r == 0x05 || r == 0x07 || r == 0x08 {
				continue
			}
			// Keep 0x0C for page break detection but strip from display
			if r == 0x0C {
				continue
			}
			// For 0x01 (inline image), resolve BSE index from PicLocation
			if r == 0x01 && rr.props.HasPicLocation && picLocToBSE != nil {
				if bseIdx, ok := picLocToBSE[rr.props.PicLocation]; ok {
					imageRef = bseIdx
				}
			}
			visible = append(visible, r)
		}
		text := string(visible)
		if text != "" {
			runs = append(runs, TextRun{Text: text, Props: rr.props, ImageRef: imageRef})
		}
	}

	if len(runs) == 0 {
		runs = []TextRun{{Text: "", Props: defaultProps, ImageRef: -1}}
	}

	return runs
}

// resolveStyleCharProps resolves character formatting from a style, following
// the style inheritance chain (istdBase).
func resolveStyleCharProps(istd uint16, styles []styleDef) CharacterFormatting {
	var result CharacterFormatting
	visited := make(map[uint16]bool)

	for {
		if int(istd) >= len(styles) || visited[istd] {
			break
		}
		visited[istd] = true
		s := styles[istd]

		if s.charProps != nil {
			// Apply properties from this style (base styles are applied first)
			if result.FontName == "" && s.charProps.FontName != "" {
				result.FontName = s.charProps.FontName
			}
			if result.FontSize == 0 && s.charProps.FontSize != 0 {
				result.FontSize = s.charProps.FontSize
			}
			if !result.Bold && s.charProps.Bold {
				result.Bold = true
			}
			if !result.Italic && s.charProps.Italic {
				result.Italic = true
			}
			if result.Underline == 0 && s.charProps.Underline != 0 {
				result.Underline = s.charProps.Underline
			}
			if result.Color == "" && s.charProps.Color != "" {
				result.Color = s.charProps.Color
			}
		}

		// Follow inheritance chain
		if s.istdBase == 0xFFF || s.istdBase == istd {
			break
		}
		istd = s.istdBase
	}

	// Apply document defaults if style chain didn't provide them
	// Default font size in Word is 10pt = 20 half-points
	if result.FontSize == 0 {
		result.FontSize = 21 // 10.5pt (common Chinese doc default) = 21 half-points
	}

	return result
}

// mergeCharProps merges explicit CHPX properties with style defaults.
// Explicit properties take precedence over defaults.
func mergeCharProps(defaults, explicit CharacterFormatting) CharacterFormatting {
	result := explicit
	if result.FontName == "" {
		result.FontName = defaults.FontName
	}
	if result.FontSize == 0 {
		result.FontSize = defaults.FontSize
	}
	if result.Color == "" {
		result.Color = defaults.Color
	}
	// Bold, Italic, Underline: explicit values always win (they're set explicitly in CHPX)
	// But if the CHPX didn't set them (they're false/0), inherit from style
	// Note: we can't distinguish "explicitly set to false" from "not set" with bool,
	// so we only inherit true values from defaults
	if !result.Bold && defaults.Bold {
		result.Bold = true
	}
	if !result.Italic && defaults.Italic {
		result.Italic = true
	}
	if result.Underline == 0 && defaults.Underline != 0 {
		result.Underline = defaults.Underline
	}
	return result
}

// resolveStyleParaProps resolves paragraph formatting from a style, following
// the style inheritance chain (istdBase).
func resolveStyleParaProps(istd uint16, styles []styleDef) ParagraphFormatting {
	var result ParagraphFormatting
	visited := make(map[uint16]bool)

	for {
		if int(istd) >= len(styles) || visited[istd] {
			break
		}
		visited[istd] = true
		s := styles[istd]

		if s.paraProps != nil {
			pp := s.paraProps
			if !result.AlignmentSet && pp.AlignmentSet {
				result.Alignment = pp.Alignment
				result.AlignmentSet = true
			}
			if result.IndentLeft == 0 && pp.IndentLeft != 0 {
				result.IndentLeft = pp.IndentLeft
			}
			if result.IndentRight == 0 && pp.IndentRight != 0 {
				result.IndentRight = pp.IndentRight
			}
			if result.IndentFirst == 0 && pp.IndentFirst != 0 {
				result.IndentFirst = pp.IndentFirst
			}
			if result.SpaceBefore == 0 && pp.SpaceBefore != 0 {
				result.SpaceBefore = pp.SpaceBefore
			}
			if result.SpaceAfter == 0 && pp.SpaceAfter != 0 {
				result.SpaceAfter = pp.SpaceAfter
			}
			if result.LineSpacing == 0 && pp.LineSpacing != 0 {
				result.LineSpacing = pp.LineSpacing
				result.LineRule = pp.LineRule
			}
		}

		if s.istdBase == 0xFFF || s.istdBase == istd {
			break
		}
		istd = s.istdBase
	}

	return result
}

// mergeStyleParaProps merges style paragraph properties into direct paragraph
// properties. Direct PAPX values take precedence; style values fill in gaps.
func mergeStyleParaProps(direct *ParagraphFormatting, istd uint16, styles []styleDef) {
	stylePP := resolveStyleParaProps(istd, styles)
	if !direct.AlignmentSet && stylePP.AlignmentSet {
		direct.Alignment = stylePP.Alignment
		direct.AlignmentSet = true
	}
	if direct.IndentLeft == 0 && stylePP.IndentLeft != 0 {
		direct.IndentLeft = stylePP.IndentLeft
	}
	if direct.IndentRight == 0 && stylePP.IndentRight != 0 {
		direct.IndentRight = stylePP.IndentRight
	}
	if direct.IndentFirst == 0 && stylePP.IndentFirst != 0 {
		direct.IndentFirst = stylePP.IndentFirst
	}
	if direct.SpaceBefore == 0 && stylePP.SpaceBefore != 0 {
		direct.SpaceBefore = stylePP.SpaceBefore
	}
	if direct.SpaceAfter == 0 && stylePP.SpaceAfter != 0 {
		direct.SpaceAfter = stylePP.SpaceAfter
	}
	if direct.LineSpacing == 0 && stylePP.LineSpacing != 0 {
		direct.LineSpacing = stylePP.LineSpacing
		direct.LineRule = stylePP.LineRule
	}
}
