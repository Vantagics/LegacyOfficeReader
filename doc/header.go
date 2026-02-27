package doc

import (
	"encoding/binary"
	"strings"
)

// extractHeaderFooter extracts header and footer text from the document.
// In DOC format, header/footer text follows the main document text in the
// character stream. The Plcfhdd structure in the Table stream specifies
// the CP ranges for each header/footer story.
//
// The header document contains stories in this order (per section):
//   [0] even page header
//   [1] odd page header
//   [2] even page footer
//   [3] odd page footer
//   [4] first page header
//   [5] first page footer
//
// Returns separate header and footer text slices.
func extractHeaderFooter(fullText string, ccpText, ccpFtn, ccpHdd uint32, tableData []byte, fcPlcfHdd, lcbPlcfHdd uint32, hdrShapeMappings []shapeImageMapping) ([]string, []string, []string, []string, [][]int, [][]int, []HeaderFooterEntry, []HeaderFooterEntry) {
	if ccpHdd == 0 || lcbPlcfHdd == 0 {
		return nil, nil, nil, nil, nil, nil, nil, nil
	}

	runes := []rune(fullText)
	totalRunes := uint32(len(runes))

	// Header/footer text starts after main document text and footnote text.
	// Per [MS-DOC], the character stream order is: main text, footnote text, header text.
	// The Plcfhdd CPs are relative to the start of the header document.
	hddStart := ccpText + ccpFtn

	// Parse Plcfhdd: array of CPs (uint32 each), no data elements
	if uint64(fcPlcfHdd)+uint64(lcbPlcfHdd) > uint64(len(tableData)) {
		return nil, nil, nil, nil, nil, nil, nil, nil
	}
	plcData := tableData[fcPlcfHdd : fcPlcfHdd+lcbPlcfHdd]

	// Number of CPs = lcbPlcfHdd / 4
	nCPs := lcbPlcfHdd / 4
	if nCPs < 2 {
		return nil, nil, nil, nil, nil, nil, nil, nil
	}

	cps := make([]uint32, nCPs)
	for i := uint32(0); i < nCPs; i++ {
		if i*4+4 > uint32(len(plcData)) {
			break
		}
		cps[i] = binary.LittleEndian.Uint32(plcData[i*4:])
	}

	// The header area ends at hddStart + ccpHdd. CPs beyond that belong to
	// other sub-documents (text boxes, etc.) and must not be read.
	hddEnd := hddStart + ccpHdd

	// Build a mapping from header-relative CP to shape BSE indices.
	// hdrShapeMappings contains shapes from PlcSpaHdr with CPs relative to the header subdocument.
	hdrShapesByCP := make(map[uint32][]int) // CP -> list of BSE indices
	for _, sm := range hdrShapeMappings {
		if sm.bseIndex >= 0 {
			hdrShapesByCP[sm.cp] = append(hdrShapesByCP[sm.cp], sm.bseIndex)
		}
	}

	// Extract text for each story
	var headers []string
	var footers []string
	var headersRaw []string
	var footersRaw []string
	var headerImages [][]int
	var footerImages [][]int
	var headerEntries []HeaderFooterEntry
	var footerEntries []HeaderFooterEntry

	// Stories come in groups of 6 per section
	numStories := int(nCPs) - 1
	for i := 0; i+1 < int(nCPs) && i < numStories; i++ {
		cpStart := hddStart + cps[i]
		cpEnd := hddStart + cps[i+1]

		// Clamp to header area boundary
		if cpEnd > hddEnd {
			cpEnd = hddEnd
		}

		if cpStart >= totalRunes || cpEnd > totalRunes || cpStart >= cpEnd {
			continue
		}

		storyText := string(runes[cpStart:cpEnd])

		// Collect drawn image BSE indices from 0x08 characters in this story.
		// Match 0x08 chars to PlcSpaHdr entries by their CP position.
		var storyImages []int
		for j, r := range []rune(storyText) {
			if r == 0x08 {
				// The CP of this 0x08 char relative to header subdocument start
				relCP := cps[i] + uint32(j)
				if bseList, ok := hdrShapesByCP[relCP]; ok {
					storyImages = append(storyImages, bseList...)
				}
			}
		}

		// Keep raw text (with field codes) before cleaning
		rawText := strings.TrimRight(storyText, "\r\n")
		// Strip only non-field special chars from raw text, keep 0x13/0x14/0x15
		rawText = stripNonFieldSpecialChars(rawText)

		// Clean up: remove trailing \r and special chars
		storyText = strings.TrimRight(storyText, "\r\n")
		storyText = cleanHeaderText(storyText)

		// Skip stories that have neither text nor images
		if storyText == "" && len(storyImages) == 0 {
			continue
		}

		// Determine story type within the 6-story group
		storyIdx := i % 6
		switch storyIdx {
		case 1: // odd page header (primary - "default" in DOCX)
			headers = appendIfNew(headers, storyText)
			headersRaw = append(headersRaw, rawText)
			headerImages = append(headerImages, storyImages)
			headerEntries = append(headerEntries, HeaderFooterEntry{Type: "default", Text: storyText, RawText: rawText, Images: storyImages})
		case 0: // even page header
			headers = appendIfNew(headers, storyText)
			headersRaw = append(headersRaw, rawText)
			headerImages = append(headerImages, storyImages)
			headerEntries = append(headerEntries, HeaderFooterEntry{Type: "even", Text: storyText, RawText: rawText, Images: storyImages})
		case 4: // first page header
			headers = appendIfNew(headers, storyText)
			headersRaw = append(headersRaw, rawText)
			headerImages = append(headerImages, storyImages)
			headerEntries = append(headerEntries, HeaderFooterEntry{Type: "first", Text: storyText, RawText: rawText, Images: storyImages})
		case 3: // odd page footer (primary - "default" in DOCX)
			footers = appendIfNew(footers, storyText)
			footersRaw = append(footersRaw, rawText)
			footerImages = append(footerImages, storyImages)
			footerEntries = append(footerEntries, HeaderFooterEntry{Type: "default", Text: storyText, RawText: rawText, Images: storyImages})
		case 2: // even page footer
			footers = appendIfNew(footers, storyText)
			footersRaw = append(footersRaw, rawText)
			footerImages = append(footerImages, storyImages)
			footerEntries = append(footerEntries, HeaderFooterEntry{Type: "even", Text: storyText, RawText: rawText, Images: storyImages})
		case 5: // first page footer
			footers = appendIfNew(footers, storyText)
			footersRaw = append(footersRaw, rawText)
			footerImages = append(footerImages, storyImages)
			footerEntries = append(footerEntries, HeaderFooterEntry{Type: "first", Text: storyText, RawText: rawText, Images: storyImages})
		}
	}

	return headers, footers, headersRaw, footersRaw, headerImages, footerImages, headerEntries, footerEntries
}

// cleanHeaderText strips DOC field codes and special characters from header/footer text.
// Field codes use 0x13 (begin), 0x14 (separator), 0x15 (end). Text between
// 0x13 and 0x14 is the field instruction and is removed. Text between 0x14
// and 0x15 is the field result (visible text) and is kept.
func cleanHeaderText(s string) string {
	var result []rune
	depth := 0
	inInstruction := false

	for _, r := range s {
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
		case 0x01, 0x02, 0x03, 0x04, 0x05, 0x07, 0x08:
			continue
		case 0x0C:
			continue
		default:
			if depth > 0 && inInstruction {
				continue // skip field instruction text
			}
			result = append(result, r)
		}
	}
	text := strings.TrimSpace(string(result))
	// Filter out strings that are only whitespace or control chars
	hasContent := false
	for _, r := range text {
		if r > 0x20 {
			hasContent = true
			break
		}
	}
	if !hasContent {
		return ""
	}
	return text
}

// stripNonFieldSpecialChars removes special characters except field codes (0x13/0x14/0x15)
// and tab (0x09) from header/footer text.
func stripNonFieldSpecialChars(s string) string {
	var result []rune
	for _, r := range s {
		switch r {
		case 0x01, 0x02, 0x03, 0x04, 0x05, 0x07, 0x08, 0x0C:
			continue
		default:
			result = append(result, r)
		}
	}
	return strings.TrimSpace(string(result))
}

// appendIfNew appends text to the slice only if it's not already present.
func appendIfNew(slice []string, text string) []string {
	for _, s := range slice {
		if s == text {
			return slice
		}
	}
	return append(slice, text)
}
