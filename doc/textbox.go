package doc

import (
	"encoding/binary"
	"strings"
)

// textBoxMapping maps a text box shape's anchor CP to its text content.
type textBoxMapping struct {
	anchorCP uint32 // CP in main document where the text box shape is anchored
	spid     uint32 // shape ID
	text     string // text box content
}

// extractTextBoxMappings extracts text box content and maps it to anchor CPs.
// It parses PlcftxbxTxt to find text box text ranges, and the DggInfo to find
// which shapes have text boxes (lTxid property).
func extractTextBoxMappings(fullText string, tableData []byte, f *fib) []textBoxMapping {
	if f.lcbPlcftxbxTxt == 0 {
		return nil
	}

	fullRunes := []rune(fullText)

	// Text box text starts after main + footnote + header + annotation + endnote
	// Read ccpAtn and ccpEdn from FIB (they're in FibRgLw97)
	// For now, use the fields we have
	txbxStart := f.ccpText + f.ccpFtn + f.ccpHdd
	// ccpAtn and ccpEdn are at FibRgLw97 indices 7 and 8 (after ccpHdd at index 5)
	// But we don't have them in the fib struct yet. For this document, they're 0.
	// TODO: add ccpAtn, ccpEdn to fib struct for full correctness

	// Parse PlcftxbxTxt: (n+1) CPs + n FTXBXS (22 bytes each)
	// Total: (n+1)*4 + n*22 = 4 + 26n
	if uint64(f.fcPlcftxbxTxt)+uint64(f.lcbPlcftxbxTxt) > uint64(len(tableData)) {
		return nil
	}
	txbxData := tableData[f.fcPlcftxbxTxt : f.fcPlcftxbxTxt+f.lcbPlcftxbxTxt]

	if f.lcbPlcftxbxTxt < 4 {
		return nil
	}
	n := (f.lcbPlcftxbxTxt - 4) / 26
	if n == 0 {
		return nil
	}

	// Read CPs
	cps := make([]uint32, n+1)
	for i := uint32(0); i <= n; i++ {
		if i*4+4 > uint32(len(txbxData)) {
			break
		}
		cps[i] = binary.LittleEndian.Uint32(txbxData[i*4:])
	}

	// Extract text for each text box
	var texts []string
	for i := uint32(0); i < n; i++ {
		cpStart := txbxStart + cps[i]
		cpEnd := txbxStart + cps[i+1]

		if int(cpStart) >= len(fullRunes) || int(cpEnd) > len(fullRunes) || cpStart >= cpEnd {
			texts = append(texts, "")
			continue
		}

		text := string(fullRunes[cpStart:cpEnd])
		text = strings.TrimRight(text, "\r\n")
		// Remove control characters
		var clean []rune
		for _, r := range text {
			if r >= 0x20 || r == '\t' {
				clean = append(clean, r)
			}
		}
		texts = append(texts, string(clean))
	}

	// Now find which shapes have text boxes by parsing DggInfo
	// Build SPID -> txid mapping
	spidToTxid := parseDggInfoTextBoxes(tableData, f.fcDggInfo, f.lcbDggInfo)
	if len(spidToTxid) == 0 {
		return nil
	}

	// Build SPID -> CP mapping from PlcSpaMom
	spidToCP := parsePlcSpaMom(tableData, f.fcPlcSpaMom, f.lcbPlcSpaMom)

	// Map text boxes to anchor CPs
	// txid is 1-based index into the text box array (shifted by 16 bits)
	var mappings []textBoxMapping
	for spid, txid := range spidToTxid {
		// txid format: the text box index is (txid >> 16) - 1 (0-based)
		// or sometimes txid is just the 1-based index
		txbxIdx := int(txid>>16) - 1
		if txbxIdx < 0 {
			txbxIdx = int(txid) - 1
		}
		if txbxIdx < 0 || txbxIdx >= len(texts) {
			continue
		}
		text := texts[txbxIdx]
		if text == "" {
			continue
		}

		cp, ok := spidToCP[spid]
		if !ok {
			continue
		}

		mappings = append(mappings, textBoxMapping{
			anchorCP: cp,
			spid:     spid,
			text:     text,
		})
	}

	return mappings
}

// parseDggInfoTextBoxes parses the OfficeArtContent to find shapes with lTxid property.
// Returns SPID -> txid mapping.
func parseDggInfoTextBoxes(tableData []byte, fcDggInfo, lcbDggInfo uint32) map[uint32]uint32 {
	if lcbDggInfo == 0 || uint64(fcDggInfo)+uint64(lcbDggInfo) > uint64(len(tableData)) {
		return nil
	}
	data := tableData[fcDggInfo : fcDggInfo+lcbDggInfo]

	if len(data) < 8 {
		return nil
	}
	recType := binary.LittleEndian.Uint16(data[2:])
	verInst := binary.LittleEndian.Uint16(data[0:])
	ver := verInst & 0x0F
	recLen := binary.LittleEndian.Uint32(data[4:])

	if recType != 0xF000 || ver != 0x0F {
		return nil
	}

	result := make(map[uint32]uint32)
	pos := 8 + int(recLen) // skip DggContainer

	// Scan for DgContainers
	for pos < len(data) {
		found := false
		for pos+8 <= len(data) {
			vi := binary.LittleEndian.Uint16(data[pos:])
			rt := binary.LittleEndian.Uint16(data[pos+2:])
			v := vi & 0x0F
			if rt == 0xF002 && v == 0x0F {
				found = true
				break
			}
			pos++
		}
		if !found {
			break
		}

		rl := binary.LittleEndian.Uint32(data[pos+4:])
		containerEnd := pos + 8 + int(rl)
		if containerEnd > len(data) {
			containerEnd = len(data)
		}

		extractTextBoxShapes(data, pos+8, containerEnd, result)
		pos = containerEnd
	}

	return result
}

// extractTextBoxShapes recursively finds SpContainers with lTxid property.
func extractTextBoxShapes(data []byte, offset, end int, result map[uint32]uint32) {
	for offset+8 <= end {
		verInst := binary.LittleEndian.Uint16(data[offset:])
		recType := binary.LittleEndian.Uint16(data[offset+2:])
		recLen := binary.LittleEndian.Uint32(data[offset+4:])
		ver := verInst & 0x0F

		childEnd := offset + 8 + int(recLen)
		if childEnd > end {
			childEnd = end
		}

		if ver == 0x0F {
			if recType == 0xF004 { // SpContainer
				spid, txid := parseSpContainerForTxid(data, offset+8, childEnd)
				if spid != 0 && txid != 0 {
					result[spid] = txid
				}
			} else {
				extractTextBoxShapes(data, offset+8, childEnd, result)
			}
		}

		offset = childEnd
	}
}

// parseSpContainerForTxid extracts SPID and lTxid from a SpContainer.
func parseSpContainerForTxid(data []byte, offset, end int) (spid, txid uint32) {
	for offset+8 <= end {
		verInst := binary.LittleEndian.Uint16(data[offset:])
		recType := binary.LittleEndian.Uint16(data[offset+2:])
		recLen := binary.LittleEndian.Uint32(data[offset+4:])
		inst := verInst >> 4

		childEnd := offset + 8 + int(recLen)
		if childEnd > end {
			childEnd = end
		}
		recData := data[offset+8 : childEnd]

		if recType == 0xF00A && len(recData) >= 8 {
			spid = binary.LittleEndian.Uint32(recData[0:])
		}
		if recType == 0xF00B { // OPT
			for p := uint16(0); p < inst; p++ {
				off := int(p) * 6
				if off+6 > len(recData) {
					break
				}
				propID := binary.LittleEndian.Uint16(recData[off:])
				propVal := binary.LittleEndian.Uint32(recData[off+2:])
				pid := propID & 0x3FFF
				if pid == 0x0080 { // lTxid
					txid = propVal
				}
			}
		}
		offset = childEnd
	}
	return
}

// applyTextBoxMappings applies text box content to the appropriate paragraphs.
func applyTextBoxMappings(fc *FormattedContent, textBoxMappings []textBoxMapping, shapeMappings []shapeImageMapping) {
	if len(textBoxMappings) == 0 || fc == nil {
		return
	}

	// Build a CP -> paragraph index mapping
	// We need to find which paragraph contains each text box anchor CP
	cpPos := uint32(0)
	type paraRange struct {
		cpStart uint32
		cpEnd   uint32
		idx     int
	}
	var ranges []paraRange
	for i, p := range fc.Paragraphs {
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		pLen := uint32(len([]rune(text)))
		ranges = append(ranges, paraRange{cpStart: cpPos, cpEnd: cpPos + pLen, idx: i})
		cpPos += pLen + 1 // +1 for separator
	}

	// Apply text box text to paragraphs
	for _, tbm := range textBoxMappings {
		for _, pr := range ranges {
			if tbm.anchorCP >= pr.cpStart && tbm.anchorCP <= pr.cpEnd {
				fc.Paragraphs[pr.idx].TextBoxText = tbm.text
				break
			}
		}
	}
}
