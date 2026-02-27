package doc

import (
	"encoding/binary"
	"errors"
	"io"

	"github.com/shakinm/xlsReader/cfb"
)

// OpenFile opens a DOC document from a file path.
func OpenFile(fileName string) (Document, error) {
	adaptor, err := cfb.OpenFile(fileName)
	defer adaptor.CloseFile()
	if err != nil {
		return Document{}, err
	}
	return openCfb(adaptor)
}

// OpenReader opens a DOC document from an io.ReadSeeker.
func OpenReader(reader io.ReadSeeker) (Document, error) {
	adaptor, err := cfb.OpenReader(reader)
	if err != nil {
		return Document{}, err
	}
	return openCfb(adaptor)
}

// lidToCodepage maps a Windows language ID (LCID) to its default ANSI codepage.
func lidToCodepage(lid uint16) uint16 {
	switch lid & 0x03FF { // primary language ID
	case 0x04: // Chinese
		sublang := lid >> 10
		if sublang == 1 { // Simplified Chinese
			return 936
		}
		return 950 // Traditional Chinese
	case 0x11: // Japanese
		return 932
	case 0x12: // Korean
		return 949
	case 0x19: // Russian
		return 1251
	case 0x15: // Polish
		return 1250
	case 0x0E: // Hungarian
		return 1250
	case 0x05: // Czech
		return 1250
	case 0x1F: // Turkish
		return 1254
	case 0x08: // Greek
		return 1253
	case 0x0D: // Hebrew
		return 1255
	case 0x01: // Arabic
		return 1256
	case 0x1E: // Thai
		return 874
	case 0x2A: // Vietnamese
		return 1258
	default:
		return 1252 // Western European default
	}
}

func openCfb(adaptor cfb.Cfb) (Document, error) {
	var wordDoc *cfb.Directory
	var table0 *cfb.Directory
	var table1 *cfb.Directory
	var root *cfb.Directory
	var dataDir *cfb.Directory
	var picturesDir *cfb.Directory

	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "WordDocument":
			wordDoc = dir
		case "0Table":
			table0 = dir
		case "1Table":
			table1 = dir
		case "Root Entry":
			root = dir
		case "Data":
			dataDir = dir
		case "Pictures":
			picturesDir = dir
		}
	}

	if wordDoc == nil {
		return Document{}, errors.New("WordDocument stream not found")
	}

	// Read WordDocument stream
	wordDocReader, err := adaptor.OpenObject(wordDoc, root)
	if err != nil {
		return Document{}, err
	}
	wordDocSize := binary.LittleEndian.Uint32(wordDoc.StreamSize[:])
	wordDocData := make([]byte, wordDocSize)
	if _, err := wordDocReader.Read(wordDocData); err != nil {
		return Document{}, err
	}

	// Parse FIB to determine which table stream to use
	f, err := parseFIB(wordDocData)
	if err != nil {
		return Document{}, err
	}

	// Select table stream based on fWhichTblStm flag
	var tableDir *cfb.Directory
	if f.fWhichTblStm == 1 {
		tableDir = table1
	} else {
		tableDir = table0
	}
	if tableDir == nil {
		return Document{}, errors.New("table stream not found")
	}

	// Read table stream
	tableReader, err := adaptor.OpenObject(tableDir, root)
	if err != nil {
		return Document{}, err
	}
	tableSize := binary.LittleEndian.Uint32(tableDir.StreamSize[:])
	tableData := make([]byte, tableSize)
	if _, err := tableReader.Read(tableData); err != nil {
		return Document{}, err
	}

	// Extract text using piece table
	pieces, err := parsePieceTable(tableData, &f)
	if err != nil {
		return Document{}, err
	}

	// Build text from pieces using the correct codepage
	codepage := lidToCodepage(f.lid)
	fullText, err := extractTextFromPiecesWithCodepage(wordDocData, pieces, codepage)
	if err != nil {
		return Document{}, err
	}

	// Separate main body text from header/footer/footnote text.
	// Per [MS-DOC], the character stream order is: main text, footnote text, header text, ...
	// ccpText is the character count of the main document text (including the final \r).
	// Only the main body text should be used for paragraph building.
	fullRunes := []rune(fullText)
	mainTextEnd := int(f.ccpText)
	// Debug: check full text length
	expectedTotal := int(f.ccpText + f.ccpFtn + f.ccpHdd + 1) // +1 for final \r
	if len(fullRunes) < expectedTotal {
		// Full text is shorter than expected - header/footer text may be missing
		// This can happen if piece table doesn't cover the full character stream
		_ = expectedTotal // suppress unused warning
	}
	if mainTextEnd > len(fullRunes) {
		mainTextEnd = len(fullRunes)
	}
	text := string(fullRunes[:mainTextEnd])

	// Extract images: prefer BSE entries from DggContainer (finds all images),
	// fall back to Pictures/Data stream scanning.
	// BSE images are ordered by BSE index; Data stream images are ordered by appearance.
	// For inline image placement (0x01), Data stream order matches text order.
	// BSE images are kept for potential drawn object mapping.
	bseImages := extractImagesFromBSE(wordDocData, tableData)
	dataImages := extractImagesFromDoc(adaptor, root, dataDir, picturesDir)

	// Read Data stream bytes for shape parsing
	var dataStreamBytes []byte
	if dataDir != nil {
		dReader, dErr := adaptor.OpenObject(dataDir, root)
		if dErr == nil {
			dSize := binary.LittleEndian.Uint32(dataDir.StreamSize[:])
			dataStreamBytes = make([]byte, dSize)
			dReader.Read(dataStreamBytes)
		}
	}

	// Build shape-to-image mappings for drawn objects (0x08)
	shapeMappings := buildShapeImageMappings(tableData, f.fcPlcSpaMom, f.lcbPlcSpaMom, dataStreamBytes, f.fcDggInfo, f.lcbDggInfo)

	// Use BSE images as the canonical image set (indexed by BSE index).
	// Data stream images are a subset used for inline placement.
	// If BSE images are available, use them; otherwise fall back to Data stream.
	images := bseImages
	if len(images) == 0 {
		images = dataImages
	}

	// Build PicLocation-to-image mapping for inline images (0x01)
	// sprmCPicLocation gives an offset into the Data stream where a PICF structure
	// contains a SpContainer and an embedded BSE with the actual image data.
	// The extracted images are appended to the images slice.
	picLocToBSE := buildPicLocationMappingWithImages(dataStreamBytes, &images)

	// Format parsing - best effort, failures don't affect text/images
	var formattedContent *FormattedContent

	fonts := parseSttbfFfn(tableData, f.fcSttbfFfn, f.lcbSttbfFfn)
	styles, stshErr := parseSTSH(tableData, f.fcStshf, f.lcbStshf, fonts)
	charRuns, chpxErr := parsePlcBteChpx(wordDocData, tableData, f.fcPlcfBteChpx, f.lcbPlcfBteChpx, styles, fonts, pieces)
	paraRuns, papxErr := parsePlcBtePapx(wordDocData, tableData, f.fcPlcfBtePapx, f.lcbPlcfBtePapx, pieces)
	lists, _ := parsePlcfLst(tableData, f.fcPlcfLst, f.lcbPlcfLst)
	listOverrides, _ := parsePlfLfo(tableData, f.fcPlfLfo, f.lcbPlfLfo)
	sections := parsePlcfSed(wordDocData, tableData, f.fcPlcfSed, f.lcbPlcfSed)

	// Only build formatted content if at least one format structure parsed successfully
	if stshErr == nil || chpxErr == nil || papxErr == nil {
		formattedContent = buildFormattedContent(text, charRuns, paraRuns, styles, lists, listOverrides, sections, shapeMappings, picLocToBSE)
	}

	// Extract text box content and map to paragraphs
	textBoxMappings := extractTextBoxMappings(fullText, tableData, &f)
	if formattedContent != nil && len(textBoxMappings) > 0 {
		applyTextBoxMappings(formattedContent, textBoxMappings, shapeMappings)
	}

	// Extract header/footer text (best effort) - uses the full text stream
	// Build shape-to-image mappings for header/footer drawn objects
	hdrShapeMappings := buildShapeImageMappings(tableData, f.fcPlcSpaHdr, f.lcbPlcSpaHdr, dataStreamBytes, f.fcDggInfo, f.lcbDggInfo)
	headers, footers, headersRaw, footersRaw, headerImages, footerImages, headerEntries, footerEntries := extractHeaderFooter(fullText, f.ccpText, f.ccpFtn, f.ccpHdd, tableData, f.fcPlcfHdd, f.lcbPlcfHdd, hdrShapeMappings)
	if formattedContent != nil {
		formattedContent.Headers = headers
		formattedContent.Footers = footers
		formattedContent.HeadersRaw = headersRaw
		formattedContent.FootersRaw = footersRaw
		formattedContent.HeaderImages = headerImages
		formattedContent.FooterImages = footerImages
		formattedContent.HeaderEntries = headerEntries
		formattedContent.FooterEntries = footerEntries
	}

	return Document{text: text, images: images, formattedContent: formattedContent, lid: f.lid, codepage: codepage, fonts: fonts, styles: styles, charRuns: charRuns, paraRuns: paraRuns, fcPlcfSed: f.fcPlcfSed, lcbPlcfSed: f.lcbPlcfSed, picLocToBSE: picLocToBSE}, nil
}
