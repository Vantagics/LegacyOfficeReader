package ppt

import (
	"encoding/binary"
	"errors"
	"fmt"
	"io"

	"github.com/shakinm/xlsReader/cfb"
)

// OpenFile opens a PPT presentation from a file path.
func OpenFile(fileName string) (Presentation, error) {
	adaptor, err := cfb.OpenFile(fileName)
	defer adaptor.CloseFile()
	if err != nil {
		return Presentation{}, err
	}
	return openCfb(adaptor)
}

// OpenReader opens a PPT presentation from an io.ReadSeeker.
func OpenReader(reader io.ReadSeeker) (Presentation, error) {
	adaptor, err := cfb.OpenReader(reader)
	if err != nil {
		return Presentation{}, err
	}
	return openCfb(adaptor)
}

func openCfb(adaptor cfb.Cfb) (Presentation, error) {
	var pptDoc *cfb.Directory
	var currentUser *cfb.Directory
	var root *cfb.Directory
	var picturesDir *cfb.Directory

	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "PowerPoint Document":
			pptDoc = dir
		case "Current User":
			currentUser = dir
		case "Root Entry":
			root = dir
		case "Pictures":
			picturesDir = dir
		}
	}

	if pptDoc == nil {
		return Presentation{}, errors.New("PowerPoint Document stream not found")
	}
	if currentUser == nil {
		return Presentation{}, errors.New("Current User stream not found")
	}

	// Read PowerPoint Document stream
	pptDocReader, err := adaptor.OpenObject(pptDoc, root)
	if err != nil {
		return Presentation{}, err
	}
	pptDocSize := binary.LittleEndian.Uint32(pptDoc.StreamSize[:])
	pptDocData := make([]byte, pptDocSize)
	if _, err := pptDocReader.Read(pptDocData); err != nil {
		return Presentation{}, err
	}

	// Read Current User stream
	cuReader, err := adaptor.OpenObject(currentUser, root)
	if err != nil {
		return Presentation{}, err
	}
	cuSize := binary.LittleEndian.Uint32(currentUser.StreamSize[:])
	cuData := make([]byte, cuSize)
	if _, err := cuReader.Read(cuData); err != nil {
		return Presentation{}, err
	}

	// Parse Current User stream to get offsetToCurrentEdit
	offsetToCurrentEdit, err := parseCurrentUser(cuData)
	if err != nil {
		return Presentation{}, err
	}

	// Build persist directory from UserEdit chain
	persistDir, err := buildPersistDirectory(pptDocData, offsetToCurrentEdit)
	if err != nil {
		return Presentation{}, err
	}

	// Parse SlideListWithText to extract slides
	slides, err := parseSlideListWithTextAndPersist(pptDocData, persistDir)
	if err != nil {
		return Presentation{}, err
	}

	// Extract embedded images from Pictures stream
	images := extractImagesFromPpt(adaptor, root, picturesDir, pptDocData)

	// Parse DocumentContainer for fonts and slide size (best-effort)
	fonts := parseFontCollection(pptDocData)
	slideWidth, slideHeight := parseSlideSize(pptDocData)

	// Parse Environment's TextMasterStyleAtom records for document-level text defaults.
	// These provide default text styles per text type when the MainMasterContainer
	// doesn't have them (e.g., text type 4 "other").
	envTextStyles := parseEnvironmentTextStyles(pptDocData, fonts)

	// Parse SlideContainers for shape formatting (best-effort)
	// Parse master slides first (needed for scheme color resolution)
	masters := parseMasters(pptDocData, persistDir, fonts)

	// Merge environment text styles into masters that don't have them
	for ref, m := range masters {
		if m.TextTypeStyles == nil {
			m.TextTypeStyles = make(map[int][5]MasterTextStyle)
		}
		for tt, styles := range envTextStyles {
			if _, ok := m.TextTypeStyles[tt]; !ok {
				m.TextTypeStyles[tt] = styles
			}
		}
		// Also resolve scheme colors in environment-sourced text styles
		for tt, styles := range m.TextTypeStyles {
			changed := false
			for i := range styles {
				if styles[i].Color != "" && styles[i].ColorRaw != 0 {
					resolved := ResolveSchemeColor(styles[i].Color, styles[i].ColorRaw, m.ColorScheme)
					if resolved != styles[i].Color {
						styles[i].Color = resolved
						changed = true
					}
				}
			}
			if changed {
				m.TextTypeStyles[tt] = styles
			}
		}
		masters[ref] = m
	}

	enrichSlidesWithShapes(pptDocData, persistDir, slides, fonts, masters)

	// Also resolve scheme colors in master shapes
	for ref, m := range masters {
		resolveShapeSchemeColors(m.Shapes, m.ColorScheme)
		masters[ref] = m
	}

	return Presentation{
		slides:      slides,
		images:      images,
		fonts:       fonts,
		slideWidth:  slideWidth,
		slideHeight: slideHeight,
		masters:     masters,
	}, nil
}


// PPT record types for slide size
const (
	rtSlideSize    = 0x03DE // 990
	rtDocumentAtom = 0x03E9 // 1001
)

// parseSlideSize scans pptDocData for a DocumentAtom or SlideSize record and returns
// the slide width and height in EMU. Returns default size if not found.
func parseSlideSize(pptDocData []byte) (int32, int32) {
	dataLen := uint32(len(pptDocData))
	offset := uint32(0)

	for offset+recordHeaderSize <= dataLen {
		rh, err := readRecordHeader(pptDocData, offset)
		if err != nil {
			break
		}
		recDataStart := offset + recordHeaderSize
		recDataEnd := recDataStart + rh.recLen
		if recDataEnd > dataLen {
			break
		}

		// DocumentAtom contains slideSizeX(4) + slideSizeY(4) at the start
		if rh.recType == rtDocumentAtom && rh.recLen >= 8 {
			pptWidth := int32(binary.LittleEndian.Uint32(pptDocData[recDataStart : recDataStart+4]))
			pptHeight := int32(binary.LittleEndian.Uint32(pptDocData[recDataStart+4 : recDataStart+8]))
			emuWidth := int64(pptWidth) * 12700 / 8
			emuHeight := int64(pptHeight) * 12700 / 8
			if emuWidth > 0 && emuHeight > 0 {
				return int32(emuWidth), int32(emuHeight)
			}
		}

		if rh.recType == rtSlideSize && rh.recLen >= 8 {
			pptWidth := int32(binary.LittleEndian.Uint32(pptDocData[recDataStart : recDataStart+4]))
			pptHeight := int32(binary.LittleEndian.Uint32(pptDocData[recDataStart+4 : recDataStart+8]))
			emuWidth := int64(pptWidth) * 12700 / 8
			emuHeight := int64(pptHeight) * 12700 / 8
			if emuWidth > 0 && emuHeight > 0 {
				return int32(emuWidth), int32(emuHeight)
			}
		}

		if rh.recVer() == 0xF {
			offset = recDataStart
		} else {
			offset = recDataEnd
		}
	}

	// Default: 10" x 7.5"
	return 9144000, 6858000
}

// Environment container record type
const rtEnvironment = 0x03F2 // 1010

// parseEnvironmentTextStyles scans the DocumentContainer for the Environment
// container and extracts TextMasterStyleAtom records per text type.
// Per [MS-PPT], the Environment provides document-level default text styles
// that apply when the MainMasterContainer doesn't have styles for a text type.
func parseEnvironmentTextStyles(pptDocData []byte, fonts []string) map[int][5]MasterTextStyle {
	result := make(map[int][5]MasterTextStyle)
	dataLen := uint32(len(pptDocData))
	offset := uint32(0)

	for offset+recordHeaderSize <= dataLen {
		rh, err := readRecordHeader(pptDocData, offset)
		if err != nil {
			break
		}
		recDataStart := offset + recordHeaderSize
		recDataEnd := recDataStart + rh.recLen
		if recDataEnd > dataLen {
			break
		}

		if rh.recType == rtEnvironment {
			// Recursively scan inside Environment for TextMasterStyleAtom records
			scanForTextMasterStyles(pptDocData, recDataStart, recDataEnd, fonts, result)
			return result
		}

		if rh.recVer() == 0xF {
			offset = recDataStart
		} else {
			offset = recDataEnd
		}
	}

	return result
}

// scanForTextMasterStyles recursively scans a container for TextMasterStyleAtom records.
func scanForTextMasterStyles(data []byte, start, end uint32, fonts []string, result map[int][5]MasterTextStyle) {
	pos := start
	for pos+recordHeaderSize <= end {
		rh, err := readRecordHeader(data, pos)
		if err != nil {
			break
		}
		recDataStart := pos + recordHeaderSize
		recDataEnd := recDataStart + rh.recLen
		if recDataEnd > end {
			break
		}

		if rh.recType == rtTextMasterStyleAtom && rh.recLen >= 2 {
			textType := int(rh.recInstance())
			parsed := parseTextMasterStyleDataWithType(data[recDataStart:recDataEnd], fonts, textType)
			hasData := false
			for _, s := range parsed {
				if s.FontSize > 0 || s.Color != "" {
					hasData = true
					break
				}
			}
			if hasData {
				result[textType] = parsed
			}
		}

		if rh.recVer() == 0xF {
			// Container: recurse into it
			scanForTextMasterStyles(data, recDataStart, recDataEnd, fonts, result)
			pos = recDataEnd
		} else {
			pos = recDataEnd
		}
	}
}

// enrichSlidesWithShapes parses SlideContainers from the persist directory
// and populates each Slide's shapes and layoutType fields.
func enrichSlidesWithShapes(pptDocData []byte, persistDir map[uint32]uint32, slides []Slide, fonts []string, masters map[uint32]MasterSlide) {
	if persistDir == nil || len(slides) == 0 {
		return
	}

	// Find SlideListWithText to get persist references for each slide
	dataLen := uint32(len(pptDocData))
	offset := uint32(0)

	var persistRefs []uint32

	for offset+recordHeaderSize <= dataLen {
		rh, err := readRecordHeader(pptDocData, offset)
		if err != nil {
			break
		}
		recDataStart := offset + recordHeaderSize
		recDataEnd := recDataStart + rh.recLen
		if recDataEnd > dataLen {
			break
		}

		if rh.recType == rtSlideListWithText {
			// [MS-PPT] 2.4.14.1: only instance 0 = actual slides
			if rh.recInstance() == 0 {
				// Scan for SlidePersistAtom records inside
				pos := recDataStart
				for pos+recordHeaderSize <= recDataEnd {
					sub, err := readRecordHeader(pptDocData, pos)
					if err != nil {
						break
					}
					subDataStart := pos + recordHeaderSize
					subDataEnd := subDataStart + sub.recLen
					if subDataEnd > recDataEnd {
						break
					}
					if sub.recType == rtSlidePersistAtom && sub.recLen >= 4 {
						psrRef := binary.LittleEndian.Uint32(pptDocData[subDataStart : subDataStart+4])
						persistRefs = append(persistRefs, psrRef)
					}
					pos = subDataEnd
				}
			}
		}

		if rh.recVer() == 0xF {
			offset = recDataStart
		} else {
			offset = recDataEnd
		}
	}

	// For each slide, look up its persist reference and parse the SlideContainer
	for i := 0; i < len(slides) && i < len(persistRefs); i++ {
		slideOffset, ok := persistDir[persistRefs[i]]
		if !ok {
			continue
		}
		if slideOffset+recordHeaderSize > dataLen {
			continue
		}

		shapes, layoutType, bg, masterRef := parseSlideContainer(pptDocData, slideOffset, fonts)
		slides[i].shapes = shapes
		slides[i].layoutType = layoutType
		slides[i].background = bg
		slides[i].masterRef = masterRef

		// Resolve scheme colors using the master's color scheme
		if m, ok := masters[masterRef]; ok {
			slides[i].colorScheme = m.ColorScheme
			slides[i].defaultTextStyles = m.DefaultTextStyles
			slides[i].textTypeStyles = m.TextTypeStyles
			resolveShapeSchemeColors(slides[i].shapes, m.ColorScheme)
		}
	}
}


// MainMasterContainer record type
const rtMainMaster = 0x03F8 // 1016

// resolveShapeSchemeColors resolves scheme color references in shapes using the color scheme.
func resolveShapeSchemeColors(shapes []ShapeFormatting, scheme []string) {
	if len(scheme) < 8 {
		return
	}
	for i := range shapes {
		shapes[i].FillColor = ResolveSchemeColor(shapes[i].FillColor, shapes[i].FillColorRaw, scheme)
		shapes[i].LineColor = ResolveSchemeColor(shapes[i].LineColor, shapes[i].LineColorRaw, scheme)
		// Resolve text run colors that may be scheme color references
		for pi := range shapes[i].Paragraphs {
			for ri := range shapes[i].Paragraphs[pi].Runs {
				run := &shapes[i].Paragraphs[pi].Runs[ri]
				if run.Color != "" && run.ColorRaw != 0 {
					run.Color = ResolveSchemeColor(run.Color, run.ColorRaw, scheme)
				}
			}
			// Resolve bullet colors
			if shapes[i].Paragraphs[pi].BulletColor != "" {
				// Bullet colors don't have raw values stored, but they're typically direct RGB
			}
		}
	}
}

// parseMasters scans for MainMasterContainer records and returns parsed master data
// keyed by their slideId (which is what slides reference via masterIdRef).
func parseMasters(pptDocData []byte, persistDir map[uint32]uint32, fonts []string) map[uint32]MasterSlide {
	dataLen := uint32(len(pptDocData))
	masters := make(map[uint32]MasterSlide)

	// Step 1: Build offset → MasterSlide map from MainMasterContainer records
	masterByOffset := make(map[uint32]MasterSlide)
	offset := uint32(0)
	for offset+recordHeaderSize <= dataLen {
		rh, err := readRecordHeader(pptDocData, offset)
		if err != nil {
			break
		}
		recDataStart := offset + recordHeaderSize
		recDataEnd := recDataStart + rh.recLen
		if recDataEnd > dataLen {
			break
		}

		if rh.recType == rtMainMaster {
			mi := MasterSlide{}
			mi.Background = parseMasterDrawingBg(pptDocData, recDataStart, recDataEnd)
			mi.ColorScheme = parseMasterColorScheme(pptDocData, recDataStart, recDataEnd)
			mi.Shapes = parseMasterDrawingShapes(pptDocData, recDataStart, recDataEnd, fonts)
			mi.DefaultTextStyles, mi.TextTypeStyles = parseMasterTextStyles(pptDocData, recDataStart, recDataEnd, fonts)

			// Resolve scheme color in background
			if mi.Background.HasBackground && mi.Background.FillColor != "" && mi.Background.fillColorRaw&0xFF000000 == 0x08000000 {
				idx := int(mi.Background.fillColorRaw & 0xFF)
				if idx >= 0 && idx < len(mi.ColorScheme) {
					mi.Background.FillColor = mi.ColorScheme[idx]
				}
			}

			// Resolve scheme colors in master shapes (fill and line colors)
			resolveShapeSchemeColors(mi.Shapes, mi.ColorScheme)

			// Resolve scheme colors in master text styles
			for i := range mi.DefaultTextStyles {
				if mi.DefaultTextStyles[i].Color != "" && mi.DefaultTextStyles[i].ColorRaw != 0 {
					mi.DefaultTextStyles[i].Color = ResolveSchemeColor(
						mi.DefaultTextStyles[i].Color,
						mi.DefaultTextStyles[i].ColorRaw,
						mi.ColorScheme,
					)
				}
			}

			masterByOffset[offset] = mi
		}

		if rh.recVer() == 0xF {
			offset = recDataStart
		} else {
			offset = recDataEnd
		}
	}

	// Step 2: Parse SlideListWithText instance 1 (master slides) to get slideId→psrRef mapping
	// Then use persistDir to map psrRef→offset, and masterByOffset to get the master data
	offset = 0
	for offset+recordHeaderSize <= dataLen {
		rh, err := readRecordHeader(pptDocData, offset)
		if err != nil {
			break
		}
		recDataStart := offset + recordHeaderSize
		recDataEnd := recDataStart + rh.recLen
		if recDataEnd > dataLen {
			break
		}

		if rh.recType == rtSlideListWithText && rh.recInstance() == 1 {
			// Scan for SlidePersistAtom records
			pos := recDataStart
			for pos+recordHeaderSize <= recDataEnd {
				sub, err := readRecordHeader(pptDocData, pos)
				if err != nil {
					break
				}
				subDataStart := pos + recordHeaderSize
				subDataEnd := subDataStart + sub.recLen
				if subDataEnd > recDataEnd {
					break
				}
				// SlidePersistAtom: psrReference(4) + flags(4) + numTexts(4) + slideId(4) + reserved(4)
				if sub.recType == rtSlidePersistAtom && sub.recLen >= 16 {
					psrRef := binary.LittleEndian.Uint32(pptDocData[subDataStart : subDataStart+4])
					slideId := binary.LittleEndian.Uint32(pptDocData[subDataStart+12 : subDataStart+16])

					// Map slideId → master data via psrRef → persistDir → offset → masterByOffset
					if masterOffset, ok := persistDir[psrRef]; ok {
						if mi, ok := masterByOffset[masterOffset]; ok {
							masters[slideId] = mi
						}
					}
				}
				pos = subDataEnd
			}
		}

		if rh.recVer() == 0xF {
			offset = recDataStart
		} else {
			offset = recDataEnd
		}
	}

	return masters
}

// parseMasterDrawingShapes extracts non-background shapes from a MainMasterContainer's
// Drawing container. These are the template/decorative shapes (images, graphics).
func parseMasterDrawingShapes(data []byte, start, end uint32, fonts []string) []ShapeFormatting {
	pos := start
	for pos+recordHeaderSize <= end {
		rh, err := readRecordHeader(data, pos)
		if err != nil {
			break
		}
		recDataStart := pos + recordHeaderSize
		recDataEnd := recDataStart + rh.recLen
		if recDataEnd > end {
			break
		}

		if rh.recType == rtDrawing {
			return parseDrawingGroup(data, recDataStart, recDataEnd, fonts)
		}

		if rh.recVer() == 0xF {
			pos = recDataStart
		} else {
			pos = recDataEnd
		}
	}
	return nil
}

// parseMasterDrawingBg finds the Drawing container inside a MainMasterContainer
// and extracts the background shape's fill properties.
func parseMasterDrawingBg(data []byte, start, end uint32) SlideBackground {
	pos := start
	for pos+recordHeaderSize <= end {
		rh, err := readRecordHeader(data, pos)
		if err != nil {
			break
		}
		recDataStart := pos + recordHeaderSize
		recDataEnd := recDataStart + rh.recLen
		if recDataEnd > end {
			break
		}

		if rh.recType == rtDrawing {
			return parseMasterDrawingBgShapes(data, recDataStart, recDataEnd)
		}

		if rh.recVer() == 0xF {
			pos = recDataStart
		} else {
			pos = recDataEnd
		}
	}
	return SlideBackground{ImageIdx: -1}
}


// parseMasterDrawingBgShapes scans the Drawing container for SpContainers
// with the fBackground flag and extracts their fill properties.
func parseMasterDrawingBgShapes(data []byte, start, end uint32) SlideBackground {
	bg := SlideBackground{ImageIdx: -1}
	pos := start

	for pos+recordHeaderSize <= end {
		rh, err := readRecordHeader(data, pos)
		if err != nil {
			break
		}
		recDataStart := pos + recordHeaderSize
		recDataEnd := recDataStart + rh.recLen
		if recDataEnd > end {
			break
		}

		if rh.recType == 0xF004 { // rtSpContainer
			candidate := parseBgSpContainer(data, recDataStart, recDataEnd)
			if candidate.HasBackground {
				return candidate
			}
		}

		if rh.recVer() == 0xF {
			pos = recDataStart
		} else {
			pos = recDataEnd
		}
	}

	return bg
}


// ColorSchemeAtom record type
const rtColorSchemeAtom = 0x07F0 // 2032

// TextMasterStyleAtom record type
const rtTextMasterStyleAtom = 0x0FA3 // 4003

// parseMasterColorScheme extracts the color scheme from a MainMasterContainer.
// Returns up to 8 RGB color strings (scheme indices 0-7).
func parseMasterColorScheme(data []byte, start, end uint32) []string {
	pos := start
	for pos+recordHeaderSize <= end {
		rh, err := readRecordHeader(data, pos)
		if err != nil {
			break
		}
		recDataStart := pos + recordHeaderSize
		recDataEnd := recDataStart + rh.recLen
		if recDataEnd > end {
			break
		}

		if rh.recType == rtColorSchemeAtom && rh.recLen >= 32 {
			// ColorSchemeAtom: 8 colors × 4 bytes (RGBX format)
			colors := make([]string, 8)
			for i := 0; i < 8; i++ {
				off := recDataStart + uint32(i*4)
				r := data[off]
				g := data[off+1]
				b := data[off+2]
				colors[i] = fmt.Sprintf("%02X%02X%02X", r, g, b)
			}
			return colors
		}

		if rh.recVer() == 0xF {
			pos = recDataStart
		} else {
			pos = recDataEnd
		}
	}
	return nil
}


// parseMasterTextStyles extracts default text styles from a MainMasterContainer's
// TextMasterStyleAtom (record type 0x0FA3). This provides default font sizes,
// colors, and styles for each indent level (0-4).
func parseMasterTextStyles(data []byte, start, end uint32, fonts []string) ([5]MasterTextStyle, map[int][5]MasterTextStyle) {
	var styles [5]MasterTextStyle
	textTypeStyles := make(map[int][5]MasterTextStyle)

	pos := start
	for pos+recordHeaderSize <= end {
		rh, err := readRecordHeader(data, pos)
		if err != nil {
			break
		}
		recDataStart := pos + recordHeaderSize
		recDataEnd := recDataStart + rh.recLen
		if recDataEnd > end {
			break
		}

		// TextMasterStyleAtom: instance indicates the text type
		// instance 1 = body text (most useful for default styles)
		// instance 0 = title text
		// We prefer body text (instance 1) but fall back to any instance
		if rh.recType == rtTextMasterStyleAtom && rh.recLen >= 2 {
			textType := int(rh.recInstance())
			parsed := parseTextMasterStyleDataWithType(data[recDataStart:recDataEnd], fonts, textType)
			// Only overwrite if we got useful data
			hasData := false
			for _, s := range parsed {
				if s.FontSize > 0 {
					hasData = true
					break
				}
			}
			if hasData {
				// Store per text type
				textType := int(rh.recInstance())
				textTypeStyles[textType] = parsed

				// Body text (instance 1) takes priority for DefaultTextStyles
				if rh.recInstance() == 1 {
					styles = parsed
				} else if styles[0].FontSize == 0 {
					// Use other instances as fallback
					styles = parsed
				}
			}
		}

		if rh.recVer() == 0xF {
			pos = recDataStart
		} else {
			pos = recDataEnd
		}
	}

	return styles, textTypeStyles
}

// parseTextMasterStyleData parses the data portion of a TextMasterStyleAtom.
// Format: numLevels(2) + for each level: indentLevel(2) + paraProps(no count) + charProps(no count)
// Unlike StyleTextPropAtom, TextMasterStyleAtom does NOT have count fields.
func parseTextMasterStyleData(data []byte, fonts []string) [5]MasterTextStyle {
	return parseTextMasterStyleDataWithType(data, fonts, -1)
}

// parseTextMasterStyleDataWithType parses a TextMasterStyleAtom with knowledge of the text type.
// For text types 0 (title) and 1 (body), each level includes an explicit indent level field.
// For text types 2-8, the indent level field is omitted and levels are implicit.
func parseTextMasterStyleDataWithType(data []byte, fonts []string, textType int) [5]MasterTextStyle {
	var styles [5]MasterTextStyle
	if len(data) < 2 {
		return styles
	}

	numLevels := int(binary.LittleEndian.Uint16(data[0:2]))
	if numLevels > 5 {
		numLevels = 5
	}
	pos := 2

	// Per [MS-PPT] 2.9.36, text types 0 (title) and 1 (body) include
	// an explicit indent level field before each level's properties.
	// Text types 2-8 omit this field.
	hasIndentLevel := textType < 0 || textType == 0 || textType == 1

	for level := 0; level < numLevels && pos < len(data); level++ {
		if hasIndentLevel {
			// Read indent level (2 bytes)
			if pos+2 > len(data) {
				break
			}
			pos += 2 // skip indent level
		}

		// Parse paragraph properties (TextPFException without count)
		paraConsumed := parseMasterParagraphProps(data, pos)
		if paraConsumed == 0 {
			break
		}
		pos += paraConsumed

		// Parse character properties (TextCFException without count)
		if pos >= len(data) {
			break
		}
		run, charConsumed := parseMasterCharacterProps(data, pos, fonts)
		if charConsumed == 0 {
			break
		}
		pos += charConsumed

		if level < 5 {
			styles[level] = MasterTextStyle{
				FontSize: run.FontSize,
				FontName: run.FontName,
				Bold:     run.Bold,
				Italic:   run.Italic,
				Color:    run.Color,
				ColorRaw: run.ColorRaw,
			}
		}
	}

	return styles
}

// parseMasterParagraphProps parses a TextPFException (paragraph properties) WITHOUT
// the leading count field. Used for TextMasterStyleAtom parsing.
// Returns the number of bytes consumed.
func parseMasterParagraphProps(data []byte, pos int) int {
	dataLen := len(data)
	start := pos

	if pos+4 > dataLen {
		return 0
	}
	mask := binary.LittleEndian.Uint32(data[pos : pos+4])
	pos += 4

	if mask&uint32(pfHasBullet|pfBulletHasFont|pfBulletHasColor|pfBulletHasSize) != 0 {
		if pos+2 > dataLen {
			return pos - start
		}
		pos += 2
	}
	if mask&uint32(pfBulletChar) != 0 {
		if pos+2 > dataLen {
			return pos - start
		}
		pos += 2
	}
	if mask&uint32(pfBulletFont) != 0 {
		if pos+2 > dataLen {
			return pos - start
		}
		pos += 2
	}
	if mask&uint32(pfBulletSize) != 0 {
		if pos+2 > dataLen {
			return pos - start
		}
		pos += 2
	}
	if mask&uint32(pfBulletColor) != 0 {
		if pos+4 > dataLen {
			return pos - start
		}
		pos += 4
	}
	if mask&uint32(pfAlign) != 0 {
		if pos+2 > dataLen {
			return pos - start
		}
		pos += 2
	}
	if mask&uint32(pfLineSpacing) != 0 {
		if pos+2 > dataLen {
			return pos - start
		}
		pos += 2
	}
	if mask&uint32(pfSpaceBefore) != 0 {
		if pos+2 > dataLen {
			return pos - start
		}
		pos += 2
	}
	if mask&uint32(pfSpaceAfter) != 0 {
		if pos+2 > dataLen {
			return pos - start
		}
		pos += 2
	}
	if mask&uint32(pfLeftMargin) != 0 {
		if pos+2 > dataLen {
			return pos - start
		}
		pos += 2
	}
	if mask&uint32(pfIndent) != 0 {
		if pos+2 > dataLen {
			return pos - start
		}
		pos += 2
	}
	if mask&uint32(pfDefaultTab) != 0 {
		if pos+2 > dataLen {
			return pos - start
		}
		pos += 2
	}
	if mask&uint32(pfTabStops) != 0 {
		if pos+2 > dataLen {
			return pos - start
		}
		tabCount := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		pos += 2
		skip := tabCount * 4
		if pos+skip > dataLen {
			return pos - start
		}
		pos += skip
	}
	if mask&uint32(pfFontAlign) != 0 {
		if pos+2 > dataLen {
			return pos - start
		}
		pos += 2
	}
	if mask&uint32(pfCharWrap|pfWordWrap|pfOverflow) != 0 {
		if pos+2 > dataLen {
			return pos - start
		}
		pos += 2
	}
	if mask&uint32(pfTextDirection) != 0 {
		if pos+2 > dataLen {
			return pos - start
		}
		pos += 2
	}

	return pos - start
}

// parseMasterCharacterProps parses a TextCFException (character properties) WITHOUT
// the leading count field. Used for TextMasterStyleAtom parsing.
// Returns the SlideTextRun and bytes consumed.
func parseMasterCharacterProps(data []byte, pos int, fonts []string) (SlideTextRun, int) {
	run := SlideTextRun{}
	dataLen := len(data)
	start := pos

	if pos+4 > dataLen {
		return run, 0
	}
	mask := binary.LittleEndian.Uint32(data[pos : pos+4])
	pos += 4

	if mask&uint32(cfStyleBits) != 0 {
		if pos+2 > dataLen {
			return run, pos - start
		}
		flags := binary.LittleEndian.Uint16(data[pos : pos+2])
		run.Bold = flags&uint16(cfBold) != 0
		run.Italic = flags&uint16(cfItalic) != 0
		run.Underline = flags&uint16(cfUnderline) != 0
		pos += 2
	}
	if mask&uint32(cfTypeface) != 0 {
		if pos+2 > dataLen {
			return run, pos - start
		}
		fontIdx := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		if fontIdx >= 0 && fontIdx < len(fonts) {
			run.FontName = fonts[fontIdx]
		}
		pos += 2
	}
	if mask&uint32(cfOldEATypeface) != 0 {
		if pos+2 > dataLen {
			return run, pos - start
		}
		pos += 2
	}
	if mask&uint32(cfAnsiTypeface) != 0 {
		if pos+2 > dataLen {
			return run, pos - start
		}
		pos += 2
	}
	if mask&uint32(cfSymbolTypeface) != 0 {
		if pos+2 > dataLen {
			return run, pos - start
		}
		pos += 2
	}
	if mask&uint32(cfSize) != 0 {
		if pos+2 > dataLen {
			return run, pos - start
		}
		run.FontSize = binary.LittleEndian.Uint16(data[pos:pos+2]) * 100
		pos += 2
	}
	if mask&uint32(cfColor) != 0 {
		if pos+4 > dataLen {
			return run, pos - start
		}
		colorVal := binary.LittleEndian.Uint32(data[pos : pos+4])
		run.ColorRaw = colorVal
		r := uint8(colorVal & 0xFF)
		g := uint8((colorVal >> 8) & 0xFF)
		b := uint8((colorVal >> 16) & 0xFF)
		run.Color = fmt.Sprintf("%02X%02X%02X", r, g, b)
		pos += 4
	}
	if mask&uint32(cfPosition) != 0 {
		if pos+2 > dataLen {
			return run, pos - start
		}
		pos += 2
	}

	return run, pos - start
}
