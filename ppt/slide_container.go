package ppt

import (
	"encoding/binary"
	"fmt"
	"unicode/utf16"
)

// PPT record types for slide container parsing
const (
	rtSlideContainer = 0x03EE // 1006
	rtSlideAtom      = 0x03EF // 1007
	rtDrawing        = 0x040C // 1036
	rtSpgrContainer  = 0xF003
	rtSpContainer    = 0xF004
	rtFSPGR          = 0xF009 // OfficeArtFSPGR - group coordinate system
	rtSpAtom         = 0xF00A
	rtClientAnchor   = 0xF010
	rtChildAnchor    = 0xF00F
	rtClientTextbox  = 0xF00D
	rtTextHeaderAtom = 0x0F9F // 3999 - TextHeaderAtom (contains text type)
)

// Shape type constants
const (
	msosptRectangle    = 1
	msosptPictureFrame = 75
	msosptTextBox      = 202
)

// OfficeArtFOPT record types
const (
	rtOfficeArtFOPT          = 0xF00B
	rtOfficeArtSecondaryFOPT = 0xF121
	rtOfficeArtTertiaryFOPT  = 0xF122
)

// OfficeArtFOPT property IDs
const (
	foptRotation      = 0x0004 // rotation (fixedPoint 16.16)
	foptGroupBools    = 0x003F // fFlipH, fFlipV
	foptGeoLeft       = 0x0140 // geometry bounding box left
	foptGeoTop        = 0x0141 // geometry bounding box top
	foptGeoRight      = 0x0142 // geometry bounding box right
	foptGeoBottom     = 0x0143 // geometry bounding box bottom
	foptPVertices     = 0x0145 // freeform vertices (complex)
	foptPSegmentInfo  = 0x0146 // freeform segment info (complex)
	foptDxTextLeft    = 0x0081 // left text margin in EMU
	foptDyTextTop     = 0x0082 // top text margin in EMU
	foptDxTextRight   = 0x0083 // right text margin in EMU
	foptDyTextBottom  = 0x0084 // bottom text margin in EMU
	foptWrapText      = 0x0085 // word wrap mode
	foptAnchorText    = 0x0086 // vertical text anchor
	foptPib           = 0x0104 // blip index (1-based)
	foptCropFromTop   = 0x0100 // image crop from top (1/65536)
	foptCropFromBottom = 0x0101 // image crop from bottom (1/65536)
	foptCropFromLeft  = 0x0102 // image crop from left (1/65536)
	foptCropFromRight = 0x0103 // image crop from right (1/65536)
	foptFillColor     = 0x0181 // fillColor (BGR)
	foptFillBools     = 0x01BF // fNoFillHitTest group
	foptLineColor     = 0x01C0 // lineColor (BGR)
	foptLineDashStyle = 0x01C5 // line dash style
	foptLineWidth     = 0x01CB // lineWidth in EMU
	foptLineBools     = 0x01FF // fNoLineDrawDash group
	foptFillOpacity   = 0x0182 // fill opacity (0-65536, 65536 = fully opaque)
	// Line arrow properties (MS-ODRAW §2.3.8)
	foptLineStartArrowhead   = 0x01D0 // start arrowhead style
	foptLineEndArrowhead     = 0x01D1 // end arrowhead style
	foptLineStartArrowWidth  = 0x01D2 // start arrowhead width
	foptLineStartArrowLength = 0x01D3 // start arrowhead length
	foptLineEndArrowWidth    = 0x01D4 // end arrowhead width
	foptLineEndArrowLength   = 0x01D5 // end arrowhead length
)

// parseSlideContainer parses a SlideContainer at the given offset in data.
// Returns the list of shapes found, the layout type, background, and masterIdRef.
func parseSlideContainer(data []byte, offset uint32, fonts []string) ([]ShapeFormatting, int, SlideBackground, uint32) {
	dataLen := uint32(len(data))
	bg := SlideBackground{ImageIdx: -1}
	if offset+recordHeaderSize > dataLen {
		return nil, 0, bg, 0
	}

	rh, err := readRecordHeader(data, offset)
	if err != nil {
		return nil, 0, bg, 0
	}

	containerEnd := offset + recordHeaderSize + rh.recLen
	if containerEnd > dataLen {
		containerEnd = dataLen
	}

	layoutType := 0
	masterIdRef := uint32(0)
	var shapes []ShapeFormatting

	// Scan sub-records inside the SlideContainer
	pos := offset + recordHeaderSize
	for pos+recordHeaderSize <= containerEnd {
		sub, err := readRecordHeader(data, pos)
		if err != nil {
			break
		}
		subDataStart := pos + recordHeaderSize
		subDataEnd := subDataStart + sub.recLen
		if subDataEnd > containerEnd {
			break
		}

		switch sub.recType {
		case rtSlideAtom:
			// SlideAtom: SSlideLayoutAtom(12 bytes) + masterIdRef(4) + notesIdRef(4) + slideFlags(2) + unused(2)
			if sub.recLen >= 4 {
				layoutType = int(binary.LittleEndian.Uint32(data[subDataStart : subDataStart+4]))
			}
			// masterIdRef is at offset 12 (after 12-byte SSlideLayoutAtom)
			if sub.recLen >= 16 {
				masterIdRef = binary.LittleEndian.Uint32(data[subDataStart+12 : subDataStart+16])
			}
			pos = subDataEnd

		case rtDrawing:
			// Drawing container - step into to find SpgrContainer and background
			shapes = parseDrawingGroup(data, subDataStart, subDataEnd, fonts)
			bg = parseDrawingBackground(data, subDataStart, subDataEnd)
			pos = subDataEnd

		default:
			if sub.recVer() == 0xF {
				pos = subDataStart // step into container
			} else {
				pos = subDataEnd // skip atom
			}
		}
	}

	return shapes, layoutType, bg, masterIdRef
}

// parseDrawingGroup parses a Drawing record to find SpgrContainer and extract shapes.
func parseDrawingGroup(data []byte, start, end uint32, fonts []string) []ShapeFormatting {
	var shapes []ShapeFormatting
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

		if rh.recType == rtSpgrContainer {
			shapes = parseSpgrContainer(data, recDataStart, recDataEnd, fonts, nil)
			return shapes
		}

		if rh.recVer() == 0xF {
			pos = recDataStart
		} else {
			pos = recDataEnd
		}
	}

	return shapes
}

// parseDrawingBackground scans the OfficeArtDgContainer for a background SpContainer.
// In PPT, the background shape is an SpContainer that appears before the SpgrContainer
// in the Drawing container, and its SpAtom has the fBackground flag set (bit 2 of grfPersist).
// It can also be identified by having foptPib (image fill) or foptFillColor (solid fill).
func parseDrawingBackground(data []byte, start, end uint32) SlideBackground {
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

		if rh.recType == rtSpgrContainer {
			// Past the background area, stop looking
			break
		}

		if rh.recType == rtSpContainer {
			// This SpContainer before SpgrContainer is the background shape
			bg = parseBgSpContainer(data, recDataStart, recDataEnd)
			break
		}

		if rh.recVer() == 0xF {
			pos = recDataStart
		} else {
			pos = recDataEnd
		}
	}

	return bg
}

// parseBgSpContainer parses a background SpContainer and extracts fill properties.
func parseBgSpContainer(data []byte, start, end uint32) SlideBackground {
	bg := SlideBackground{ImageIdx: -1}
	isBackground := false

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

		switch rh.recType {
		case rtSpAtom:
			// SpAtom: 4 bytes spid + 4 bytes grfPersist
			// Bit 10 (0x400) of grfPersist = fBackground
			if rh.recLen >= 8 {
				grfPersist := binary.LittleEndian.Uint32(data[recDataStart+4 : recDataStart+8])
				if grfPersist&0x400 != 0 {
					isBackground = true
				}
			}

		case rtOfficeArtFOPT, rtOfficeArtSecondaryFOPT, rtOfficeArtTertiaryFOPT:
			// Parse FOPT properties for fill info
			numProps := rh.recInstance()
			parseBgFOPTProperties(data[recDataStart:recDataEnd], numProps, &bg)
		}

		if rh.recVer() == 0xF {
			pos = recDataStart
		} else {
			pos = recDataEnd
		}
	}

	// Only return background if the fBackground flag was set or we found fill properties
	if isBackground && (bg.FillColor != "" || bg.ImageIdx >= 0) {
		bg.HasBackground = true
	}

	return bg
}

// parseBgFOPTProperties extracts fill-related FOPT properties for background.
func parseBgFOPTProperties(data []byte, numProps uint16, bg *SlideBackground) {
	pos := 0
	for i := uint16(0); i < numProps && pos+6 <= len(data); i++ {
		propID := binary.LittleEndian.Uint16(data[pos : pos+2])
		propVal := binary.LittleEndian.Uint32(data[pos+2 : pos+6])
		basePropID := propID & 0x3FFF
		isComplex := propID&0x8000 != 0

		if !isComplex {
			switch basePropID {
			case foptPib:
				if propVal > 0 {
					bg.ImageIdx = int(propVal) - 1
				}
			case foptFillColor:
				bg.fillColorRaw = propVal
				r := uint8(propVal & 0xFF)
				g := uint8((propVal >> 8) & 0xFF)
				b := uint8((propVal >> 16) & 0xFF)
				bg.FillColor = fmt.Sprintf("%02X%02X%02X", r, g, b)
			}
		}
		pos += 6
	}
}


// groupTransform holds the mapping from a group's internal coordinate system
// to the group's position on the slide (or parent group).
type groupTransform struct {
	// Group's internal coordinate system (from OfficeArtFSPGR)
	grpLeft, grpTop, grpRight, grpBottom int64
	// Group's position in absolute EMU coordinates
	dstLeft, dstTop, dstWidth, dstHeight int64
}

// transformToEMU maps a point from group coordinates to absolute EMU.
func (gt *groupTransform) transformToEMU(anchor *[4]int32) (left, top, width, height int64) {
	// ChildAnchor: [xLeft, yTop, xRight, yBottom] in group coordinate space
	cLeft := int64(anchor[0])
	cTop := int64(anchor[1])
	cRight := int64(anchor[2])
	cBottom := int64(anchor[3])

	if cRight < cLeft {
		cLeft, cRight = cRight, cLeft
	}
	if cBottom < cTop {
		cTop, cBottom = cBottom, cTop
	}

	grpW := gt.grpRight - gt.grpLeft
	grpH := gt.grpBottom - gt.grpTop

	if grpW == 0 || grpH == 0 {
		return cLeft, cTop, cRight - cLeft, cBottom - cTop
	}

	left = gt.dstLeft + (cLeft-gt.grpLeft)*gt.dstWidth/grpW
	top = gt.dstTop + (cTop-gt.grpTop)*gt.dstHeight/grpH
	width = (cRight - cLeft) * gt.dstWidth / grpW
	height = (cBottom - cTop) * gt.dstHeight / grpH
	return
}

// applyTransform sets shape position/size from a group coordinate transform.
func (gt *groupTransform) applyTransform(shape *ShapeFormatting, anchor *[4]int32) {
	l, t, w, h := gt.transformToEMU(anchor)
	shape.Left = int32(l)
	shape.Top = int32(t)
	shape.Width = int32(w)
	shape.Height = int32(h)
}

// parseSpgrContainer recursively parses a shape group container.
// parentGT is the parent group's transform (nil for top-level).
func parseSpgrContainer(data []byte, start, end uint32, fonts []string, parentGT *groupTransform) []ShapeFormatting {
	var shapes []ShapeFormatting

	// First, find the group shape (first SpContainer) to get the coordinate transform
	var gt *groupTransform
	pos := start
	firstSp := true

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

		switch rh.recType {
		case rtSpContainer:
			if firstSp {
				firstSp = false
				gt = parseGroupShape(data, recDataStart, recDataEnd, parentGT)
			} else {
				shape := parseSpContainerInGroup(data, recDataStart, recDataEnd, fonts, gt)
				if shape != nil {
					shapes = append(shapes, *shape)
				}
			}
		case rtSpgrContainer:
			nested := parseSpgrContainer(data, recDataStart, recDataEnd, fonts, gt)
			shapes = append(shapes, nested...)
		}

		pos = recDataEnd
	}

	return shapes
}

// parseGroupShape extracts the group coordinate system (FSPGR) and anchor
// from the first SpContainer in a group. parentGT is used to transform
// ChildAnchors of nested groups.
func parseGroupShape(data []byte, start, end uint32, parentGT *groupTransform) *groupTransform {
	var fspgr *[4]int32
	var clientAnchor *[4]int32
	var childAnchor *[4]int32

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

		switch rh.recType {
		case rtFSPGR:
			if rh.recLen >= 16 {
				var coords [4]int32
				coords[0] = int32(binary.LittleEndian.Uint32(data[recDataStart : recDataStart+4]))
				coords[1] = int32(binary.LittleEndian.Uint32(data[recDataStart+4 : recDataStart+8]))
				coords[2] = int32(binary.LittleEndian.Uint32(data[recDataStart+8 : recDataStart+12]))
				coords[3] = int32(binary.LittleEndian.Uint32(data[recDataStart+12 : recDataStart+16]))
				fspgr = &coords
			}
		case rtClientAnchor:
			if rh.recLen >= 8 {
				var anchor [4]int32
				if rh.recLen >= 16 {
					anchor[0] = int32(binary.LittleEndian.Uint32(data[recDataStart : recDataStart+4]))
					anchor[1] = int32(binary.LittleEndian.Uint32(data[recDataStart+4 : recDataStart+8]))
					anchor[2] = int32(binary.LittleEndian.Uint32(data[recDataStart+8 : recDataStart+12]))
					anchor[3] = int32(binary.LittleEndian.Uint32(data[recDataStart+12 : recDataStart+16]))
				} else {
					anchor[0] = int32(int16(binary.LittleEndian.Uint16(data[recDataStart : recDataStart+2])))
					anchor[1] = int32(int16(binary.LittleEndian.Uint16(data[recDataStart+2 : recDataStart+4])))
					anchor[2] = int32(int16(binary.LittleEndian.Uint16(data[recDataStart+4 : recDataStart+6])))
					anchor[3] = int32(int16(binary.LittleEndian.Uint16(data[recDataStart+6 : recDataStart+8])))
				}
				clientAnchor = &anchor
			}
		case rtChildAnchor:
			if rh.recLen >= 16 {
				var anchor [4]int32
				anchor[0] = int32(binary.LittleEndian.Uint32(data[recDataStart : recDataStart+4]))
				anchor[1] = int32(binary.LittleEndian.Uint32(data[recDataStart+4 : recDataStart+8]))
				anchor[2] = int32(binary.LittleEndian.Uint32(data[recDataStart+8 : recDataStart+12]))
				anchor[3] = int32(binary.LittleEndian.Uint32(data[recDataStart+12 : recDataStart+16]))
				childAnchor = &anchor
			}
		}

		if rh.recVer() == 0xF && rh.recType != rtClientTextbox {
			pos = recDataStart
		} else {
			pos = recDataEnd
		}
	}

	if fspgr == nil {
		return nil
	}

	// Determine the group's position in EMU
	var dstLeft, dstTop, dstWidth, dstHeight int64

	if clientAnchor != nil {
		// ClientAnchor: [top, left, right, bottom] in master units
		top := int64(clientAnchor[0])
		left := int64(clientAnchor[1])
		right := int64(clientAnchor[2])
		bottom := int64(clientAnchor[3])
		if right < left {
			left, right = right, left
		}
		if bottom < top {
			top, bottom = bottom, top
		}
		dstLeft = left * 12700 / 8
		dstTop = top * 12700 / 8
		dstWidth = (right - left) * 12700 / 8
		dstHeight = (bottom - top) * 12700 / 8
	} else if childAnchor != nil && parentGT != nil {
		// ChildAnchor in nested group: transform through parent
		dstLeft, dstTop, dstWidth, dstHeight = parentGT.transformToEMU(childAnchor)
	} else {
		return nil
	}

	return &groupTransform{
		grpLeft:   int64(fspgr[0]),
		grpTop:    int64(fspgr[1]),
		grpRight:  int64(fspgr[2]),
		grpBottom: int64(fspgr[3]),
		dstLeft:   dstLeft,
		dstTop:    dstTop,
		dstWidth:  dstWidth,
		dstHeight: dstHeight,
	}
}

// parseSpContainerInGroup parses a shape container within a group,
// applying the group's coordinate transform to ChildAnchors.
func parseSpContainerInGroup(data []byte, start, end uint32, fonts []string, gt *groupTransform) *ShapeFormatting {
	shape := &ShapeFormatting{ImageIdx: -1, LineDash: -1, TextMarginLeft: -1, TextMarginTop: -1, TextMarginRight: -1, TextMarginBottom: -1, TextAnchor: -1, TextWordWrap: -1, FillOpacity: -1, LineStartArrowWidth: -1, LineStartArrowLength: -1, LineEndArrowWidth: -1, LineEndArrowLength: -1, TextType: -1}
	hasAnchor := false
	var clientAnchor *[4]int32
	var childAnchor *[4]int32

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

		switch rh.recType {
		case rtSpAtom:
			shape.ShapeType = rh.recInstance()
			if shape.ShapeType == msosptTextBox || shape.ShapeType == msosptRectangle {
				shape.IsText = true
			}
			if shape.ShapeType == msosptPictureFrame {
				shape.IsImage = true
			}

		case rtOfficeArtFOPT, rtOfficeArtSecondaryFOPT, rtOfficeArtTertiaryFOPT:
			parseFOPTProperties(data[recDataStart:recDataEnd], rh.recInstance(), shape)

		case rtClientAnchor:
			if rh.recLen >= 16 {
				var anchor [4]int32
				anchor[0] = int32(binary.LittleEndian.Uint32(data[recDataStart : recDataStart+4]))
				anchor[1] = int32(binary.LittleEndian.Uint32(data[recDataStart+4 : recDataStart+8]))
				anchor[2] = int32(binary.LittleEndian.Uint32(data[recDataStart+8 : recDataStart+12]))
				anchor[3] = int32(binary.LittleEndian.Uint32(data[recDataStart+12 : recDataStart+16]))
				clientAnchor = &anchor
			} else if rh.recLen >= 8 {
				var anchor [4]int32
				anchor[0] = int32(int16(binary.LittleEndian.Uint16(data[recDataStart : recDataStart+2])))
				anchor[1] = int32(int16(binary.LittleEndian.Uint16(data[recDataStart+2 : recDataStart+4])))
				anchor[2] = int32(int16(binary.LittleEndian.Uint16(data[recDataStart+4 : recDataStart+6])))
				anchor[3] = int32(int16(binary.LittleEndian.Uint16(data[recDataStart+6 : recDataStart+8])))
				clientAnchor = &anchor
			}

		case rtChildAnchor:
			if rh.recLen >= 16 {
				var anchor [4]int32
				anchor[0] = int32(binary.LittleEndian.Uint32(data[recDataStart : recDataStart+4]))
				anchor[1] = int32(binary.LittleEndian.Uint32(data[recDataStart+4 : recDataStart+8]))
				anchor[2] = int32(binary.LittleEndian.Uint32(data[recDataStart+8 : recDataStart+12]))
				anchor[3] = int32(binary.LittleEndian.Uint32(data[recDataStart+12 : recDataStart+16]))
				childAnchor = &anchor
			}

		case rtClientTextbox:
			parseClientTextbox(data, recDataStart, recDataEnd, shape, fonts)
		}

		if rh.recVer() == 0xF && rh.recType != rtClientTextbox {
			pos = recDataStart
		} else {
			pos = recDataEnd
		}
	}

	// Apply anchor: ClientAnchor takes priority
	if clientAnchor != nil {
		applyClientAnchor(shape, clientAnchor)
		hasAnchor = true
	} else if childAnchor != nil {
		if gt != nil {
			gt.applyTransform(shape, childAnchor)
		} else {
			applyChildAnchor(shape, childAnchor)
		}
		hasAnchor = true
	}

	if !hasAnchor && shape.ShapeType == 0 {
		return nil
	}

	return shape
}

// parseSpContainer parses a single shape container and extracts its properties.
func parseSpContainer(data []byte, start, end uint32, fonts []string) *ShapeFormatting {
	shape := &ShapeFormatting{ImageIdx: -1, LineDash: -1, TextMarginLeft: -1, TextMarginTop: -1, TextMarginRight: -1, TextMarginBottom: -1, TextAnchor: -1, TextWordWrap: -1, FillOpacity: -1, LineStartArrowWidth: -1, LineStartArrowLength: -1, LineEndArrowWidth: -1, LineEndArrowLength: -1, TextType: -1}
	hasAnchor := false
	var clientAnchor *[4]int32 // left, top, width, height
	var childAnchor *[4]int32

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

		switch rh.recType {
		case rtSpAtom:
			// ShapeAtom: first 4 bytes = shape type (as part of spid/grfPersist)
			// Actually: first 4 bytes = spid, next 4 bytes = grfPersist
			// The shape type is in the recInstance of the SpAtom header
			shape.ShapeType = rh.recInstance()
			if shape.ShapeType == msosptTextBox || shape.ShapeType == msosptRectangle {
				shape.IsText = true
			}
			if shape.ShapeType == msosptPictureFrame {
				shape.IsImage = true
			}

		case rtOfficeArtFOPT, rtOfficeArtSecondaryFOPT, rtOfficeArtTertiaryFOPT:
			parseFOPTProperties(data[recDataStart:recDataEnd], rh.recInstance(), shape)

		case rtClientAnchor:
			// ClientAnchor for PPT: 8 bytes (4 x int16) or 16 bytes (4 x int32)
			if rh.recLen >= 16 {
				var anchor [4]int32
				anchor[0] = int32(binary.LittleEndian.Uint32(data[recDataStart : recDataStart+4]))     // top (or left)
				anchor[1] = int32(binary.LittleEndian.Uint32(data[recDataStart+4 : recDataStart+8]))   // left (or top)
				anchor[2] = int32(binary.LittleEndian.Uint32(data[recDataStart+8 : recDataStart+12]))  // right (or width)
				anchor[3] = int32(binary.LittleEndian.Uint32(data[recDataStart+12 : recDataStart+16])) // bottom (or height)
				clientAnchor = &anchor
			} else if rh.recLen >= 8 {
				var anchor [4]int32
				anchor[0] = int32(int16(binary.LittleEndian.Uint16(data[recDataStart : recDataStart+2])))
				anchor[1] = int32(int16(binary.LittleEndian.Uint16(data[recDataStart+2 : recDataStart+4])))
				anchor[2] = int32(int16(binary.LittleEndian.Uint16(data[recDataStart+4 : recDataStart+6])))
				anchor[3] = int32(int16(binary.LittleEndian.Uint16(data[recDataStart+6 : recDataStart+8])))
				clientAnchor = &anchor
			}

		case rtChildAnchor:
			// ChildAnchor: 16 bytes (4 x int32)
			if rh.recLen >= 16 {
				var anchor [4]int32
				anchor[0] = int32(binary.LittleEndian.Uint32(data[recDataStart : recDataStart+4]))
				anchor[1] = int32(binary.LittleEndian.Uint32(data[recDataStart+4 : recDataStart+8]))
				anchor[2] = int32(binary.LittleEndian.Uint32(data[recDataStart+8 : recDataStart+12]))
				anchor[3] = int32(binary.LittleEndian.Uint32(data[recDataStart+12 : recDataStart+16]))
				childAnchor = &anchor
			}

		case rtClientTextbox:
			// ClientTextbox container - parse text and formatting inside
			parseClientTextbox(data, recDataStart, recDataEnd, shape, fonts)
		}

		if rh.recVer() == 0xF && rh.recType != rtClientTextbox {
			pos = recDataStart // step into container
		} else {
			pos = recDataEnd
		}
	}

	// Apply anchor: ClientAnchor takes priority over ChildAnchor
	if clientAnchor != nil {
		applyClientAnchor(shape, clientAnchor)
		hasAnchor = true
	} else if childAnchor != nil {
		applyChildAnchor(shape, childAnchor)
		hasAnchor = true
	}

	// Skip shapes without any anchor (e.g., the group shape itself)
	if !hasAnchor && shape.ShapeType == 0 {
		return nil
	}

	return shape
}

// parseFOPTProperties extracts shape properties from an OfficeArtFOPT record.
func parseFOPTProperties(data []byte, numProps uint16, shape *ShapeFormatting) {
	// Fixed property table: numProps × 6 bytes (2 propID + 4 value)
	// Complex properties (bit 15 set) have their value as a length of extra data
	// appended after the fixed table.
	fixedEnd := int(numProps) * 6

	// First pass: read simple properties and record complex property offsets
	type complexProp struct {
		basePropID uint16
		dataLen    uint32
	}
	var complexProps []complexProp

	pos := 0
	for i := uint16(0); i < numProps && pos+6 <= len(data); i++ {
		propID := binary.LittleEndian.Uint16(data[pos : pos+2])
		propVal := binary.LittleEndian.Uint32(data[pos+2 : pos+6])
		basePropID := propID & 0x3FFF
		isComplex := propID&0x8000 != 0

		if isComplex {
			complexProps = append(complexProps, complexProp{basePropID: basePropID, dataLen: propVal})
		} else {
			switch basePropID {
			case foptPib:
				if propVal > 0 {
					shape.IsImage = true
					shape.ImageIdx = int(propVal) - 1
				}
			case foptCropFromTop:
				shape.CropFromTop = int32(propVal)
			case foptCropFromBottom:
				shape.CropFromBottom = int32(propVal)
			case foptCropFromLeft:
				shape.CropFromLeft = int32(propVal)
			case foptCropFromRight:
				shape.CropFromRight = int32(propVal)
			case foptFillColor:
				r := uint8(propVal & 0xFF)
				g := uint8((propVal >> 8) & 0xFF)
				b := uint8((propVal >> 16) & 0xFF)
				shape.FillColor = fmt.Sprintf("%02X%02X%02X", r, g, b)
				shape.FillColorRaw = propVal
			case foptFillBools:
				if propVal&0x00100000 != 0 && propVal&0x00000010 == 0 {
					shape.NoFill = true
				}
			case foptLineColor:
				r := uint8(propVal & 0xFF)
				g := uint8((propVal >> 8) & 0xFF)
				b := uint8((propVal >> 16) & 0xFF)
				shape.LineColor = fmt.Sprintf("%02X%02X%02X", r, g, b)
				shape.LineColorRaw = propVal
			case foptLineWidth:
				shape.LineWidth = int32(propVal)
			case foptLineDashStyle:
				shape.LineDash = int32(propVal)
			case foptLineBools:
				if propVal&0x00080000 != 0 && propVal&0x00000008 == 0 {
					shape.NoLine = true
				}
			case foptRotation:
				shape.Rotation = int32(propVal)
			case foptGroupBools:
				if propVal&0x00010000 != 0 {
					shape.FlipH = propVal&0x0001 != 0
				}
				if propVal&0x00020000 != 0 {
					shape.FlipV = propVal&0x0002 != 0
				}
			case foptDxTextLeft:
				shape.TextMarginLeft = int32(propVal)
			case foptDyTextTop:
				shape.TextMarginTop = int32(propVal)
			case foptDxTextRight:
				shape.TextMarginRight = int32(propVal)
			case foptDyTextBottom:
				shape.TextMarginBottom = int32(propVal)
			case foptAnchorText:
				shape.TextAnchor = int32(propVal)
			case foptWrapText:
				shape.TextWordWrap = int32(propVal)
			case foptFillOpacity:
				shape.FillOpacity = int32(propVal)
			case foptGeoLeft:
				shape.GeoLeft = int32(propVal)
			case foptGeoTop:
				shape.GeoTop = int32(propVal)
			case foptGeoRight:
				shape.GeoRight = int32(propVal)
			case foptGeoBottom:
				shape.GeoBottom = int32(propVal)
			case foptLineStartArrowhead:
				shape.LineStartArrowHead = int32(propVal)
			case foptLineEndArrowhead:
				shape.LineEndArrowHead = int32(propVal)
			case foptLineStartArrowWidth:
				shape.LineStartArrowWidth = int32(propVal)
			case foptLineStartArrowLength:
				shape.LineStartArrowLength = int32(propVal)
			case foptLineEndArrowWidth:
				shape.LineEndArrowWidth = int32(propVal)
			case foptLineEndArrowLength:
				shape.LineEndArrowLength = int32(propVal)
			}
		}
		pos += 6
	}

	// Second pass: read complex property data from after the fixed table
	complexOffset := fixedEnd
	for _, cp := range complexProps {
		dataEnd := complexOffset + int(cp.dataLen)
		if dataEnd > len(data) {
			break
		}
		cpData := data[complexOffset:dataEnd]

		// For IMsoArray properties (pVertices, pSegmentInfo), the dataLen in the
		// FOPT may not include the 6-byte IMsoArray header when cbElem >= 0xFFF0.
		// In this case, the actual data extends 6 bytes beyond dataLen.
		// Detect this by checking the IMsoArray header's cbElem field.
		actualDataLen := int(cp.dataLen)
		if len(cpData) >= 6 {
			arrayCbElem := binary.LittleEndian.Uint16(cpData[4:6])
			if arrayCbElem >= 0xFFF0 {
				// cbElem >= 0xFFF0 indicates truncated elements. The FOPT dataLen
				// only covers nElems * actualElemSize, not the 6-byte header.
				// Extend the data to include the header.
				actualDataLen = int(cp.dataLen) + 6
				extEnd := complexOffset + actualDataLen
				if extEnd <= len(data) {
					cpData = data[complexOffset:extEnd]
				}
			}
		}

		switch cp.basePropID {
		case foptPVertices:
			shape.GeoVertices = parseGeoVertices(cpData)
		case foptPSegmentInfo:
			shape.GeoSegments = parseGeoSegments(cpData)
		}

		complexOffset = complexOffset + actualDataLen
	}
}

// applyClientAnchor sets shape position/size from a ClientAnchor (SmallRectStruct).
// Field order: [top, left, right, bottom] in master units (1/576 inch).
func applyClientAnchor(shape *ShapeFormatting, anchor *[4]int32) {
	// PPT ClientAnchor (SmallRectStruct): [top, left, right, bottom] in master units
	// Convert to EMU: multiply by 12700 / 8, using int64 to avoid overflow
	top := int64(anchor[0])
	left := int64(anchor[1])
	right := int64(anchor[2])
	bottom := int64(anchor[3])

	if right < left {
		left, right = right, left
	}
	if bottom < top {
		top, bottom = bottom, top
	}

	shape.Left = int32(left * 12700 / 8)
	shape.Top = int32(top * 12700 / 8)
	shape.Width = int32((right - left) * 12700 / 8)
	shape.Height = int32((bottom - top) * 12700 / 8)
}

func applyChildAnchor(shape *ShapeFormatting, anchor *[4]int32) {
	// ChildAnchor (OfficeArtChildAnchor): [xLeft, yTop, xRight, yBottom] in master units
	// Convert to EMU: multiply by 12700 / 8, using int64 to avoid overflow
	left := int64(anchor[0])
	top := int64(anchor[1])
	right := int64(anchor[2])
	bottom := int64(anchor[3])

	if right < left {
		left, right = right, left
	}
	if bottom < top {
		top, bottom = bottom, top
	}

	shape.Left = int32(left * 12700 / 8)
	shape.Top = int32(top * 12700 / 8)
	shape.Width = int32((right - left) * 12700 / 8)
	shape.Height = int32((bottom - top) * 12700 / 8)
}



// parseClientTextbox parses text atoms and StyleTextPropAtom inside a ClientTextbox.
func parseClientTextbox(data []byte, start, end uint32, shape *ShapeFormatting, fonts []string) {
	var textContent string
	var styleData []byte
	shape.TextType = -1 // default: unknown

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

		switch rh.recType {
		case rtTextHeaderAtom:
			// TextHeaderAtom: 4 bytes containing textType
			// 0=title, 1=body, 2=notes, 4=other, 5=centerBody, 6=centerTitle
			if rh.recLen >= 4 {
				shape.TextType = int(binary.LittleEndian.Uint32(data[recDataStart : recDataStart+4]))
			}
		case rtTextCharsAtom:
			if rh.recLen > 0 {
				textContent = decodeUTF16LELocal(data[recDataStart:recDataEnd])
			}
		case rtTextBytesAtom:
			if rh.recLen > 0 {
				textContent = decodeANSILocal(data[recDataStart:recDataEnd])
			}
		case rtStyleTextPropAtom:
			styleData = make([]byte, rh.recLen)
			copy(styleData, data[recDataStart:recDataEnd])
		}

		pos = recDataEnd
	}

	if textContent == "" {
		return
	}

	// Parse style properties if available
	if len(styleData) > 0 {
		paras, runs := parseStyleTextPropAtomWithCounts(styleData, len([]rune(textContent)), fonts)
		shape.Paragraphs = buildFormattedParagraphs(textContent, paras, runs)
	} else {
		// No formatting - create a single paragraph with the text
		lines := splitByNewline(textContent)
		result := make([]SlideParagraph, len(lines))
		for i, line := range lines {
			result[i] = SlideParagraph{Runs: []SlideTextRun{{Text: line}}}
		}
		shape.Paragraphs = result
	}
}

// buildFormattedParagraphs combines text content with parsed paragraph and character
// properties to produce formatted paragraphs with text runs.
// It uses character counts from the StyleTextPropAtom to properly distribute
// character formatting runs across paragraph boundaries.
func buildFormattedParagraphs(text string, paras []paraWithCount, runs []runWithCount) []SlideParagraph {
	runes := []rune(text)
	textLen := len(runes)

	if len(paras) == 0 && len(runs) == 0 {
		lines := splitByNewline(text)
		result := make([]SlideParagraph, len(lines))
		for i, line := range lines {
			result[i] = SlideParagraph{Runs: []SlideTextRun{{Text: line}}}
		}
		return result
	}

	// Build a flat list of character-level formatting spans.
	// Each span covers [startChar, startChar+count) in the rune array.
	type charSpan struct {
		start int
		end   int
		run   SlideTextRun
	}
	var charSpans []charSpan
	charPos := 0
	for _, rc := range runs {
		end := charPos + rc.Count
		if end > textLen {
			end = textLen
		}
		charSpans = append(charSpans, charSpan{start: charPos, end: end, run: rc.Run})
		charPos = end
	}
	// If runs don't cover all text, add an unformatted span for the rest
	if charPos < textLen {
		charSpans = append(charSpans, charSpan{start: charPos, end: textLen, run: SlideTextRun{}})
	}

	// Walk through paragraphs by character count, splitting at newlines.
	// Each TextPFRun's count includes the trailing CR/LF character.
	var result []SlideParagraph
	textPos := 0 // current position in runes
	spanIdx := 0 // current index into charSpans

	for _, pc := range paras {
		paraEnd := textPos + pc.Count
		if paraEnd > textLen {
			paraEnd = textLen
		}

		// Find all lines within this paragraph's character range
		// (a single TextPFRun may contain multiple newlines in some edge cases,
		// but typically covers exactly one paragraph + its trailing newline)
		lineStart := textPos
		for lineStart < paraEnd {
			// Find the end of this line (next newline or paraEnd)
			lineEnd := paraEnd
			for i := lineStart; i < paraEnd; i++ {
				if runes[i] == '\r' || runes[i] == '\n' {
					lineEnd = i
					break
				}
			}

			para := pc.Para
			para.Runs = nil

			// Collect character runs that overlap with [lineStart, lineEnd)
			for spanIdx < len(charSpans) && charSpans[spanIdx].start < lineEnd {
				span := charSpans[spanIdx]
				// Calculate the overlap between this span and the current line
				overlapStart := span.start
				if overlapStart < lineStart {
					overlapStart = lineStart
				}
				overlapEnd := span.end
				if overlapEnd > lineEnd {
					overlapEnd = lineEnd
				}

				if overlapStart < overlapEnd {
					r := span.run
					r.Text = string(runes[overlapStart:overlapEnd])
					para.Runs = append(para.Runs, r)
				}

				// If this span extends beyond the current line, don't advance spanIdx
				if span.end > lineEnd {
					break
				}
				spanIdx++
			}

			if len(para.Runs) == 0 {
				para.Runs = []SlideTextRun{{Text: ""}}
			}

			result = append(result, para)

			// Skip past the newline character(s)
			lineStart = lineEnd
			if lineStart < paraEnd && runes[lineStart] == '\r' {
				lineStart++
				if lineStart < paraEnd && runes[lineStart] == '\n' {
					lineStart++
				}
			} else if lineStart < paraEnd && runes[lineStart] == '\n' {
				lineStart++
			}
		}

		textPos = paraEnd
	}

	// Handle any remaining text not covered by paragraph props
	if textPos < textLen {
		remaining := string(runes[textPos:])
		lines := splitByNewline(remaining)
		for _, line := range lines {
			result = append(result, SlideParagraph{
				Runs: []SlideTextRun{{Text: line}},
			})
		}
	}

	if len(result) == 0 {
		result = []SlideParagraph{{Runs: []SlideTextRun{{Text: text}}}}
	}

	return result
}


// splitByNewline splits text by \r, \n, or \r\n.
func splitByNewline(text string) []string {
	var lines []string
	start := 0
	runes := []rune(text)
	for i := 0; i < len(runes); i++ {
		if runes[i] == '\r' || runes[i] == '\n' {
			lines = append(lines, string(runes[start:i]))
			if runes[i] == '\r' && i+1 < len(runes) && runes[i+1] == '\n' {
				i++
			}
			start = i + 1
		}
	}
	if start <= len(runes) {
		lines = append(lines, string(runes[start:]))
	}
	if len(lines) == 0 {
		lines = []string{text}
	}
	return lines
}

// decodeUTF16LELocal decodes UTF-16LE bytes to string, handling surrogate pairs.
func decodeUTF16LELocal(data []byte) string {
	if len(data) < 2 {
		return ""
	}
	numChars := len(data) / 2
	chars := make([]uint16, numChars)
	for i := 0; i < numChars; i++ {
		chars[i] = binary.LittleEndian.Uint16(data[i*2 : i*2+2])
	}
	// Trim trailing nulls
	for len(chars) > 0 && chars[len(chars)-1] == 0 {
		chars = chars[:len(chars)-1]
	}
	// Use utf16.Decode to properly handle surrogate pairs
	return string(utf16.Decode(chars))
}


// decodeANSILocal decodes ANSI bytes to string (local helper).
func decodeANSILocal(data []byte) string {
	// Simple byte-to-rune conversion for ASCII/Latin-1
	runes := make([]rune, 0, len(data))
	for _, b := range data {
		if b == 0 {
			break
		}
		runes = append(runes, rune(b))
	}
	return string(runes)
}


// parseGeoVertices parses the pVertices complex property data.
// Format per [MS-ODRAW] IMsoArray: 2 bytes nElems, 2 bytes nElemsAlloc, 2 bytes cbElem,
// then nElems × cbElem bytes of vertex data.
// cbElem is typically 4 (2×int16) or 8 (2×int32). 0xFFF0 means truncated 8-byte (4 bytes stored).
func parseGeoVertices(data []byte) []GeoVertex {
	if len(data) < 6 {
		return nil
	}
	nElems := int(binary.LittleEndian.Uint16(data[0:2]))
	// nElemsAlloc at data[2:4] — skip
	cbElem := int(binary.LittleEndian.Uint16(data[4:6]))

	if cbElem == 0 || nElems == 0 {
		return nil
	}

	// 0xFFF0 means truncated 8-byte elements (only 4 low-order bytes stored)
	actualSize := cbElem
	if cbElem == 0xFFF0 {
		actualSize = 4
	}

	vertices := make([]GeoVertex, 0, nElems)
	offset := 6
	for i := 0; i < nElems && offset+actualSize <= len(data); i++ {
		var v GeoVertex
		if cbElem == 8 || cbElem == 0xFFF0 {
			// 2 × int32 (or truncated: 2 × int16 stored in 4 bytes)
			if actualSize >= 8 {
				v.X = int32(binary.LittleEndian.Uint32(data[offset : offset+4]))
				v.Y = int32(binary.LittleEndian.Uint32(data[offset+4 : offset+8]))
			} else {
				// truncated: 2 × int16
				v.X = int32(int16(binary.LittleEndian.Uint16(data[offset : offset+2])))
				v.Y = int32(int16(binary.LittleEndian.Uint16(data[offset+2 : offset+4])))
			}
		} else if cbElem == 4 {
			// 2 × int16
			v.X = int32(int16(binary.LittleEndian.Uint16(data[offset : offset+2])))
			v.Y = int32(int16(binary.LittleEndian.Uint16(data[offset+2 : offset+4])))
		} else if cbElem == 2 {
			// 2 × int8
			v.X = int32(int8(data[offset]))
			v.Y = int32(int8(data[offset+1]))
		} else {
			break
		}
		vertices = append(vertices, v)
		offset += actualSize
	}
	return vertices
}

// parseGeoSegments parses the pSegmentInfo complex property data.
// Format per [MS-ODRAW] IMsoArray: 2 bytes nElems, 2 bytes nElemsAlloc, 2 bytes cbElem,
// then nElems × cbElem bytes.
// Each segment is a uint16 encoding per MSOPATHINFO:
// bits 15-13 = segment type, bits 12-0 = count.
// Segment types: 0=lineTo, 1=curveTo, 2=moveTo, 3=close, 4=end, 5=escape
func parseGeoSegments(data []byte) []GeoSegment {
	if len(data) < 6 {
		return nil
	}
	nElems := int(binary.LittleEndian.Uint16(data[0:2]))
	// nElemsAlloc at data[2:4] — skip
	cbElem := int(binary.LittleEndian.Uint16(data[4:6]))
	if cbElem == 0 {
		cbElem = 2
	}

	segments := make([]GeoSegment, 0, nElems)
	offset := 6
	for i := 0; i < nElems && offset+cbElem <= len(data); i++ {
		val := binary.LittleEndian.Uint16(data[offset : offset+2])
		segType := (val >> 13) & 0x07
		count := val & 0x1FFF
		segments = append(segments, GeoSegment{SegType: segType, Count: count})
		offset += cbElem
	}
	return segments
}
