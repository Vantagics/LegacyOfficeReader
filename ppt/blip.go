package ppt

import (
	"encoding/binary"
	"fmt"

	"github.com/shakinm/xlsReader/cfb"
	"github.com/shakinm/xlsReader/common"
)

// OfficeArt Blip RecordType constants
const (
	rtBlipEMF   = 0xF01A
	rtBlipWMF   = 0xF01B
	rtBlipPICT  = 0xF01C
	rtBlipJPEG  = 0xF01D
	rtBlipPNG   = 0xF01E
	rtBlipDIB   = 0xF01F
	rtBlipTIFF  = 0xF029
	rtBlipJPEG2 = 0xF02A
)

// OfficeArt FBSE RecordType
const rtFBSE = 0xF007

// Blip metadata header sizes
const (
	blipUIDSize       = 16 // 16 bytes UID (MD4 hash)
	blipMetafileExtra = 34 // EMF/WMF/PICT: 34 bytes extra metadata
	blipBitmapExtra   = 1  // JPEG/PNG/DIB/TIFF: 1 byte tag
	fbseHeaderSize    = 36 // FBSE fixed header size (after RecordHeader)
)

// baseRecInstance maps each recType to its base recInstance value (1 UID).
// If actual recInstance > base, there are 2 UIDs.
var baseRecInstance = map[uint16]uint16{
	rtBlipEMF:   0x3D5,
	rtBlipWMF:   0x216,
	rtBlipPICT:  0x542,
	rtBlipJPEG:  0x46A,
	rtBlipJPEG2: 0x6E2,
	rtBlipPNG:   0x6E0,
	rtBlipDIB:   0x7A8,
	rtBlipTIFF:  0x6E4,
}

// parseFBSEAndBlip parses a single FBSE record and its embedded Blip at the given offset.
// Returns the extracted Image, total bytes consumed (8 + FBSE recLen), and any error.
func parseFBSEAndBlip(data []byte, offset uint32) (common.Image, uint32, error) {
	dataLen := uint32(len(data))

	// Need at least 8 bytes for FBSE RecordHeader
	if dataLen < offset+8 {
		return common.Image{}, 0, fmt.Errorf("not enough data for FBSE RecordHeader")
	}

	// Read FBSE RecordHeader
	fbseRecType := binary.LittleEndian.Uint16(data[offset+2 : offset+4])
	fbseRecLen := binary.LittleEndian.Uint32(data[offset+4 : offset+8])

	totalConsumed := uint32(8) + fbseRecLen

	// Verify this is an FBSE record
	if fbseRecType != rtFBSE {
		return common.Image{}, totalConsumed, fmt.Errorf("expected FBSE recType 0xF007, got 0x%04X", fbseRecType)
	}

	// Need enough data for the full FBSE record
	if dataLen < offset+totalConsumed {
		return common.Image{}, totalConsumed, fmt.Errorf("truncated FBSE record")
	}

	// Need at least FBSE header + embedded Blip RecordHeader
	blipHeaderOffset := offset + 8 + fbseHeaderSize
	if dataLen < blipHeaderOffset+8 {
		return common.Image{}, totalConsumed, fmt.Errorf("not enough data for embedded Blip RecordHeader")
	}

	// Read embedded Blip RecordHeader
	blipRecInstance := binary.LittleEndian.Uint16(data[blipHeaderOffset:blipHeaderOffset+2]) >> 4
	blipRecType := binary.LittleEndian.Uint16(data[blipHeaderOffset+2 : blipHeaderOffset+4])
	blipRecLen := binary.LittleEndian.Uint32(data[blipHeaderOffset+4 : blipHeaderOffset+8])

	// Map recType to ImageFormat and extra metadata size
	var format common.ImageFormat
	var extraBytes uint32
	switch blipRecType {
	case rtBlipEMF:
		format = common.ImageFormatEMF
		extraBytes = blipMetafileExtra
	case rtBlipWMF:
		format = common.ImageFormatWMF
		extraBytes = blipMetafileExtra
	case rtBlipPICT:
		format = common.ImageFormatPICT
		extraBytes = blipMetafileExtra
	case rtBlipJPEG, rtBlipJPEG2:
		format = common.ImageFormatJPEG
		extraBytes = blipBitmapExtra
	case rtBlipPNG:
		format = common.ImageFormatPNG
		extraBytes = blipBitmapExtra
	case rtBlipDIB:
		format = common.ImageFormatDIB
		extraBytes = blipBitmapExtra
	case rtBlipTIFF:
		format = common.ImageFormatTIFF
		extraBytes = blipBitmapExtra
	default:
		return common.Image{}, totalConsumed, fmt.Errorf("unknown Blip recType: 0x%04X", blipRecType)
	}

	// Determine UID count based on recInstance vs base value
	base, ok := baseRecInstance[blipRecType]
	uidCount := uint32(1)
	if ok && blipRecInstance > base {
		uidCount = 2
	}

	// Calculate bytes to skip after Blip RecordHeader: UIDs + format-specific metadata
	skipBytes := uidCount*blipUIDSize + extraBytes

	// Validate that the embedded Blip data fits within the FBSE record
	blipDataStart := blipHeaderOffset + 8 + skipBytes
	blipDataEnd := blipHeaderOffset + 8 + blipRecLen

	if blipDataEnd > offset+totalConsumed {
		return common.Image{}, totalConsumed, fmt.Errorf("embedded Blip exceeds FBSE record boundary")
	}

	if skipBytes > blipRecLen {
		return common.Image{}, totalConsumed, fmt.Errorf("Blip metadata exceeds Blip record length")
	}

	if blipDataStart > dataLen || blipDataEnd > dataLen {
		return common.Image{}, totalConsumed, fmt.Errorf("embedded Blip data out of bounds")
	}

	// Extract image data
	imgData := make([]byte, blipDataEnd-blipDataStart)
	copy(imgData, data[blipDataStart:blipDataEnd])

	return common.Image{Format: format, Data: imgData}, totalConsumed, nil
}

// isBlipRecordType returns true if the recType is a known Blip type.
func isBlipRecordType(recType uint16) bool {
	switch recType {
	case rtBlipEMF, rtBlipWMF, rtBlipPICT, rtBlipJPEG, rtBlipPNG, rtBlipDIB, rtBlipTIFF, rtBlipJPEG2:
		return true
	}
	return false
}

// parseDirectBlip parses a direct Blip record (not wrapped in FBSE) at the given offset.
// Returns the extracted Image, total bytes consumed, and any error.
func parseDirectBlip(data []byte, offset uint32) (common.Image, uint32, error) {
	dataLen := uint32(len(data))
	if dataLen < offset+8 {
		return common.Image{}, 0, fmt.Errorf("not enough data for Blip RecordHeader")
	}

	blipRecInstance := binary.LittleEndian.Uint16(data[offset:offset+2]) >> 4
	blipRecType := binary.LittleEndian.Uint16(data[offset+2 : offset+4])
	blipRecLen := binary.LittleEndian.Uint32(data[offset+4 : offset+8])

	totalConsumed := uint32(8) + blipRecLen
	if offset+totalConsumed > dataLen {
		return common.Image{}, totalConsumed, fmt.Errorf("truncated direct Blip record")
	}

	var format common.ImageFormat
	var extraBytes uint32
	switch blipRecType {
	case rtBlipEMF:
		format = common.ImageFormatEMF
		extraBytes = blipMetafileExtra
	case rtBlipWMF:
		format = common.ImageFormatWMF
		extraBytes = blipMetafileExtra
	case rtBlipPICT:
		format = common.ImageFormatPICT
		extraBytes = blipMetafileExtra
	case rtBlipJPEG, rtBlipJPEG2:
		format = common.ImageFormatJPEG
		extraBytes = blipBitmapExtra
	case rtBlipPNG:
		format = common.ImageFormatPNG
		extraBytes = blipBitmapExtra
	case rtBlipDIB:
		format = common.ImageFormatDIB
		extraBytes = blipBitmapExtra
	case rtBlipTIFF:
		format = common.ImageFormatTIFF
		extraBytes = blipBitmapExtra
	default:
		return common.Image{}, totalConsumed, fmt.Errorf("unknown Blip recType: 0x%04X", blipRecType)
	}

	base, ok := baseRecInstance[blipRecType]
	uidCount := uint32(1)
	if ok && blipRecInstance > base {
		uidCount = 2
	}

	skipBytes := uidCount*blipUIDSize + extraBytes
	if skipBytes > blipRecLen {
		return common.Image{}, totalConsumed, fmt.Errorf("Blip metadata exceeds Blip record length")
	}

	blipDataStart := offset + 8 + skipBytes
	blipDataEnd := offset + 8 + blipRecLen

	if blipDataStart > dataLen || blipDataEnd > dataLen {
		return common.Image{}, totalConsumed, fmt.Errorf("direct Blip data out of bounds")
	}

	imgData := make([]byte, blipDataEnd-blipDataStart)
	copy(imgData, data[blipDataStart:blipDataEnd])

	return common.Image{Format: format, Data: imgData}, totalConsumed, nil
}

// parsePicturesStream parses a PPT Pictures stream containing either FBSE records
// or direct Blip records.
func parsePicturesStream(data []byte) []common.Image {
	var images []common.Image
	offset := uint32(0)
	dataLen := uint32(len(data))

	for offset+8 <= dataLen {
		recType := binary.LittleEndian.Uint16(data[offset+2 : offset+4])

		if recType == rtFBSE {
			// FBSE-wrapped blip
			img, consumed, err := parseFBSEAndBlip(data, offset)
			if err != nil {
				if consumed > 0 {
					offset += consumed
				} else {
					break
				}
				if offset > dataLen {
					break
				}
				continue
			}
			images = append(images, img)
			offset += consumed
		} else if isBlipRecordType(recType) {
			// Direct blip record (not wrapped in FBSE)
			img, consumed, err := parseDirectBlip(data, offset)
			if err != nil {
				if consumed > 0 {
					offset += consumed
				} else {
					break
				}
				if offset > dataLen {
					break
				}
				continue
			}
			images = append(images, img)
			offset += consumed
		} else {
			// Unknown record, try to skip by reading recLen
			recLen := binary.LittleEndian.Uint32(data[offset+4 : offset+8])
			next := offset + 8 + recLen
			if next <= offset || next > dataLen {
				break
			}
			offset = next
		}
	}

	if images == nil {
		return []common.Image{}
	}
	return images
}


// extractImagesFromPpt extracts images from a PPT file using the BStoreContainer
// in the PowerPoint Document stream to correctly map image indices.
// The BStoreContainer's FBSE entries define the canonical ordering that shapes
// reference via pib (1-based). Each FBSE entry either embeds a blip directly
// or uses foDelay to point to a blip offset in the Pictures stream.
// Falls back to sequential Pictures stream parsing if BStoreContainer is not found.
func extractImagesFromPpt(adaptor cfb.Cfb, root, picturesDir *cfb.Directory, pptDocData []byte) []common.Image {
	var picData []byte
	if picturesDir != nil {
		reader, err := adaptor.OpenObject(picturesDir, root)
		if err == nil {
			size := binary.LittleEndian.Uint32(picturesDir.StreamSize[:])
			picData = make([]byte, size)
			reader.Read(picData)
		}
	}

	// Try BStoreContainer-based extraction first
	images := extractImagesViaBStore(pptDocData, picData)
	if len(images) > 0 {
		return images
	}

	// Fallback: sequential Pictures stream parsing
	if len(picData) > 0 {
		return parsePicturesStream(picData)
	}
	return []common.Image{}
}

// rtBStoreContainer is the RecordType for OfficeArtBStoreContainer.
const rtBStoreContainer = 0xF001

// extractImagesViaBStore parses the BStoreContainer from the PowerPoint Document
// stream and builds an images array ordered by FBSE index. Each FBSE entry either
// contains an embedded blip or references one in the Pictures stream via foDelay.
func extractImagesViaBStore(pptDocData []byte, picData []byte) []common.Image {
	pptLen := uint32(len(pptDocData))

	// Scan for BStoreContainer (recType=0xF001, recVer=0x0F)
	var bstoreStart, bstoreEnd uint32
	found := false
	for off := uint32(0); off+8 <= pptLen; off++ {
		recVerInst := binary.LittleEndian.Uint16(pptDocData[off : off+2])
		recType := binary.LittleEndian.Uint16(pptDocData[off+2 : off+4])
		recLen := binary.LittleEndian.Uint32(pptDocData[off+4 : off+8])
		recVer := recVerInst & 0x0F

		if recType == rtBStoreContainer && recVer == 0x0F {
			bstoreStart = off + 8
			bstoreEnd = off + 8 + recLen
			if bstoreEnd > pptLen {
				bstoreEnd = pptLen
			}
			found = true
			break
		}
	}
	if !found {
		return nil
	}

	var images []common.Image
	pos := bstoreStart
	for pos+8 <= bstoreEnd {
		fRecType := binary.LittleEndian.Uint16(pptDocData[pos+2 : pos+4])
		fRecLen := binary.LittleEndian.Uint32(pptDocData[pos+4 : pos+8])

		if fRecType != rtFBSE {
			pos += 8 + fRecLen
			continue
		}

		// Parse FBSE header fields
		// offset+8: btWin32(1) btMacOS(1) rgbUid(16) tag(2) size(4) cRef(4) foDelay(4) unused(4)
		if pos+8+fbseHeaderSize > bstoreEnd {
			break
		}

		fbseSize := binary.LittleEndian.Uint32(pptDocData[pos+28 : pos+32])
		foDelay := binary.LittleEndian.Uint32(pptDocData[pos+36 : pos+40])

		var img common.Image
		parsed := false

		if foDelay != 0 && len(picData) > 0 {
			// Blip is in the Pictures stream at offset foDelay
			if foDelay+8 <= uint32(len(picData)) {
				blipRecType := binary.LittleEndian.Uint16(picData[foDelay+2 : foDelay+4])
				if isBlipRecordType(blipRecType) {
					// Direct blip at this offset
					if i, _, err := parseDirectBlip(picData, foDelay); err == nil {
						img = i
						parsed = true
					}
				} else if blipRecType == rtFBSE {
					// FBSE-wrapped blip
					if i, _, err := parseFBSEAndBlip(picData, foDelay); err == nil {
						img = i
						parsed = true
					}
				}
			}
		} else if foDelay == 0 && fbseSize > 0 {
			// Embedded blip within the FBSE record itself
			blipOff := pos + 8 + fbseHeaderSize
			if blipOff+8 <= bstoreEnd {
				blipRecType := binary.LittleEndian.Uint16(pptDocData[blipOff+2 : blipOff+4])
				if isBlipRecordType(blipRecType) {
					if i, _, err := parseDirectBlip(pptDocData, blipOff); err == nil {
						img = i
						parsed = true
					}
				}
			}
		}

		if !parsed {
			// Empty or unparseable entry - add placeholder
			img = common.Image{Format: common.ImageFormatPNG, Data: nil}
		}

		images = append(images, img)
		pos += 8 + fRecLen
	}

	if len(images) == 0 {
		return nil
	}
	return images
}

