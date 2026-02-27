package doc

import (
	"bytes"
	"compress/zlib"
	"encoding/binary"
	"fmt"
	"io"

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

// Blip metadata header sizes
const (
	blipUIDSize       = 16 // 16 bytes UID (MD4 hash)
	blipMetafileExtra = 34 // EMF/WMF/PICT: 34 bytes extra metadata
	blipBitmapExtra   = 1  // JPEG/PNG/DIB/TIFF: 1 byte tag
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

// parseBlip parses a single OfficeArt Blip record at the given offset.
// Returns the extracted Image, total bytes consumed, and any error.
func parseBlip(data []byte, offset uint32) (common.Image, uint32, error) {
	// Need at least 8 bytes for RecordHeader
	if uint32(len(data)) < offset+8 {
		return common.Image{}, 0, fmt.Errorf("not enough data for RecordHeader")
	}

	// Read RecordHeader fields
	recInstance := binary.LittleEndian.Uint16(data[offset:offset+2]) >> 4
	recType := binary.LittleEndian.Uint16(data[offset+2 : offset+4])
	recLen := binary.LittleEndian.Uint32(data[offset+4 : offset+8])

	// Map recType to ImageFormat and extra metadata size
	var format common.ImageFormat
	var extraBytes uint32
	switch recType {
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
		return common.Image{}, 8 + recLen, fmt.Errorf("unknown recType: 0x%04X", recType)
	}

	// Determine UID count based on recInstance vs base value
	base, ok := baseRecInstance[recType]
	uidCount := uint32(1)
	if ok && recInstance > base {
		uidCount = 2
	}

	// Calculate bytes to skip after RecordHeader: UIDs + format-specific metadata
	skipBytes := uidCount*blipUIDSize + extraBytes

	// Validate that we have enough data for the full record
	totalConsumed := uint32(8) + recLen
	if uint32(len(data)) < offset+totalConsumed {
		return common.Image{}, totalConsumed, fmt.Errorf("truncated blip record")
	}

	// Validate that skipBytes doesn't exceed recLen
	if skipBytes > recLen {
		return common.Image{}, totalConsumed, fmt.Errorf("metadata exceeds record length")
	}

	// Extract image data
	imgStart := offset + 8 + skipBytes
	imgEnd := offset + 8 + recLen
	imgData := make([]byte, imgEnd-imgStart)
	copy(imgData, data[imgStart:imgEnd])

	// EMF, WMF, and PICT blips are zlib-compressed. Decompress them.
	if (format == common.ImageFormatEMF || format == common.ImageFormatWMF || format == common.ImageFormatPICT) && len(imgData) > 2 {
		// Check for zlib header (0x78 0x01, 0x78 0x9C, or 0x78 0xDA)
		if imgData[0] == 0x78 {
			decompressed, err := decompressZlib(imgData)
			if err == nil && len(decompressed) > 0 {
				imgData = decompressed
			}
		}
	}

	return common.Image{Format: format, Data: imgData}, totalConsumed, nil
}

// decompressZlib decompresses zlib-compressed data.
func decompressZlib(data []byte) ([]byte, error) {
	r, err := zlib.NewReader(bytes.NewReader(data))
	if err != nil {
		return nil, err
	}
	defer r.Close()
	return io.ReadAll(r)
}

// parseBlipStream parses a byte stream containing OfficeArt records (possibly
// wrapped in container records like DggContainer / BStoreContainer / BSE).
// It walks the record tree recursively and extracts all Blip images.
// If the stream contains non-OfficeArt data (common in DOC Data streams),
// it scans for OfficeArt record signatures before walking.
func parseBlipStream(data []byte) []common.Image {
	var images []common.Image
	dataLen := uint32(len(data))

	// First try walking from offset 0
	walkOfficeArt(data, 0, dataLen, &images)

	// If no images found, scan for OfficeArt records at arbitrary offsets.
	// The Data stream in DOC files often has non-OfficeArt data before the
	// actual drawing records.
	if len(images) == 0 {
		scanOfficeArtRecords(data, dataLen, &images)
	}

	if images == nil {
		return []common.Image{}
	}
	return images
}

// scanOfficeArtRecords scans the data for OfficeArt BSE or Blip records
// that may be embedded at arbitrary offsets within the stream.
func scanOfficeArtRecords(data []byte, dataLen uint32, images *[]common.Image) {
	for i := uint32(0); i+8 <= dataLen; i++ {
		recType := binary.LittleEndian.Uint16(data[i+2 : i+4])
		verInst := binary.LittleEndian.Uint16(data[i : i+2])
		recVer := verInst & 0x0F
		recLen := binary.LittleEndian.Uint32(data[i+4 : i+8])

		// Validate recLen is reasonable
		if recLen == 0 || i+8+recLen > dataLen {
			continue
		}

		// Container record — recurse into it
		if recVer == 0xF && isOfficeArtType(recType) {
			walkOfficeArt(data, i+8, i+8+recLen, images)
			i += 7 + recLen // loop will add 1
			continue
		}

		// BSE record
		if recType == 0xF007 && recVer == 0x2 {
			parseBSE(data, i, recLen, images)
			i += 7 + recLen
			continue
		}

		// Direct Blip record
		if isBlipType(recType) {
			img, consumed, err := parseBlip(data, i)
			if err == nil {
				*images = append(*images, img)
				i += consumed - 1
				continue
			}
		}
	}
}

// isOfficeArtType returns true if the recType is in the OfficeArt range (0xF000-0xF02A).
func isOfficeArtType(recType uint16) bool {
	return recType >= 0xF000 && recType <= 0xF02A
}

// walkOfficeArt recursively walks OfficeArt records between offset and limit,
// extracting Blip images into the provided slice.
func walkOfficeArt(data []byte, offset, limit uint32, images *[]common.Image) {
	for offset+8 <= limit {
		if offset+8 > uint32(len(data)) {
			break
		}

		verInst := binary.LittleEndian.Uint16(data[offset : offset+2])
		recVer := verInst & 0x0F
		recType := binary.LittleEndian.Uint16(data[offset+2 : offset+4])
		recLen := binary.LittleEndian.Uint32(data[offset+4 : offset+8])

		childStart := offset + 8
		childEnd := childStart + recLen
		if childEnd > limit || childEnd > uint32(len(data)) {
			break
		}

		// Check if this is a Blip record
		if isBlipType(recType) {
			img, _, err := parseBlip(data, offset)
			if err == nil {
				*images = append(*images, img)
			}
			offset = childEnd
			continue
		}

		// If container (recVer == 0xF), recurse into children
		if recVer == 0xF {
			walkOfficeArt(data, childStart, childEnd, images)
			offset = childEnd
			continue
		}

		// BSE record (recType 0xF007) — contains a Blip as child
		if recType == 0xF007 {
			parseBSE(data, offset, recLen, images)
			offset = childEnd
			continue
		}

		// Skip unknown records
		offset = childEnd
	}
}

// isBlipType returns true if the recType is in the OfficeArt Blip range.
func isBlipType(recType uint16) bool {
	return recType >= 0xF01A && recType <= 0xF02A
}

// parseBSE extracts the embedded Blip from a BSE record (recType 0xF007).
// The BSE record header is 8 bytes, followed by 36 bytes of BSE-specific data,
// then the embedded Blip record.
func parseBSE(data []byte, offset, recLen uint32, images *[]common.Image) {
	const bseHeaderSize = uint32(36) // BSE-specific header after the 8-byte record header
	childStart := offset + 8
	childEnd := childStart + recLen

	if recLen > bseHeaderSize {
		blipOffset := childStart + bseHeaderSize
		if blipOffset+8 <= childEnd && blipOffset+8 <= uint32(len(data)) {
			img, _, err := parseBlip(data, blipOffset)
			if err == nil {
				*images = append(*images, img)
			}
		}
	}
}


// extractImagesFromDoc extracts images from a DOC file's Pictures or Data stream.
// Prefers Pictures stream; falls back to Data stream. Returns empty slice if neither exists.
func extractImagesFromDoc(adaptor cfb.Cfb, root, dataDir, picturesDir *cfb.Directory) []common.Image {
	var dir *cfb.Directory
	if picturesDir != nil {
		dir = picturesDir
	} else if dataDir != nil {
		dir = dataDir
	} else {
		return []common.Image{}
	}

	reader, err := adaptor.OpenObject(dir, root)
	if err != nil {
		return []common.Image{}
	}

	size := binary.LittleEndian.Uint32(dir.StreamSize[:])
	data := make([]byte, size)
	_, err = reader.Read(data)
	if err != nil {
		return []common.Image{}
	}

	return parseBlipStream(data)
}

// extractImagesFromBSE extracts images by parsing BSE entries from the DggContainer
// in the Table stream and reading blip data from the WordDocument stream.
// This finds images that are not in the Data/Pictures stream.
func extractImagesFromBSE(wordDocData, tableData []byte) []common.Image {
	// Scan the Table stream for a DggContainer (0xF000)
	var dggStart, dggEnd uint32
	for i := uint32(0); i+8 <= uint32(len(tableData)); i++ {
		verInst := binary.LittleEndian.Uint16(tableData[i : i+2])
		recVer := verInst & 0x0F
		recType := binary.LittleEndian.Uint16(tableData[i+2 : i+4])
		recLen := binary.LittleEndian.Uint32(tableData[i+4 : i+8])

		if recType == 0xF000 && recVer == 0xF && recLen > 0 && i+8+recLen <= uint32(len(tableData)) {
			dggStart = i + 8
			dggEnd = i + 8 + recLen
			break
		}
	}

	if dggStart == 0 {
		return nil
	}

	// Find BStoreContainer (0xF001) inside DggContainer
	var bscStart, bscEnd uint32
	for offset := dggStart; offset+8 <= dggEnd; {
		recType := binary.LittleEndian.Uint16(tableData[offset+2 : offset+4])
		recLen := binary.LittleEndian.Uint32(tableData[offset+4 : offset+8])
		verInst := binary.LittleEndian.Uint16(tableData[offset : offset+2])
		recVer := verInst & 0x0F

		childEnd := offset + 8 + recLen
		if childEnd > dggEnd {
			break
		}

		if recType == 0xF001 && recVer == 0xF {
			bscStart = offset + 8
			bscEnd = childEnd
			break
		}
		offset = childEnd
	}

	if bscStart == 0 {
		return nil
	}

	// Parse BSE entries from BStoreContainer
	var images []common.Image
	for offset := bscStart; offset+8 <= bscEnd; {
		recType := binary.LittleEndian.Uint16(tableData[offset+2 : offset+4])
		recLen := binary.LittleEndian.Uint32(tableData[offset+4 : offset+8])

		childEnd := offset + 8 + recLen
		if childEnd > bscEnd {
			break
		}

		if recType == 0xF007 && offset+8+36 <= bscEnd {
			// BSE record: extract blip from WordDocument stream using foDelay
			bseData := tableData[offset+8 : offset+8+36]
			foDelay := binary.LittleEndian.Uint32(bseData[28:32])

			if foDelay > 0 && foDelay+8 <= uint32(len(wordDocData)) {
				img, _, err := parseBlip(wordDocData, foDelay)
				if err == nil {
					images = append(images, img)
				}
			}
		}

		offset = childEnd
	}

	return images
}

// shapeImageMapping maps a character position (CP) of a drawn object (0x08)
// to a BSE image index (0-based).
type shapeImageMapping struct {
	cp       uint32 // character position of the 0x08 character
	bseIndex int    // 0-based BSE image index
}

// parsePlcSpaMom parses the PlcSpaMom structure from the Table stream.
// Returns a map of SPID -> CP for main document shapes.
func parsePlcSpaMom(tableData []byte, fc, lcb uint32) map[uint32]uint32 {
	if lcb == 0 || uint64(fc)+uint64(lcb) > uint64(len(tableData)) {
		return nil
	}
	spaData := tableData[fc : fc+lcb]
	// SPA is 26 bytes each: lcb = (n+1)*4 + n*26 => n = (lcb - 4) / 30
	if lcb < 4 {
		return nil
	}
	n := (lcb - 4) / 30
	if n == 0 {
		return nil
	}
	result := make(map[uint32]uint32)
	for i := uint32(0); i < n; i++ {
		cp := binary.LittleEndian.Uint32(spaData[i*4:])
		spaOff := (n+1)*4 + i*26
		if spaOff+4 > uint32(len(spaData)) {
			break
		}
		spid := binary.LittleEndian.Uint32(spaData[spaOff:])
		result[spid] = cp
	}
	return result
}

// parseDataStreamShapes scans the Data stream for SpContainer records and
// extracts SPID -> pib (1-based BSE index) mappings.
func parseDataStreamShapes(dataStreamBytes []byte) map[uint32]uint32 {
	if len(dataStreamBytes) == 0 {
		return nil
	}
	result := make(map[uint32]uint32)
	dataLen := uint32(len(dataStreamBytes))

	for i := uint32(0); i+8 <= dataLen; i++ {
		verInst := binary.LittleEndian.Uint16(dataStreamBytes[i : i+2])
		recVer := verInst & 0x0F
		recType := binary.LittleEndian.Uint16(dataStreamBytes[i+2 : i+4])
		recLen := binary.LittleEndian.Uint32(dataStreamBytes[i+4 : i+8])

		if recType == 0xF004 && recVer == 0xF && recLen > 0 && i+8+recLen <= dataLen {
			spid, pib := parseSpContainerForPib(dataStreamBytes, i+8, i+8+recLen)
			if spid != 0 && pib != 0 {
				result[spid] = pib
			}
			i += 7 + recLen // skip past this container
		}
	}
	return result
}

// parseSpContainerForPib parses a single SpContainer to extract SPID and pib values.
func parseSpContainerForPib(data []byte, offset, limit uint32) (spid, pib uint32) {
	for offset+8 <= limit {
		verInst := binary.LittleEndian.Uint16(data[offset : offset+2])
		recVer := verInst & 0x0F
		recInst := verInst >> 4
		recType := binary.LittleEndian.Uint16(data[offset+2 : offset+4])
		recLen := binary.LittleEndian.Uint32(data[offset+4 : offset+8])

		childEnd := offset + 8 + recLen
		if childEnd > limit {
			break
		}

		// Sp record (0xF00A) - shape info
		if recType == 0xF00A && recLen >= 8 {
			spid = binary.LittleEndian.Uint32(data[offset+8:])
		}

		// Opt record (0xF00B) - shape properties
		if recType == 0xF00B && recLen > 0 {
			nProps := recInst
			propOff := offset + 8
			for p := uint16(0); p < nProps && propOff+6 <= childEnd; p++ {
				propID := binary.LittleEndian.Uint16(data[propOff:])
				propVal := binary.LittleEndian.Uint32(data[propOff+2:])
				pid := propID & 0x3FFF
				if pid == 260 { // pib property
					pib = propVal
				}
				propOff += 6
			}
		}

		if recVer == 0xF {
			// Recurse into container children
			s, p := parseSpContainerForPib(data, offset+8, childEnd)
			if s != 0 {
				spid = s
			}
			if p != 0 {
				pib = p
			}
		}

		offset = childEnd
	}
	return
}

// buildShapeImageMappings builds the complete mapping from CP positions of
// drawn objects (0x08) to BSE image indices.
// It combines PlcSpaMom (CP->SPID) with shape data from both the DggInfo
// (OfficeArtContent in table stream) and the Data stream.
// Shapes without images (pib=0, e.g. text boxes, group shapes) are included
// with bseIndex=-1 so they consume their 0x08 slot without mapping to an image.
func buildShapeImageMappings(tableData []byte, fcPlcSpaMom, lcbPlcSpaMom uint32, dataStreamBytes []byte, fcDggInfo, lcbDggInfo uint32) []shapeImageMapping {
	// Step 1: Parse PlcSpaMom to get SPID -> CP
	spidToCP := parsePlcSpaMom(tableData, fcPlcSpaMom, lcbPlcSpaMom)
	if len(spidToCP) == 0 {
		return nil
	}

	// Step 2: Build SPID -> pib mapping from all available sources
	spidToPib := make(map[uint32]uint32)

	// 2a: Parse DggInfo (OfficeArtContent) in table stream for shapes in DgContainers
	parseDggInfoShapes(tableData, fcDggInfo, lcbDggInfo, spidToPib)

	// 2b: Parse Data stream shapes (fallback/additional)
	dataSpidToPib := parseDataStreamShapes(dataStreamBytes)
	for spid, pib := range dataSpidToPib {
		if _, exists := spidToPib[spid]; !exists {
			spidToPib[spid] = pib
		}
	}

	// Step 3: Map PlcSpaMom SPIDs to image indices.
	// Include ALL shapes from PlcSpaMom, even those without images (pib=0),
	// so that 0x08 characters for non-image shapes (text boxes, etc.) are
	// properly consumed and don't shift the image mapping.
	var mappings []shapeImageMapping
	for spid, cp := range spidToCP {
		pib, hasPib := spidToPib[spid]
		if hasPib && pib > 0 {
			mappings = append(mappings, shapeImageMapping{
				cp:       cp,
				bseIndex: int(pib) - 1, // pib is 1-based, convert to 0-based
			})
		} else {
			// Non-image shape (text box, group shape, line, etc.)
			mappings = append(mappings, shapeImageMapping{
				cp:       cp,
				bseIndex: -1,
			})
		}
	}

	// Sort by CP
	for i := 1; i < len(mappings); i++ {
		for j := i; j > 0 && mappings[j].cp < mappings[j-1].cp; j-- {
			mappings[j], mappings[j-1] = mappings[j-1], mappings[j]
		}
	}

	return mappings
}

// parseDggInfoShapes parses the OfficeArtContent in the table stream to extract
// SPID -> pib mappings from all DgContainers.
func parseDggInfoShapes(tableData []byte, fcDggInfo, lcbDggInfo uint32, result map[uint32]uint32) {
	if lcbDggInfo == 0 || uint64(fcDggInfo)+uint64(lcbDggInfo) > uint64(len(tableData)) {
		return
	}
	data := tableData[fcDggInfo : fcDggInfo+lcbDggInfo]

	// OfficeArtContent: DggContainer followed by DgContainers
	// Skip the DggContainer first
	if len(data) < 8 {
		return
	}
	verInst := binary.LittleEndian.Uint16(data[0:])
	recType := binary.LittleEndian.Uint16(data[2:])
	recLen := binary.LittleEndian.Uint32(data[4:])
	ver := verInst & 0x0F

	if recType != 0xF000 || ver != 0x0F {
		return // Not a valid DggContainer
	}

	pos := 8 + int(recLen) // skip past DggContainer

	// Parse remaining DgContainers
	for pos < len(data) {
		// There may be a 1-byte padding/type indicator between records
		// Scan for the next DgContainer (0xF002)
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

		// Parse shapes inside this DgContainer
		extractShapesFromContainer(data, pos+8, containerEnd, result)
		pos = containerEnd
	}
}

// extractShapesFromContainer recursively parses OfficeArt containers to find
// SpContainers and extract their SPID and pib values.
func extractShapesFromContainer(data []byte, offset, end int, result map[uint32]uint32) {
	for offset+8 <= end {
		verInst := binary.LittleEndian.Uint16(data[offset:])
		recType := binary.LittleEndian.Uint16(data[offset+2:])
		recLen := binary.LittleEndian.Uint32(data[offset+4:])
		ver := verInst & 0x0F

		childEnd := offset + 8 + int(recLen)
		if childEnd > end {
			childEnd = end
		}

		if ver == 0x0F { // container
			if recType == 0xF004 { // SpContainer
				spid, pib := parseSpContainerForPib(data, uint32(offset+8), uint32(childEnd))
				if spid != 0 && pib != 0 {
					result[spid] = pib
				}
			} else {
				// Recurse into SpgrContainer, DgContainer, etc.
				extractShapesFromContainer(data, offset+8, childEnd, result)
			}
		}

		offset = childEnd
	}
}

// buildPicLocationMapping scans the Data stream for PICFAndOfficeArtData structures
// and builds a mapping from Data stream offset (sprmCPicLocation value) to BSE index (0-based).
// Each PICF structure has a header (typically 68 bytes) followed by an OfficeArt SpContainer,
// then a BSE record containing the actual embedded blip data.
// The function extracts these embedded images and appends them to the provided images slice,
// returning the mapping from PicLocation offset to the new image index.
func buildPicLocationMapping(dataStreamBytes []byte) map[int32]int {
	return buildPicLocationMappingWithImages(dataStreamBytes, nil)
}

// buildPicLocationMappingWithImages extracts inline images from the Data stream and
// appends them to the images slice. Returns a mapping from PicLocation to image index.
func buildPicLocationMappingWithImages(dataStreamBytes []byte, images *[]common.Image) map[int32]int {
	if len(dataStreamBytes) == 0 {
		return nil
	}
	result := make(map[int32]int)
	dataLen := len(dataStreamBytes)

	// Scan for SpContainers in the Data stream
	for i := 0; i+8 <= dataLen; i++ {
		verInst := binary.LittleEndian.Uint16(dataStreamBytes[i : i+2])
		recVer := verInst & 0x0F
		recType := binary.LittleEndian.Uint16(dataStreamBytes[i+2 : i+4])
		recLen := binary.LittleEndian.Uint32(dataStreamBytes[i+4 : i+8])

		if recType == 0xF004 && recVer == 0xF && recLen > 0 && uint32(i)+8+recLen <= uint32(dataLen) {
			spContainerEnd := i + 8 + int(recLen)

			// Check if there's a PICF header before this SpContainer
			picfOffset := -1
			for _, hdrSize := range []int{68, 44} {
				pOff := i - hdrSize
				if pOff >= 0 && pOff+6 <= dataLen {
					cb := int(binary.LittleEndian.Uint16(dataStreamBytes[pOff+4:]))
					if cb == hdrSize && pOff+cb == i {
						picfOffset = pOff
						break
					}
				}
			}

			if picfOffset >= 0 {
				// Read PICF lcb to know the total structure size
				lcb := int(binary.LittleEndian.Uint32(dataStreamBytes[picfOffset:]))
				picfEnd := picfOffset + lcb

				// After the SpContainer, there should be a BSE record (0xF007)
				// containing the embedded blip
				if spContainerEnd+8 <= picfEnd && spContainerEnd+8 <= dataLen {
					bseRT := binary.LittleEndian.Uint16(dataStreamBytes[spContainerEnd+2:])
					bseRL := binary.LittleEndian.Uint32(dataStreamBytes[spContainerEnd+4:])

					if bseRT == 0xF007 && bseRL > 36 {
						// BSE record: 36-byte header + blip data
						blipStart := uint32(spContainerEnd) + 8 + 36
						blipEnd := uint32(spContainerEnd) + 8 + bseRL
						if blipEnd <= uint32(dataLen) && blipStart < blipEnd {
							// Parse the blip
							img, _, err := parseBlip(dataStreamBytes, blipStart)
							if err == nil && len(img.Data) > 0 {
								if images != nil {
									idx := len(*images)
									*images = append(*images, img)
									result[int32(picfOffset)] = idx
								}
							}
						}
					}
				}

				// Fallback: if no embedded BSE found, use pib from SpContainer
				if _, ok := result[int32(picfOffset)]; !ok {
					_, pib := parseSpContainerForPib(dataStreamBytes, uint32(i)+8, uint32(i)+8+recLen)
					if pib > 0 {
						result[int32(picfOffset)] = int(pib) - 1
					}
				}
			}

			i += int(7 + recLen)
		}
	}
	return result
}
