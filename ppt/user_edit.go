package ppt

import (
	"encoding/binary"
	"fmt"
)

// parseCurrentUser parses the Current User stream and returns the
// offsetToCurrentEdit value pointing to the latest UserEdit record.
//
// Current User Stream layout:
//   Offset 0:  RecordHeader (8 bytes)
//   Offset 8:  size (4 bytes, fixed 0x14)
//   Offset 12: headerToken (4 bytes)
//   Offset 16: offsetToCurrentEdit (4 bytes, uint32 LE)
func parseCurrentUser(data []byte) (uint32, error) {
	const minSize = 20 // 8 (rh) + 4 (size) + 4 (headerToken) + 4 (offsetToCurrentEdit)
	if len(data) < minSize {
		return 0, fmt.Errorf("current user stream too short: %d bytes, need at least %d", len(data), minSize)
	}

	// Read RecordHeader at offset 0 to validate structure
	_, err := readRecordHeader(data, 0)
	if err != nil {
		return 0, fmt.Errorf("failed to read current user record header: %w", err)
	}

	offsetToCurrentEdit := binary.LittleEndian.Uint32(data[16:20])
	return offsetToCurrentEdit, nil
}

// buildPersistDirectory traverses the UserEdit chain starting from
// offsetToCurrentEdit and builds a complete persistId-to-offset map.
//
// The first UserEdit in the chain (most recent edit) takes precedence
// when there are duplicate persistId entries.
func buildPersistDirectory(pptDocData []byte, offsetToCurrentEdit uint32) (map[uint32]uint32, error) {
	result := make(map[uint32]uint32)
	currentOffset := offsetToCurrentEdit

	for {
		// Read RecordHeader at currentOffset
		rh, err := readRecordHeader(pptDocData, currentOffset)
		if err != nil {
			return nil, fmt.Errorf("failed to read UserEdit record header at offset %d: %w", currentOffset, err)
		}

		// Verify recType == rtUserEdit (0x0FF5)
		if rh.recType != rtUserEdit {
			return nil, fmt.Errorf("expected UserEdit record (0x%04X) at offset %d, got 0x%04X", rtUserEdit, currentOffset, rh.recType)
		}

		// Read offsetLastEdit at currentOffset+16
		if uint32(len(pptDocData)) < currentOffset+24 {
			return nil, fmt.Errorf("not enough data for UserEdit fields at offset %d", currentOffset)
		}
		offsetLastEdit := binary.LittleEndian.Uint32(pptDocData[currentOffset+16 : currentOffset+20])
		offsetPersistDir := binary.LittleEndian.Uint32(pptDocData[currentOffset+20 : currentOffset+24])

		// Parse PersistDirectoryAtom at offsetPersistDir
		if err := parsePersistDirectoryAtom(pptDocData, offsetPersistDir, result); err != nil {
			return nil, fmt.Errorf("failed to parse PersistDirectoryAtom at offset %d: %w", offsetPersistDir, err)
		}

		// If offsetLastEdit == 0, we've reached the end of the chain
		if offsetLastEdit == 0 {
			break
		}
		currentOffset = offsetLastEdit
	}

	return result, nil
}

// parsePersistDirectoryAtom parses a PersistDirectoryAtom at the given offset
// and adds persistId→offset mappings to the result map.
// Only adds entries if the key doesn't already exist (first UserEdit takes precedence).
func parsePersistDirectoryAtom(pptDocData []byte, offset uint32, result map[uint32]uint32) error {
	// Read RecordHeader
	rh, err := readRecordHeader(pptDocData, offset)
	if err != nil {
		return fmt.Errorf("failed to read PersistDirectoryAtom header: %w", err)
	}

	// Verify recType == rtPersistDirectoryAtom (0x1772)
	if rh.recType != rtPersistDirectoryAtom {
		return fmt.Errorf("expected PersistDirectoryAtom (0x%04X), got 0x%04X", rtPersistDirectoryAtom, rh.recType)
	}

	// The atom data starts after the 8-byte header
	dataStart := offset + recordHeaderSize
	dataEnd := dataStart + rh.recLen
	if uint32(len(pptDocData)) < dataEnd {
		return fmt.Errorf("not enough data for PersistDirectoryAtom body (need %d bytes from offset %d)", rh.recLen, dataStart)
	}

	atomData := pptDocData[dataStart:dataEnd]
	pos := uint32(0)

	for pos < uint32(len(atomData)) {
		// Each entry starts with a 4-byte value:
		//   low 20 bits = persistId (starting)
		//   high 12 bits = cPersist (count)
		if pos+4 > uint32(len(atomData)) {
			return fmt.Errorf("not enough data for persist directory entry at position %d", pos)
		}

		entryHeader := binary.LittleEndian.Uint32(atomData[pos : pos+4])
		pos += 4

		persistId := entryHeader & 0x000FFFFF       // low 20 bits
		cPersist := (entryHeader >> 20) & 0x00000FFF // high 12 bits

		// Read cPersist uint32 offsets
		if pos+cPersist*4 > uint32(len(atomData)) {
			return fmt.Errorf("not enough data for %d persist offsets at position %d", cPersist, pos)
		}

		for i := uint32(0); i < cPersist; i++ {
			offsetVal := binary.LittleEndian.Uint32(atomData[pos : pos+4])
			pos += 4

			id := persistId + i
			// Only add if key doesn't already exist (first UserEdit takes precedence)
			if _, exists := result[id]; !exists {
				result[id] = offsetVal
			}
		}
	}

	return nil
}
