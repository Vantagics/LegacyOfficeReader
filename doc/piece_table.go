package doc

import (
	"encoding/binary"
	"errors"
	"fmt"
	"strings"

	"github.com/shakinm/xlsReader/helpers"
)

// pieceDescriptor describes a text fragment in the WordDocument stream.
type pieceDescriptor struct {
	fc  uint32 // file character position (bit 30 marks encoding type)
	prm uint16 // property modifier
}

// piece represents a resolved text fragment with position and encoding info.
type piece struct {
	cpStart   uint32 // character position start
	cpEnd     uint32 // character position end
	fc        uint32 // byte offset in stream (actual, after bit manipulation)
	isUnicode bool   // true=UTF-16LE, false=ANSI
}

// extractText extracts text from the WordDocument stream using the piece table
// found in the table stream at the offset specified by the FIB.
func extractText(wordDocData, tableData []byte, f *fib) (string, error) {
	pieces, err := parsePieceTable(tableData, f)
	if err != nil {
		return "", err
	}

	// Extract text fragments from WordDocument stream
	var sb strings.Builder
	for _, p := range pieces {
		charCount := p.cpEnd - p.cpStart
		if charCount == 0 {
			continue
		}

		var byteCount uint32
		if p.isUnicode {
			byteCount = charCount * 2
		} else {
			byteCount = charCount
		}

		start := p.fc
		end := start + byteCount
		if uint64(end) > uint64(len(wordDocData)) {
			return "", fmt.Errorf("piece text data out of bounds: offset=%d, length=%d, streamSize=%d", start, byteCount, len(wordDocData))
		}

		fragment := wordDocData[start:end]
		if p.isUnicode {
			sb.WriteString(helpers.DecodeUTF16LE(fragment))
		} else {
			sb.WriteString(helpers.DecodeANSI(fragment))
		}
	}

	return sb.String(), nil
}

// parsePieceTable parses the piece table from the Clx structure in the table stream.
// Returns the list of pieces that map CP ranges to FC byte offsets.
func parsePieceTable(tableData []byte, f *fib) ([]piece, error) {
	// Validate Clx bounds
	clxStart := f.fcClx
	clxLen := f.lcbClx
	if uint64(clxStart)+uint64(clxLen) > uint64(len(tableData)) {
		return nil, fmt.Errorf("Clx data out of bounds: offset=%d, length=%d, tableSize=%d", clxStart, clxLen, len(tableData))
	}

	clxData := tableData[clxStart : clxStart+clxLen]

	// Walk through Clx, skip any Prc entries (type 0x01), find PlcPcd (type 0x02)
	pos := uint32(0)
	for pos < uint32(len(clxData)) {
		typeByte := clxData[pos]

		if typeByte == 0x01 {
			// Prc: 1 byte type + 2 bytes size + size bytes data
			if pos+3 > uint32(len(clxData)) {
				return nil, errors.New("Clx: Prc entry truncated")
			}
			prcSize := binary.LittleEndian.Uint16(clxData[pos+1:])
			pos += 3 + uint32(prcSize)
			continue
		}

		if typeByte == 0x02 {
			// PlcPcd found
			pos++ // skip type byte
			break
		}

		return nil, fmt.Errorf("Clx: unexpected type byte 0x%02X at offset %d", typeByte, pos)
	}

	if pos >= uint32(len(clxData)) {
		return nil, errors.New("Clx: PlcPcd not found")
	}

	// Read 4-byte length of PlcPcd data
	if pos+4 > uint32(len(clxData)) {
		return nil, errors.New("Clx: PlcPcd length truncated")
	}
	plcPcdLen := binary.LittleEndian.Uint32(clxData[pos:])
	pos += 4

	if pos+plcPcdLen > uint32(len(clxData)) {
		return nil, fmt.Errorf("Clx: PlcPcd data out of bounds: need %d bytes, have %d", plcPcdLen, uint32(len(clxData))-pos)
	}

	plcPcdData := clxData[pos : pos+plcPcdLen]

	// Calculate n: plcPcdLen = (n+1)*4 + n*8 = 4 + 12*n => n = (plcPcdLen - 4) / 12
	if plcPcdLen < 4 {
		return nil, errors.New("Clx: PlcPcd data too short")
	}
	n := (plcPcdLen - 4) / 12

	if n == 0 {
		return nil, nil // no pieces
	}

	// Read (n+1) CP values (uint32 each)
	cpArraySize := (n + 1) * 4
	if cpArraySize > plcPcdLen {
		return nil, errors.New("Clx: CP array exceeds PlcPcd data")
	}

	cps := make([]uint32, n+1)
	for i := uint32(0); i <= n; i++ {
		cps[i] = binary.LittleEndian.Uint32(plcPcdData[i*4:])
	}

	// Read n PieceDescriptors (8 bytes each), starting after CP array
	pdOffset := cpArraySize
	pieces := make([]piece, n)
	for i := uint32(0); i < n; i++ {
		pdStart := pdOffset + i*8
		if pdStart+8 > plcPcdLen {
			return nil, fmt.Errorf("Clx: PieceDescriptor %d out of bounds", i)
		}

		// PieceDescriptor layout: 2 bytes reserved, 4 bytes fc, 2 bytes prm
		fc := binary.LittleEndian.Uint32(plcPcdData[pdStart+2:])

		var actualOffset uint32
		var isUnicode bool

		if fc&0x40000000 != 0 {
			// Bit 30 set: ANSI encoding
			actualOffset = (fc & ^uint32(0x40000000)) >> 1
			isUnicode = false
		} else {
			// UTF-16LE encoding
			actualOffset = fc
			isUnicode = true
		}

		pieces[i] = piece{
			cpStart:   cps[i],
			cpEnd:     cps[i+1],
			fc:        actualOffset,
			isUnicode: isUnicode,
		}
	}

	return pieces, nil
}

// fcToCP converts a file character (FC) byte offset to a character position (CP)
// using the piece table. Returns the CP and true if found, or 0 and false if not.
func fcToCP(fc uint32, pieces []piece) (uint32, bool) {
	for _, p := range pieces {
		var byteLen uint32
		charCount := p.cpEnd - p.cpStart
		if p.isUnicode {
			byteLen = charCount * 2
		} else {
			byteLen = charCount
		}
		fcEnd := p.fc + byteLen

		if fc >= p.fc && fc < fcEnd {
			// FC is within this piece
			var cpOffset uint32
			if p.isUnicode {
				cpOffset = (fc - p.fc) / 2
			} else {
				cpOffset = fc - p.fc
			}
			return p.cpStart + cpOffset, true
		}
	}
	return 0, false
}

// convertFCRangesToCP converts a slice of FC-based format runs to CP-based ranges
// using the piece table. This is needed because FKP pages store FC byte offsets
// but the text processing works with CP character positions.
func convertFCRangesToCP(fcStart, fcEnd uint32, pieces []piece) (uint32, uint32) {
	cpStart, ok1 := fcToCP(fcStart, pieces)
	cpEnd, ok2 := fcToCP(fcEnd, pieces)
	if !ok1 || !ok2 {
		// Fallback: try to find closest match
		if !ok1 {
			cpStart = 0
		}
		if !ok2 {
			// If fcEnd is past the last piece, use the last CP
			if len(pieces) > 0 {
				cpEnd = pieces[len(pieces)-1].cpEnd
			}
		}
	}
	return cpStart, cpEnd
}

// extractTextFromPieces extracts text from the WordDocument stream using pre-parsed pieces.
func extractTextFromPieces(wordDocData []byte, pieces []piece) (string, error) {
	return extractTextFromPiecesWithCodepage(wordDocData, pieces, 0)
}

// extractTextFromPiecesWithCodepage extracts text using the specified codepage for ANSI pieces.
func extractTextFromPiecesWithCodepage(wordDocData []byte, pieces []piece, codepage uint16) (string, error) {
	var sb strings.Builder
	for _, p := range pieces {
		charCount := p.cpEnd - p.cpStart
		if charCount == 0 {
			continue
		}

		var byteCount uint32
		if p.isUnicode {
			byteCount = charCount * 2
		} else {
			byteCount = charCount
		}

		start := p.fc
		end := start + byteCount
		if uint64(end) > uint64(len(wordDocData)) {
			return "", fmt.Errorf("piece text data out of bounds: offset=%d, length=%d, streamSize=%d", start, byteCount, len(wordDocData))
		}

		fragment := wordDocData[start:end]
		if p.isUnicode {
			sb.WriteString(helpers.DecodeUTF16LE(fragment))
		} else {
			if codepage != 0 {
				sb.WriteString(helpers.DecodeWithCodepage(fragment, codepage))
			} else {
				sb.WriteString(helpers.DecodeANSI(fragment))
			}
		}
	}

	return sb.String(), nil
}
