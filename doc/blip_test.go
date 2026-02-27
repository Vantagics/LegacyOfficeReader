package doc

import (
	"bytes"
	"encoding/binary"
	"testing"
	"testing/quick"

	"github.com/shakinm/xlsReader/cfb"
	"github.com/shakinm/xlsReader/common"
)

// blipConfig holds the configuration for constructing a valid Blip binary record.
type blipConfig struct {
	recType      uint16
	baseInstance uint16
	format       common.ImageFormat
	extraBytes   int
}

var blipConfigs = []blipConfig{
	{rtBlipEMF, 0x3D5, common.ImageFormatEMF, blipMetafileExtra},
	{rtBlipWMF, 0x216, common.ImageFormatWMF, blipMetafileExtra},
	{rtBlipPICT, 0x542, common.ImageFormatPICT, blipMetafileExtra},
	{rtBlipJPEG, 0x46A, common.ImageFormatJPEG, blipBitmapExtra},
	{rtBlipPNG, 0x6E0, common.ImageFormatPNG, blipBitmapExtra},
	{rtBlipDIB, 0x7A8, common.ImageFormatDIB, blipBitmapExtra},
	{rtBlipTIFF, 0x6E4, common.ImageFormatTIFF, blipBitmapExtra},
	{rtBlipJPEG2, 0x6E2, common.ImageFormatJPEG, blipBitmapExtra},
}

// buildBlipRecord constructs a valid Blip binary record from the given parameters.
func buildBlipRecord(cfg blipConfig, imageData []byte, hasSecondUID bool) []byte {
	uidCount := 1
	recInstance := cfg.baseInstance
	if hasSecondUID {
		uidCount = 2
		recInstance = cfg.baseInstance + 1
	}

	recLen := uint32(uidCount*blipUIDSize + cfg.extraBytes + len(imageData))

	buf := new(bytes.Buffer)

	// RecordHeader: first 2 bytes = (recInstance << 4) | recVer (recVer=0)
	var header [8]byte
	binary.LittleEndian.PutUint16(header[0:2], recInstance<<4)
	binary.LittleEndian.PutUint16(header[2:4], cfg.recType)
	binary.LittleEndian.PutUint32(header[4:8], recLen)
	buf.Write(header[:])

	// UID bytes (zeros)
	buf.Write(make([]byte, uidCount*blipUIDSize))

	// Extra metadata bytes (zeros)
	buf.Write(make([]byte, cfg.extraBytes))

	// Image data
	buf.Write(imageData)

	return buf.Bytes()
}

// Feature: doc-ppt-image-extraction, Property 1: Blip 数据提取往返（Round-Trip）
// **Validates: Requirements 1.3, 1.4, 1.5, 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 4.7, 4.8, 4.9**
func TestBlipRoundTrip(t *testing.T) {
	config := &quick.Config{MaxCount: 100}

	prop := func(formatIndex uint8, imageData []byte, hasSecondUID bool) bool {
		// Ensure non-empty image data
		if len(imageData) == 0 {
			imageData = []byte{0xFF}
		}

		cfg := blipConfigs[int(formatIndex)%len(blipConfigs)]

		record := buildBlipRecord(cfg, imageData, hasSecondUID)
		images := parseBlipStream(record)

		if len(images) != 1 {
			t.Logf("expected 1 image, got %d (format=0x%04X, dataLen=%d, secondUID=%v)",
				len(images), cfg.recType, len(imageData), hasSecondUID)
			return false
		}

		if !bytes.Equal(images[0].Data, imageData) {
			t.Logf("image data mismatch: got %d bytes, want %d bytes", len(images[0].Data), len(imageData))
			return false
		}

		if images[0].Format != cfg.format {
			t.Logf("format mismatch: got %d, want %d", images[0].Format, cfg.format)
			return false
		}

		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property 1 (Blip round-trip) failed: %v", err)
	}
}

// buildCorruptRecord constructs a corrupted record with an unknown recType (0xFFFF)
// and a valid 8-byte header so parseBlipStream can read recLen and skip it.
func buildCorruptRecord(bodyLen int) []byte {
	buf := new(bytes.Buffer)
	var header [8]byte
	binary.LittleEndian.PutUint16(header[0:2], 0) // recInstance=0, recVer=0
	binary.LittleEndian.PutUint16(header[2:4], 0xFFFF) // unknown recType
	binary.LittleEndian.PutUint32(header[4:8], uint32(bodyLen))
	buf.Write(header[:])
	buf.Write(make([]byte, bodyLen))
	return buf.Bytes()
}

// Feature: doc-ppt-image-extraction, Property 3: 损坏记录容错
// **Validates: Requirements 6.1, 6.2, 6.3**
func TestFaultTolerance(t *testing.T) {
	config := &quick.Config{MaxCount: 100}

	prop := func(
		nRaw uint8, mRaw uint8,
		formatIndices []uint8, imageDatas [][]byte, secondUIDs []bool,
		corruptLens []uint8,
		interleaveOrder []uint8,
	) bool {
		// N = 1..5 valid records, M = 0..3 corrupted records
		n := int(nRaw%5) + 1
		m := int(mRaw % 4)

		// Build N valid records
		type validEntry struct {
			data   []byte
			imgData []byte
			format common.ImageFormat
		}
		validRecords := make([]validEntry, n)
		for i := 0; i < n; i++ {
			cfgIdx := 0
			if i < len(formatIndices) {
				cfgIdx = int(formatIndices[i]) % len(blipConfigs)
			}
			cfg := blipConfigs[cfgIdx]

			imgData := []byte{byte(i + 1)} // default non-empty
			if i < len(imageDatas) && len(imageDatas[i]) > 0 {
				imgData = imageDatas[i]
			}

			hasSecond := false
			if i < len(secondUIDs) {
				hasSecond = secondUIDs[i]
			}

			validRecords[i] = validEntry{
				data:    buildBlipRecord(cfg, imgData, hasSecond),
				imgData: imgData,
				format:  cfg.format,
			}
		}

		// Build M corrupted records with unknown recType
		corruptRecords := make([][]byte, m)
		for i := 0; i < m; i++ {
			bodyLen := 4 // default
			if i < len(corruptLens) && corruptLens[i] > 0 {
				bodyLen = int(corruptLens[i]%64) + 1
			}
			corruptRecords[i] = buildCorruptRecord(bodyLen)
		}

		// Interleave: place all records in a slice, then shuffle based on interleaveOrder
		type taggedRecord struct {
			data      []byte
			isValid   bool
			validIdx  int
		}
		allRecords := make([]taggedRecord, 0, n+m)
		for i, v := range validRecords {
			allRecords = append(allRecords, taggedRecord{data: v.data, isValid: true, validIdx: i})
		}
		for _, c := range corruptRecords {
			allRecords = append(allRecords, taggedRecord{data: c, isValid: false})
		}

		// Simple deterministic shuffle using interleaveOrder bytes
		total := len(allRecords)
		for i := 0; i < total && i < len(interleaveOrder); i++ {
			j := int(interleaveOrder[i]) % total
			allRecords[i], allRecords[j] = allRecords[j], allRecords[i]
		}

		// Build combined stream and track valid record order
		var stream bytes.Buffer
		var expectedOrder []int
		for _, rec := range allRecords {
			stream.Write(rec.data)
			if rec.isValid {
				expectedOrder = append(expectedOrder, rec.validIdx)
			}
		}

		// Parse
		images := parseBlipStream(stream.Bytes())

		// Verify exactly N images extracted
		if len(images) != n {
			t.Logf("expected %d images, got %d (n=%d, m=%d)", n, len(images), n, m)
			return false
		}

		// Verify each image matches the corresponding valid record in stream order
		for i, validIdx := range expectedOrder {
			if i >= len(images) {
				return false
			}
			if !bytes.Equal(images[i].Data, validRecords[validIdx].imgData) {
				t.Logf("image %d data mismatch: got %d bytes, want %d bytes",
					i, len(images[i].Data), len(validRecords[validIdx].imgData))
				return false
			}
			if images[i].Format != validRecords[validIdx].format {
				t.Logf("image %d format mismatch: got %d, want %d",
					i, images[i].Format, validRecords[validIdx].format)
				return false
			}
		}

		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property 3 (Fault tolerance) failed: %v", err)
	}
}

// --- Unit Tests (Task 2.4) ---
// Requirements: 1.1, 1.2, 1.3, 1.4, 1.5, 4.1-4.10, 6.1, 6.2, 6.3

// TestParseBlipStream_Empty verifies that an empty (or nil) stream returns an empty slice.
func TestParseBlipStream_Empty(t *testing.T) {
	tests := []struct {
		name  string
		input []byte
	}{
		{"nil input", nil},
		{"empty slice", []byte{}},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			result := parseBlipStream(tc.input)
			if len(result) != 0 {
				t.Errorf("expected empty slice, got %d images", len(result))
			}
		})
	}
}

// TestParseBlipStream_SingleJPEG verifies that a single JPEG Blip record is correctly extracted.
func TestParseBlipStream_SingleJPEG(t *testing.T) {
	imageData := []byte{0xFF, 0xD8, 0xFF, 0xE0}
	cfg := blipConfigs[3] // JPEG, recType=0xF01D, baseInstance=0x46A
	record := buildBlipRecord(cfg, imageData, false)

	images := parseBlipStream(record)

	if len(images) != 1 {
		t.Fatalf("expected 1 image, got %d", len(images))
	}
	if images[0].Format != common.ImageFormatJPEG {
		t.Errorf("expected JPEG format, got %d", images[0].Format)
	}
	if !bytes.Equal(images[0].Data, imageData) {
		t.Errorf("image data mismatch: got %v, want %v", images[0].Data, imageData)
	}
}

// TestParseBlipStream_MultipleMixed verifies that multiple Blip records of different formats
// are correctly extracted in order.
func TestParseBlipStream_MultipleMixed(t *testing.T) {
	emfData := []byte{0x01, 0x02, 0x03}
	pngData := []byte{0x89, 0x50, 0x4E, 0x47}
	jpegData := []byte{0xFF, 0xD8, 0xFF, 0xE0, 0x00}

	emfRecord := buildBlipRecord(blipConfigs[0], emfData, false)   // EMF
	pngRecord := buildBlipRecord(blipConfigs[4], pngData, false)   // PNG
	jpegRecord := buildBlipRecord(blipConfigs[3], jpegData, false) // JPEG

	stream := make([]byte, 0, len(emfRecord)+len(pngRecord)+len(jpegRecord))
	stream = append(stream, emfRecord...)
	stream = append(stream, pngRecord...)
	stream = append(stream, jpegRecord...)

	images := parseBlipStream(stream)

	if len(images) != 3 {
		t.Fatalf("expected 3 images, got %d", len(images))
	}

	expected := []struct {
		format common.ImageFormat
		data   []byte
	}{
		{common.ImageFormatEMF, emfData},
		{common.ImageFormatPNG, pngData},
		{common.ImageFormatJPEG, jpegData},
	}

	for i, exp := range expected {
		if images[i].Format != exp.format {
			t.Errorf("image[%d] format: got %d, want %d", i, images[i].Format, exp.format)
		}
		if !bytes.Equal(images[i].Data, exp.data) {
			t.Errorf("image[%d] data mismatch: got %v, want %v", i, images[i].Data, exp.data)
		}
	}
}

// TestParseBlipStream_UnknownRecType verifies that records with unknown recType are skipped
// and valid records after them are still extracted.
func TestParseBlipStream_UnknownRecType(t *testing.T) {
	corruptRecord := buildCorruptRecord(16) // unknown recType 0xFFFF, 16 bytes body
	jpegData := []byte{0xFF, 0xD8, 0xFF, 0xE0}
	jpegRecord := buildBlipRecord(blipConfigs[3], jpegData, false)

	stream := make([]byte, 0, len(corruptRecord)+len(jpegRecord))
	stream = append(stream, corruptRecord...)
	stream = append(stream, jpegRecord...)

	images := parseBlipStream(stream)

	if len(images) != 1 {
		t.Fatalf("expected 1 image (unknown recType skipped), got %d", len(images))
	}
	if images[0].Format != common.ImageFormatJPEG {
		t.Errorf("expected JPEG format, got %d", images[0].Format)
	}
	if !bytes.Equal(images[0].Data, jpegData) {
		t.Errorf("image data mismatch: got %v, want %v", images[0].Data, jpegData)
	}
}

// TestParseBlipStream_TruncatedRecord verifies that a truncated stream (less than 8 bytes
// for a RecordHeader) results in an empty slice.
func TestParseBlipStream_TruncatedRecord(t *testing.T) {
	// Only 4 bytes — not enough for an 8-byte RecordHeader
	truncated := []byte{0x00, 0x00, 0x1D, 0xF0}

	images := parseBlipStream(truncated)

	if len(images) != 0 {
		t.Errorf("expected empty slice for truncated stream, got %d images", len(images))
	}
}

// TestExtractImagesFromDoc_NoStreams verifies that when both dataDir and picturesDir are nil,
// extractImagesFromDoc returns an empty slice.
func TestExtractImagesFromDoc_NoStreams(t *testing.T) {
	var adaptor cfb.Cfb

	images := extractImagesFromDoc(adaptor, nil, nil, nil)

	if images == nil {
		t.Error("expected non-nil empty slice, got nil")
	}
	if len(images) != 0 {
		t.Errorf("expected 0 images, got %d", len(images))
	}
}
