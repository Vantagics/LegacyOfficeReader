package ppt

import (
	"bytes"
	"encoding/binary"
	"testing"

	"github.com/shakinm/xlsReader/cfb"
	"github.com/shakinm/xlsReader/common"
)

// pptBlipConfig holds the configuration for constructing a valid FBSE+Blip binary record.
type pptBlipConfig struct {
	recType      uint16
	baseInstance uint16
	format       common.ImageFormat
	extraBytes   int
}

var pptBlipConfigs = []pptBlipConfig{
	{rtBlipEMF, 0x3D5, common.ImageFormatEMF, blipMetafileExtra},
	{rtBlipWMF, 0x216, common.ImageFormatWMF, blipMetafileExtra},
	{rtBlipPICT, 0x542, common.ImageFormatPICT, blipMetafileExtra},
	{rtBlipJPEG, 0x46A, common.ImageFormatJPEG, blipBitmapExtra},
	{rtBlipPNG, 0x6E0, common.ImageFormatPNG, blipBitmapExtra},
	{rtBlipDIB, 0x7A8, common.ImageFormatDIB, blipBitmapExtra},
	{rtBlipTIFF, 0x6E4, common.ImageFormatTIFF, blipBitmapExtra},
	{rtBlipJPEG2, 0x6E2, common.ImageFormatJPEG, blipBitmapExtra},
}

// buildFBSERecord constructs a valid FBSE+Blip binary record for testing.
func buildFBSERecord(cfg pptBlipConfig, imageData []byte, hasSecondUID bool) []byte {
	uidCount := 1
	blipRecInstance := cfg.baseInstance
	if hasSecondUID {
		uidCount = 2
		blipRecInstance = cfg.baseInstance + 1
	}

	blipRecLen := uint32(uidCount*blipUIDSize + cfg.extraBytes + len(imageData))
	fbseRecLen := uint32(fbseHeaderSize) + 8 + blipRecLen // FBSE header + Blip header + Blip data

	buf := new(bytes.Buffer)

	// FBSE RecordHeader (8 bytes)
	var fbseHeader [8]byte
	binary.LittleEndian.PutUint16(fbseHeader[0:2], 0) // recInstance=0, recVer=0
	binary.LittleEndian.PutUint16(fbseHeader[2:4], rtFBSE)
	binary.LittleEndian.PutUint32(fbseHeader[4:8], fbseRecLen)
	buf.Write(fbseHeader[:])

	// FBSE fixed header (36 bytes zeros)
	buf.Write(make([]byte, fbseHeaderSize))

	// Embedded Blip RecordHeader (8 bytes)
	var blipHeader [8]byte
	binary.LittleEndian.PutUint16(blipHeader[0:2], blipRecInstance<<4)
	binary.LittleEndian.PutUint16(blipHeader[2:4], cfg.recType)
	binary.LittleEndian.PutUint32(blipHeader[4:8], blipRecLen)
	buf.Write(blipHeader[:])

	// UIDs (zeros)
	buf.Write(make([]byte, uidCount*blipUIDSize))

	// Extra metadata (zeros)
	buf.Write(make([]byte, cfg.extraBytes))

	// Image data
	buf.Write(imageData)

	return buf.Bytes()
}

// buildCorruptFBSERecord constructs a record with a non-FBSE recType so parsePicturesStream
// can read the recLen and skip it.
func buildCorruptFBSERecord(bodyLen int) []byte {
	buf := new(bytes.Buffer)
	var header [8]byte
	binary.LittleEndian.PutUint16(header[0:2], 0)      // recInstance=0, recVer=0
	binary.LittleEndian.PutUint16(header[2:4], 0xFFFF)  // unknown recType (not rtFBSE)
	binary.LittleEndian.PutUint32(header[4:8], uint32(bodyLen))
	buf.Write(header[:])
	buf.Write(make([]byte, bodyLen))
	return buf.Bytes()
}

// --- Unit Tests (Task 5.2) ---
// Requirements: 2.1, 2.2, 2.3, 2.4, 2.5, 6.1, 6.2, 6.3

// TestParsePicturesStream_Empty verifies that nil and empty input return an empty slice.
func TestParsePicturesStream_Empty(t *testing.T) {
	tests := []struct {
		name  string
		input []byte
	}{
		{"nil input", nil},
		{"empty slice", []byte{}},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			result := parsePicturesStream(tc.input)
			if result == nil {
				t.Error("expected non-nil empty slice, got nil")
			}
			if len(result) != 0 {
				t.Errorf("expected 0 images, got %d", len(result))
			}
		})
	}
}

// TestParsePicturesStream_SingleFBSE verifies that a single FBSE+JPEG Blip record
// is correctly parsed and the image data and format are extracted.
func TestParsePicturesStream_SingleFBSE(t *testing.T) {
	imageData := []byte{0xFF, 0xD8, 0xFF, 0xE0}
	cfg := pptBlipConfigs[3] // JPEG
	record := buildFBSERecord(cfg, imageData, false)

	images := parsePicturesStream(record)

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

// TestParsePicturesStream_MultipleFBSE verifies that multiple FBSE records of different
// formats are correctly extracted in order.
func TestParsePicturesStream_MultipleFBSE(t *testing.T) {
	emfData := []byte{0x01, 0x02, 0x03}
	pngData := []byte{0x89, 0x50, 0x4E, 0x47}
	jpegData := []byte{0xFF, 0xD8, 0xFF, 0xE0, 0x00}

	emfRecord := buildFBSERecord(pptBlipConfigs[0], emfData, false)   // EMF
	pngRecord := buildFBSERecord(pptBlipConfigs[4], pngData, false)   // PNG
	jpegRecord := buildFBSERecord(pptBlipConfigs[3], jpegData, false) // JPEG

	stream := make([]byte, 0, len(emfRecord)+len(pngRecord)+len(jpegRecord))
	stream = append(stream, emfRecord...)
	stream = append(stream, pngRecord...)
	stream = append(stream, jpegRecord...)

	images := parsePicturesStream(stream)

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

// TestParsePicturesStream_CorruptFBSE verifies that a corrupt FBSE record (wrong recType)
// is skipped and valid records after it are still extracted.
func TestParsePicturesStream_CorruptFBSE(t *testing.T) {
	corruptRecord := buildCorruptFBSERecord(20) // non-FBSE recType, 20 bytes body
	jpegData := []byte{0xFF, 0xD8, 0xFF, 0xE0}
	jpegRecord := buildFBSERecord(pptBlipConfigs[3], jpegData, false)

	stream := make([]byte, 0, len(corruptRecord)+len(jpegRecord))
	stream = append(stream, corruptRecord...)
	stream = append(stream, jpegRecord...)

	images := parsePicturesStream(stream)

	if len(images) != 1 {
		t.Fatalf("expected 1 image (corrupt FBSE skipped), got %d", len(images))
	}
	if images[0].Format != common.ImageFormatJPEG {
		t.Errorf("expected JPEG format, got %d", images[0].Format)
	}
	if !bytes.Equal(images[0].Data, jpegData) {
		t.Errorf("image data mismatch: got %v, want %v", images[0].Data, jpegData)
	}
}

// TestExtractImagesFromPpt_NoPicturesStream verifies that when picturesDir is nil,
// extractImagesFromPpt returns an empty slice (not nil).
func TestExtractImagesFromPpt_NoPicturesStream(t *testing.T) {
	var adaptor cfb.Cfb

	images := extractImagesFromPpt(adaptor, nil, nil, nil)

	if images == nil {
		t.Error("expected non-nil empty slice, got nil")
	}
	if len(images) != 0 {
		t.Errorf("expected 0 images, got %d", len(images))
	}
}
