package common

import (
	"strings"
	"testing"
	"testing/quick"
)

// Feature: doc-ppt-image-extraction, Property 2: ImageFormat 扩展名映射
// **Validates: Requirements 3.4**
func TestImageFormatExtensionProperty(t *testing.T) {
	allFormats := []ImageFormat{
		ImageFormatEMF,
		ImageFormatWMF,
		ImageFormatPICT,
		ImageFormatJPEG,
		ImageFormatPNG,
		ImageFormatDIB,
		ImageFormatTIFF,
	}

	// Property: every valid ImageFormat returns a non-empty extension starting with "."
	config := &quick.Config{MaxCount: 100}
	prop := func(idx uint8) bool {
		// Map random uint8 into valid format range
		format := allFormats[int(idx)%len(allFormats)]
		img := &Image{Format: format}
		ext := img.Extension()
		if ext == "" {
			t.Logf("Extension() returned empty string for format %d", format)
			return false
		}
		if !strings.HasPrefix(ext, ".") {
			t.Logf("Extension() %q does not start with '.' for format %d", ext, format)
			return false
		}
		return true
	}
	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property failed: all valid ImageFormats should return non-empty extension starting with '.': %v", err)
	}

	// Property: different ImageFormat values return different extensions
	seen := make(map[string]ImageFormat)
	for _, format := range allFormats {
		img := &Image{Format: format}
		ext := img.Extension()
		if prev, exists := seen[ext]; exists {
			t.Errorf("Duplicate extension %q for formats %d and %d", ext, prev, format)
		}
		seen[ext] = format
	}
}

// TestExtension_AllFormats verifies that each ImageFormat returns the correct extension.
// Validates: Requirements 3.1, 3.2, 3.3, 3.4
func TestExtension_AllFormats(t *testing.T) {
	tests := []struct {
		format ImageFormat
		want   string
	}{
		{ImageFormatEMF, ".emf"},
		{ImageFormatWMF, ".wmf"},
		{ImageFormatPICT, ".pict"},
		{ImageFormatJPEG, ".jpeg"},
		{ImageFormatPNG, ".png"},
		{ImageFormatDIB, ".bmp"},
		{ImageFormatTIFF, ".tiff"},
	}

	for _, tc := range tests {
		img := &Image{Format: tc.format}
		got := img.Extension()
		if got != tc.want {
			t.Errorf("Extension() for format %d = %q, want %q", tc.format, got, tc.want)
		}
	}
}

// TestImage_EmptyData verifies that an Image with empty/nil Data behaves correctly.
// Validates: Requirements 3.1, 3.2, 3.3, 3.4
func TestImage_EmptyData(t *testing.T) {
	// nil Data
	img := &Image{Format: ImageFormatPNG, Data: nil}
	if img.Data != nil {
		t.Error("expected nil Data")
	}
	if ext := img.Extension(); ext != ".png" {
		t.Errorf("Extension() with nil Data = %q, want %q", ext, ".png")
	}

	// empty slice Data
	img2 := &Image{Format: ImageFormatJPEG, Data: []byte{}}
	if len(img2.Data) != 0 {
		t.Error("expected empty Data slice")
	}
	if ext := img2.Extension(); ext != ".jpeg" {
		t.Errorf("Extension() with empty Data = %q, want %q", ext, ".jpeg")
	}
}
