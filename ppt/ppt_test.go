package ppt

import (
	"encoding/binary"
	"fmt"
	"math/rand"
	"os"
	"testing"
	"testing/quick"

	"github.com/shakinm/xlsReader/common"
)

const testPPTPath = "../testfie/test.ppt"

func TestOpenFile(t *testing.T) {
	pres, err := OpenFile(testPPTPath)
	if err != nil {
		t.Fatalf("OpenFile returned unexpected error: %v", err)
	}
	if pres.GetNumberSlides() == 0 {
		t.Error("expected at least one slide from valid PPT file")
	}
}

func TestOpenReader(t *testing.T) {
	f, err := os.Open(testPPTPath)
	if err != nil {
		t.Fatalf("failed to open test file: %v", err)
	}
	defer f.Close()

	pres, err := OpenReader(f)
	if err != nil {
		t.Fatalf("OpenReader returned unexpected error: %v", err)
	}
	if pres.GetNumberSlides() == 0 {
		t.Error("expected at least one slide from valid PPT file")
	}
}

func TestOpenFileAndOpenReaderConsistent(t *testing.T) {
	presFromFile, err := OpenFile(testPPTPath)
	if err != nil {
		t.Fatalf("OpenFile returned unexpected error: %v", err)
	}

	f, err := os.Open(testPPTPath)
	if err != nil {
		t.Fatalf("failed to open test file: %v", err)
	}
	defer f.Close()

	presFromReader, err := OpenReader(f)
	if err != nil {
		t.Fatalf("OpenReader returned unexpected error: %v", err)
	}

	if presFromFile.GetNumberSlides() != presFromReader.GetNumberSlides() {
		t.Errorf("slide count mismatch: OpenFile=%d, OpenReader=%d",
			presFromFile.GetNumberSlides(), presFromReader.GetNumberSlides())
	}

	// Verify each slide has the same texts
	for i := 0; i < presFromFile.GetNumberSlides(); i++ {
		slideFile, _ := presFromFile.GetSlide(i)
		slideReader, _ := presFromReader.GetSlide(i)
		textsFile := slideFile.GetTexts()
		textsReader := slideReader.GetTexts()
		if len(textsFile) != len(textsReader) {
			t.Errorf("slide %d text count mismatch: OpenFile=%d, OpenReader=%d",
				i, len(textsFile), len(textsReader))
			continue
		}
		for j := range textsFile {
			if textsFile[j] != textsReader[j] {
				t.Errorf("slide %d text %d mismatch: OpenFile=%q, OpenReader=%q",
					i, j, textsFile[j], textsReader[j])
			}
		}
	}
}

func TestGetSlides(t *testing.T) {
	pres, err := OpenFile(testPPTPath)
	if err != nil {
		t.Fatalf("OpenFile returned unexpected error: %v", err)
	}
	slides := pres.GetSlides()
	numSlides := pres.GetNumberSlides()
	if len(slides) != numSlides {
		t.Errorf("GetSlides() returned %d slides, but GetNumberSlides() returned %d",
			len(slides), numSlides)
	}
	if len(slides) == 0 {
		t.Error("expected at least one slide from valid PPT file")
	}
}

func TestGetSlide(t *testing.T) {
	pres, err := OpenFile(testPPTPath)
	if err != nil {
		t.Fatalf("OpenFile returned unexpected error: %v", err)
	}
	numSlides := pres.GetNumberSlides()
	if numSlides == 0 {
		t.Fatal("expected at least one slide")
	}

	// Test valid index: first slide
	slide, err := pres.GetSlide(0)
	if err != nil {
		t.Errorf("GetSlide(0) returned unexpected error: %v", err)
	}
	if slide == nil {
		t.Error("GetSlide(0) returned nil slide")
	}

	// Test valid index: last slide
	slide, err = pres.GetSlide(numSlides - 1)
	if err != nil {
		t.Errorf("GetSlide(%d) returned unexpected error: %v", numSlides-1, err)
	}
	if slide == nil {
		t.Errorf("GetSlide(%d) returned nil slide", numSlides-1)
	}

	// Test out-of-range index: negative
	_, err = pres.GetSlide(-1)
	if err == nil {
		t.Error("expected error for GetSlide(-1), got nil")
	}

	// Test out-of-range index: too large
	_, err = pres.GetSlide(numSlides)
	if err == nil {
		t.Errorf("expected error for GetSlide(%d), got nil", numSlides)
	}
}

func TestGetTexts(t *testing.T) {
	pres, err := OpenFile(testPPTPath)
	if err != nil {
		t.Fatalf("OpenFile returned unexpected error: %v", err)
	}
	slides := pres.GetSlides()
	if len(slides) == 0 {
		t.Fatal("expected at least one slide")
	}

	// Verify at least one slide has non-empty text
	hasText := false
	for i, slide := range slides {
		texts := slide.GetTexts()
		for _, text := range texts {
			if len(text) > 0 {
				hasText = true
				t.Logf("Slide %d has text: %q", i, text)
			}
		}
	}
	if !hasText {
		t.Error("expected at least one slide to have non-empty text content")
	}
}

func TestInvalidPath(t *testing.T) {
	_, err := OpenFile("nonexistent/path/to/file.ppt")
	if err == nil {
		t.Error("expected error for invalid file path, got nil")
	}
}

func TestInvalidFormat(t *testing.T) {
	tmpFile, err := os.CreateTemp("", "notappt-*.ppt")
	if err != nil {
		t.Fatalf("failed to create temp file: %v", err)
	}
	defer os.Remove(tmpFile.Name())

	_, err = tmpFile.Write([]byte("this is not a valid CFB or PPT file at all"))
	if err != nil {
		t.Fatalf("failed to write temp file: %v", err)
	}
	tmpFile.Close()

	_, err = OpenFile(tmpFile.Name())
	if err == nil {
		t.Error("expected error for non-CFB format file, got nil")
	}
}

// Feature: doc-ppt-reader, Property 3: Out-of-range slide index returns error
// **Validates: Requirements 6.6**
func TestSlideIndexOutOfRangeProperty(t *testing.T) {
	pres, err := OpenFile(testPPTPath)
	if err != nil {
		t.Skip("test PPT file not available")
	}
	config := &quick.Config{MaxCount: 100}
	f := func(index int) bool {
		if index < 0 || index >= pres.GetNumberSlides() {
			_, err := pres.GetSlide(index)
			return err != nil
		}
		return true // valid indices not in scope for this property
	}
	if err := quick.Check(f, config); err != nil {
		t.Error(err)
	}
}

// Feature: ppt-to-pptx-format-conversion, Property 9: Backward compatibility
// For any Presentation (with random slides and images, with or without formatted shapes),
// GetSlides() should always return the same slides, GetTexts() should return the same texts,
// and GetImages() should return the same images.
func TestProperty_BackwardCompatibility(t *testing.T) {
	config := &quick.Config{MaxCount: 100}

	prop := func(numSlides uint8, numImages uint8, seed int64) bool {
		ns := int(numSlides) % 6 // 0-5 slides
		ni := int(numImages) % 4 // 0-3 images

		slides := make([]Slide, ns)
		for i := 0; i < ns; i++ {
			texts := []string{}
			for j := 0; j <= i; j++ {
				texts = append(texts, fmt.Sprintf("text_%d_%d", i, j))
			}
			slides[i] = Slide{
				texts: texts,
				shapes: []ShapeFormatting{
					{ShapeType: 202, IsText: true, Left: int32(i * 100)},
				},
				layoutType: i,
			}
		}

		images := make([]common.Image, ni)
		for i := 0; i < ni; i++ {
			images[i] = common.Image{
				Format: common.ImageFormatPNG,
				Data:   []byte{byte(i), byte(i + 1)},
			}
		}

		p := Presentation{
			slides:      slides,
			images:      images,
			fonts:       []string{"Arial", "Times"},
			slideWidth:  9144000,
			slideHeight: 6858000,
		}

		// GetSlides returns same slides
		gotSlides := p.GetSlides()
		if len(gotSlides) != ns {
			return false
		}
		for i, s := range gotSlides {
			gotTexts := s.GetTexts()
			if len(gotTexts) != len(slides[i].texts) {
				return false
			}
			for j, txt := range gotTexts {
				if txt != slides[i].texts[j] {
					return false
				}
			}
		}

		// GetImages returns same images
		gotImages := p.GetImages()
		if len(gotImages) != ni {
			return false
		}

		// GetFonts returns same fonts
		gotFonts := p.GetFonts()
		if len(gotFonts) != 2 {
			return false
		}

		// GetSlideSize returns same size
		w, h := p.GetSlideSize()
		if w != 9144000 || h != 6858000 {
			return false
		}

		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property failed: backward compatibility: %v", err)
	}
}

// Feature: ppt-to-pptx-format-conversion, Property 8: Slide size passthrough
// For any valid SlideSize record with known width and height values,
// parseSlideSize should return correctly converted EMU values.
func TestProperty_SlideSizePassthrough(t *testing.T) {
	config := &quick.Config{MaxCount: 100}

	prop := func(seed int64) bool {
		rng := rand.New(rand.NewSource(seed))

		// Generate random PPT master unit values (reasonable range)
		pptWidth := int32(1000 + rng.Intn(10000))
		pptHeight := int32(1000 + rng.Intn(10000))

		// Build a minimal data stream with a SlideSize record
		body := make([]byte, 8)
		binary.LittleEndian.PutUint32(body[0:], uint32(pptWidth))
		binary.LittleEndian.PutUint32(body[4:], uint32(pptHeight))

		header := make([]byte, 8)
		binary.LittleEndian.PutUint16(header[0:], 0x0000)
		binary.LittleEndian.PutUint16(header[2:], uint16(rtSlideSize))
		binary.LittleEndian.PutUint32(header[4:], 8)

		data := append(header, body...)

		w, h := parseSlideSize(data)

		expectedW := pptWidth * 12700 / 8
		expectedH := pptHeight * 12700 / 8

		if w != expectedW {
			t.Logf("width mismatch: expected %d, got %d", expectedW, w)
			return false
		}
		if h != expectedH {
			t.Logf("height mismatch: expected %d, got %d", expectedH, h)
			return false
		}
		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property failed: slide size passthrough: %v", err)
	}
}

func TestParseSlideSize_Default(t *testing.T) {
	// Empty data should return default size
	w, h := parseSlideSize(nil)
	if w != 9144000 || h != 6858000 {
		t.Errorf("expected default size (9144000, 6858000), got (%d, %d)", w, h)
	}
}
