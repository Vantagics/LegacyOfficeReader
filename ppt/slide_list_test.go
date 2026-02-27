package ppt

import (
	"encoding/binary"
	"testing"
)

// buildRecord creates a raw PPT record (header + body).
func buildRecord(recVer uint16, recType uint16, body []byte) []byte {
	buf := make([]byte, 8+len(body))
	// recVerAndInstance: low 4 bits = recVer, high 12 bits = recInstance (0)
	binary.LittleEndian.PutUint16(buf[0:], recVer)
	binary.LittleEndian.PutUint16(buf[2:], recType)
	binary.LittleEndian.PutUint32(buf[4:], uint32(len(body)))
	copy(buf[8:], body)
	return buf
}

// buildUTF16LEBytes encodes a simple ASCII string as UTF-16LE bytes.
func buildUTF16LEBytes(s string) []byte {
	buf := make([]byte, len(s)*2)
	for i, c := range s {
		binary.LittleEndian.PutUint16(buf[2*i:], uint16(c))
	}
	return buf
}

func TestParseSlideListWithText_SingleSlide(t *testing.T) {
	// Build a SlideListWithText container with one slide and one TextBytesAtom
	slidePersist := buildRecord(0x00, rtSlidePersistAtom, make([]byte, 20))
	textBytes := buildRecord(0x00, rtTextBytesAtom, []byte("Hello"))

	containerBody := append(slidePersist, textBytes...)
	container := buildRecord(0x0F, rtSlideListWithText, containerBody)

	slides, err := parseSlideListWithText(container)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if len(slides) != 1 {
		t.Fatalf("expected 1 slide, got %d", len(slides))
	}
	texts := slides[0].GetTexts()
	if len(texts) != 1 || texts[0] != "Hello" {
		t.Errorf("expected [\"Hello\"], got %v", texts)
	}
}

func TestParseSlideListWithText_MultipleSlides(t *testing.T) {
	// Slide 1: TextBytesAtom "Slide1"
	sp1 := buildRecord(0x00, rtSlidePersistAtom, make([]byte, 20))
	tb1 := buildRecord(0x00, rtTextBytesAtom, []byte("Slide1"))

	// Slide 2: TextCharsAtom "Slide2" (UTF-16LE)
	sp2 := buildRecord(0x00, rtSlidePersistAtom, make([]byte, 20))
	tc2 := buildRecord(0x00, rtTextCharsAtom, buildUTF16LEBytes("Slide2"))

	containerBody := append(sp1, tb1...)
	containerBody = append(containerBody, sp2...)
	containerBody = append(containerBody, tc2...)
	container := buildRecord(0x0F, rtSlideListWithText, containerBody)

	slides, err := parseSlideListWithText(container)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if len(slides) != 2 {
		t.Fatalf("expected 2 slides, got %d", len(slides))
	}
	if texts := slides[0].GetTexts(); len(texts) != 1 || texts[0] != "Slide1" {
		t.Errorf("slide 0: expected [\"Slide1\"], got %v", texts)
	}
	if texts := slides[1].GetTexts(); len(texts) != 1 || texts[0] != "Slide2" {
		t.Errorf("slide 1: expected [\"Slide2\"], got %v", texts)
	}
}

func TestParseSlideListWithText_MultipleTextsPerSlide(t *testing.T) {
	sp := buildRecord(0x00, rtSlidePersistAtom, make([]byte, 20))
	tb1 := buildRecord(0x00, rtTextBytesAtom, []byte("Title"))
	tb2 := buildRecord(0x00, rtTextBytesAtom, []byte("Body"))

	containerBody := append(sp, tb1...)
	containerBody = append(containerBody, tb2...)
	container := buildRecord(0x0F, rtSlideListWithText, containerBody)

	slides, err := parseSlideListWithText(container)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if len(slides) != 1 {
		t.Fatalf("expected 1 slide, got %d", len(slides))
	}
	texts := slides[0].GetTexts()
	if len(texts) != 2 || texts[0] != "Title" || texts[1] != "Body" {
		t.Errorf("expected [\"Title\", \"Body\"], got %v", texts)
	}
}

func TestParseSlideListWithText_EmptyStream(t *testing.T) {
	slides, err := parseSlideListWithText([]byte{})
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if len(slides) != 0 {
		t.Errorf("expected 0 slides, got %d", len(slides))
	}
}

func TestParseSlideListWithText_NoSlideListContainer(t *testing.T) {
	// A non-SlideListWithText record should be skipped
	other := buildRecord(0x00, 0x0001, []byte("irrelevant"))

	slides, err := parseSlideListWithText(other)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if len(slides) != 0 {
		t.Errorf("expected 0 slides, got %d", len(slides))
	}
}

func TestParseSlideListWithText_TextBeforeSlidePersist(t *testing.T) {
	// Text atoms before any SlidePersistAtom should be ignored
	tb := buildRecord(0x00, rtTextBytesAtom, []byte("orphan"))
	sp := buildRecord(0x00, rtSlidePersistAtom, make([]byte, 20))
	tb2 := buildRecord(0x00, rtTextBytesAtom, []byte("valid"))

	containerBody := append(tb, sp...)
	containerBody = append(containerBody, tb2...)
	container := buildRecord(0x0F, rtSlideListWithText, containerBody)

	slides, err := parseSlideListWithText(container)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if len(slides) != 1 {
		t.Fatalf("expected 1 slide, got %d", len(slides))
	}
	texts := slides[0].GetTexts()
	if len(texts) != 1 || texts[0] != "valid" {
		t.Errorf("expected [\"valid\"], got %v", texts)
	}
}

func TestParseSlideListWithText_MixedWithOtherRecords(t *testing.T) {
	// Other records before and after the SlideListWithText container
	other1 := buildRecord(0x00, 0x0001, []byte("before"))

	sp := buildRecord(0x00, rtSlidePersistAtom, make([]byte, 20))
	tb := buildRecord(0x00, rtTextBytesAtom, []byte("content"))
	containerBody := append(sp, tb...)
	container := buildRecord(0x0F, rtSlideListWithText, containerBody)

	other2 := buildRecord(0x00, 0x0002, []byte("after"))

	stream := append(other1, container...)
	stream = append(stream, other2...)

	slides, err := parseSlideListWithText(stream)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if len(slides) != 1 {
		t.Fatalf("expected 1 slide, got %d", len(slides))
	}
	if texts := slides[0].GetTexts(); len(texts) != 1 || texts[0] != "content" {
		t.Errorf("expected [\"content\"], got %v", texts)
	}
}
