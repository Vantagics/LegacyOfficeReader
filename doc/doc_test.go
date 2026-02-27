package doc

import (
	"os"
	"testing"
)

const testDocPath = "../testfie/test.doc"

func TestOpenFile(t *testing.T) {
	doc, err := OpenFile(testDocPath)
	if err != nil {
		t.Fatalf("OpenFile returned unexpected error: %v", err)
	}
	text := doc.GetText()
	if len(text) == 0 {
		t.Error("expected non-empty text from valid DOC file")
	}
}

func TestOpenReader(t *testing.T) {
	f, err := os.Open(testDocPath)
	if err != nil {
		t.Fatalf("failed to open test file: %v", err)
	}
	defer f.Close()

	doc, err := OpenReader(f)
	if err != nil {
		t.Fatalf("OpenReader returned unexpected error: %v", err)
	}
	text := doc.GetText()
	if len(text) == 0 {
		t.Error("expected non-empty text from valid DOC file")
	}
}

func TestOpenFileAndOpenReaderConsistent(t *testing.T) {
	docFromFile, err := OpenFile(testDocPath)
	if err != nil {
		t.Fatalf("OpenFile returned unexpected error: %v", err)
	}

	f, err := os.Open(testDocPath)
	if err != nil {
		t.Fatalf("failed to open test file: %v", err)
	}
	defer f.Close()

	docFromReader, err := OpenReader(f)
	if err != nil {
		t.Fatalf("OpenReader returned unexpected error: %v", err)
	}

	if docFromFile.GetText() != docFromReader.GetText() {
		t.Error("OpenFile and OpenReader returned different text for the same file")
	}
}

func TestGetText(t *testing.T) {
	doc, err := OpenFile(testDocPath)
	if err != nil {
		t.Fatalf("OpenFile returned unexpected error: %v", err)
	}
	text := doc.GetText()
	if len(text) == 0 {
		t.Error("GetText returned empty string for valid DOC file")
	}
}

func TestInvalidPath(t *testing.T) {
	_, err := OpenFile("nonexistent/path/to/file.doc")
	if err == nil {
		t.Error("expected error for invalid file path, got nil")
	}
}

func TestInvalidFormat(t *testing.T) {
	// Use a temporary file with random bytes to test non-CFB format
	tmpFile, err := os.CreateTemp("", "notadoc-*.doc")
	if err != nil {
		t.Fatalf("failed to create temp file: %v", err)
	}
	defer os.Remove(tmpFile.Name())

	// Write some non-CFB data
	_, err = tmpFile.Write([]byte("this is not a valid CFB or DOC file at all"))
	if err != nil {
		t.Fatalf("failed to write temp file: %v", err)
	}
	tmpFile.Close()

	_, err = OpenFile(tmpFile.Name())
	if err == nil {
		t.Error("expected error for non-CFB format file, got nil")
	}
}
