package ppt

import (
	"encoding/binary"
	"math/rand"
	"testing"
	"testing/quick"
)

// buildRecordHeader creates an 8-byte PPT record header.
func buildRecordHeader(recVer uint16, recInstance uint16, recType uint16, recLen uint32) []byte {
	header := make([]byte, 8)
	verAndInst := (recInstance << 4) | (recVer & 0x0F)
	binary.LittleEndian.PutUint16(header[0:], verAndInst)
	binary.LittleEndian.PutUint16(header[2:], recType)
	binary.LittleEndian.PutUint32(header[4:], recLen)
	return header
}

// buildClientAnchor16 creates a ClientAnchor record with 16-byte body (4 x int32).
func buildClientAnchor16(top, left, right, bottom int32) []byte {
	body := make([]byte, 16)
	binary.LittleEndian.PutUint32(body[0:], uint32(top))
	binary.LittleEndian.PutUint32(body[4:], uint32(left))
	binary.LittleEndian.PutUint32(body[8:], uint32(right))
	binary.LittleEndian.PutUint32(body[12:], uint32(bottom))
	header := buildRecordHeader(0, 0, rtClientAnchor, 16)
	return append(header, body...)
}

// buildChildAnchor creates a ChildAnchor record with 16-byte body (4 x int32).
func buildChildAnchor(left, top, right, bottom int32) []byte {
	body := make([]byte, 16)
	binary.LittleEndian.PutUint32(body[0:], uint32(left))
	binary.LittleEndian.PutUint32(body[4:], uint32(top))
	binary.LittleEndian.PutUint32(body[8:], uint32(right))
	binary.LittleEndian.PutUint32(body[12:], uint32(bottom))
	header := buildRecordHeader(0, 0, rtChildAnchor, 16)
	return append(header, body...)
}

// buildSpAtom creates a ShapeAtom record with the given shape type in recInstance.
func buildSpAtom(shapeType uint16) []byte {
	body := make([]byte, 8) // spid(4) + grfPersist(4)
	header := buildRecordHeader(2, shapeType, rtSpAtom, 8)
	return append(header, body...)
}

// buildSpContainer wraps sub-records in an SpContainer.
func buildSpContainer(subRecords ...[]byte) []byte {
	var body []byte
	for _, rec := range subRecords {
		body = append(body, rec...)
	}
	header := buildRecordHeader(0xF, 0, rtSpContainer, uint32(len(body)))
	return append(header, body...)
}

// buildSpgrContainer wraps SpContainers in an SpgrContainer.
func buildSpgrContainer(subRecords ...[]byte) []byte {
	var body []byte
	for _, rec := range subRecords {
		body = append(body, rec...)
	}
	header := buildRecordHeader(0xF, 0, rtSpgrContainer, uint32(len(body)))
	return append(header, body...)
}

// buildDrawing wraps content in a Drawing container.
func buildDrawing(content []byte) []byte {
	header := buildRecordHeader(0xF, 0, rtDrawing, uint32(len(content)))
	return append(header, content...)
}

// buildSlideAtom creates a SlideAtom with the given layout type.
func buildSlideAtom(layoutType int32) []byte {
	body := make([]byte, 24) // SlideAtom is typically 24 bytes
	binary.LittleEndian.PutUint32(body[0:], uint32(layoutType))
	header := buildRecordHeader(2, 0, rtSlideAtom, uint32(len(body)))
	return append(header, body...)
}

// buildFSPGR creates an OfficeArtFSPGR record defining the group coordinate system.
func buildFSPGR(xLeft, yTop, xRight, yBottom int32) []byte {
	body := make([]byte, 16)
	binary.LittleEndian.PutUint32(body[0:], uint32(xLeft))
	binary.LittleEndian.PutUint32(body[4:], uint32(yTop))
	binary.LittleEndian.PutUint32(body[8:], uint32(xRight))
	binary.LittleEndian.PutUint32(body[12:], uint32(yBottom))
	header := buildRecordHeader(1, 0, rtFSPGR, 16)
	return append(header, body...)
}

// buildGroupShape creates the group shape (first SpContainer in a group)
// with FSPGR and ClientAnchor.
func buildGroupShape(grpLeft, grpTop, grpRight, grpBottom, ancTop, ancLeft, ancRight, ancBottom int32) []byte {
	fspgr := buildFSPGR(grpLeft, grpTop, grpRight, grpBottom)
	spAtom := buildSpAtom(0) // group shape type = 0
	anchor := buildClientAnchor16(ancTop, ancLeft, ancRight, ancBottom)
	return buildSpContainer(fspgr, spAtom, anchor)
}

// buildSlideContainerData builds a SlideContainer with SlideAtom and Drawing.
func buildSlideContainerData(layoutType int32, shapes ...[]byte) []byte {
	var body []byte
	body = append(body, buildSlideAtom(layoutType)...)

	if len(shapes) > 0 {
		spgr := buildSpgrContainer(shapes...)
		drawing := buildDrawing(spgr)
		body = append(body, drawing...)
	}

	header := buildRecordHeader(0xF, 0, rtSlideContainer, uint32(len(body)))
	return append(header, body...)
}

func TestParseSpContainer_TextBox(t *testing.T) {
	spAtom := buildSpAtom(msosptTextBox)
	anchor := buildClientAnchor16(100, 200, 500, 400)
	sp := buildSpContainer(spAtom, anchor)

	shape := parseSpContainer(sp, 8, uint32(len(sp)), nil)
	if shape == nil {
		t.Fatal("expected non-nil shape")
	}
	if shape.ShapeType != msosptTextBox {
		t.Errorf("expected shape type %d, got %d", msosptTextBox, shape.ShapeType)
	}
	if !shape.IsText {
		t.Error("expected IsText=true for textbox")
	}
}

func TestParseSpContainer_Rectangle(t *testing.T) {
	spAtom := buildSpAtom(msosptRectangle)
	anchor := buildClientAnchor16(0, 0, 100, 100)
	sp := buildSpContainer(spAtom, anchor)

	shape := parseSpContainer(sp, 8, uint32(len(sp)), nil)
	if shape == nil {
		t.Fatal("expected non-nil shape")
	}
	if !shape.IsText {
		t.Error("expected IsText=true for rectangle")
	}
}

func TestParseSpContainer_ClientAnchorPriority(t *testing.T) {
	spAtom := buildSpAtom(msosptTextBox)
	// ChildAnchor with different values
	childAnc := buildChildAnchor(10, 20, 30, 40)
	// ClientAnchor with specific values
	clientAnc := buildClientAnchor16(100, 200, 500, 400)
	sp := buildSpContainer(spAtom, childAnc, clientAnc)

	shape := parseSpContainer(sp, 8, uint32(len(sp)), nil)
	if shape == nil {
		t.Fatal("expected non-nil shape")
	}
	// ClientAnchor should take priority: top=100, left=200, right=500, bottom=400
	expectedLeft := int32(200) * 12700 / 8
	expectedTop := int32(100) * 12700 / 8
	if shape.Left != expectedLeft {
		t.Errorf("expected Left=%d, got %d", expectedLeft, shape.Left)
	}
	if shape.Top != expectedTop {
		t.Errorf("expected Top=%d, got %d", expectedTop, shape.Top)
	}
}

func TestParseSlideContainer_NoDrawing(t *testing.T) {
	data := buildSlideContainerData(0) // no shapes
	shapes, layoutType, _, _ := parseSlideContainer(data, 0, nil)
	if len(shapes) != 0 {
		t.Errorf("expected 0 shapes for slide without Drawing, got %d", len(shapes))
	}
	if layoutType != 0 {
		t.Errorf("expected layout type 0, got %d", layoutType)
	}
}

func TestParseSlideContainer_WithShapes(t *testing.T) {
	// First SpContainer in a group is the group shape itself (skipped in output)
	grpShape := buildGroupShape(0, 0, 7680, 4320, 0, 0, 7680, 4320)
	sp1 := buildSpContainer(buildSpAtom(msosptTextBox), buildClientAnchor16(0, 0, 100, 50))
	sp2 := buildSpContainer(buildSpAtom(msosptRectangle), buildClientAnchor16(50, 50, 200, 150))
	data := buildSlideContainerData(1, grpShape, sp1, sp2)

	shapes, layoutType, _, _ := parseSlideContainer(data, 0, nil)
	if layoutType != 1 {
		t.Errorf("expected layout type 1, got %d", layoutType)
	}
	if len(shapes) != 2 {
		t.Errorf("expected 2 shapes, got %d", len(shapes))
	}
}

// Feature: ppt-to-pptx-format-conversion, Property 4: Shape position/size extraction
// For any valid SpContainer with ClientAnchor or ChildAnchor containing known
// left, top, width, height values, parseSpContainer should return a ShapeFormatting
// with matching position and size. If both anchors exist, ClientAnchor takes priority.
func TestProperty_ShapePositionExtraction(t *testing.T) {
	config := &quick.Config{MaxCount: 100}

	prop := func(seed int64) bool {
		rng := rand.New(rand.NewSource(seed))

		// Generate random anchor values (PPT master units, reasonable range)
		top := int32(rng.Intn(10000))
		left := int32(rng.Intn(10000))
		width := int32(1 + rng.Intn(5000))
		height := int32(1 + rng.Intn(5000))
		right := left + width
		bottom := top + height

		shapeType := uint16(msosptTextBox)
		if rng.Intn(2) == 0 {
			shapeType = uint16(msosptRectangle)
		}

		// Test with ClientAnchor only
		spAtom := buildSpAtom(shapeType)
		anchor := buildClientAnchor16(top, left, right, bottom)
		sp := buildSpContainer(spAtom, anchor)

		shape := parseSpContainer(sp, 8, uint32(len(sp)), nil)
		if shape == nil {
			t.Log("shape is nil")
			return false
		}

		expectedLeft := left * 12700 / 8
		expectedTop := top * 12700 / 8
		expectedWidth := width * 12700 / 8
		expectedHeight := height * 12700 / 8

		if shape.Left != expectedLeft {
			t.Logf("Left mismatch: expected %d, got %d", expectedLeft, shape.Left)
			return false
		}
		if shape.Top != expectedTop {
			t.Logf("Top mismatch: expected %d, got %d", expectedTop, shape.Top)
			return false
		}
		if shape.Width != expectedWidth {
			t.Logf("Width mismatch: expected %d, got %d", expectedWidth, shape.Width)
			return false
		}
		if shape.Height != expectedHeight {
			t.Logf("Height mismatch: expected %d, got %d", expectedHeight, shape.Height)
			return false
		}

		// Test ClientAnchor priority over ChildAnchor
		childAnc := buildChildAnchor(999, 999, 9999, 9999)
		sp2 := buildSpContainer(spAtom, childAnc, anchor)
		shape2 := parseSpContainer(sp2, 8, uint32(len(sp2)), nil)
		if shape2 == nil {
			t.Log("shape2 is nil")
			return false
		}
		if shape2.Left != expectedLeft || shape2.Top != expectedTop {
			t.Log("ClientAnchor should take priority over ChildAnchor")
			return false
		}

		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property failed: shape position extraction: %v", err)
	}
}
