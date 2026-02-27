package docconv

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/common"
	"github.com/shakinm/xlsReader/doc"
)

// textData is an internal representation of document text for DOCX generation.
type textData struct {
	Text string
}

// imageData is an internal representation of an image for DOCX generation.
type imageData struct {
	Format common.ImageFormat
	Data   []byte
}

// imageRel tracks an image file and its relationship ID inside the DOCX archive.
type imageRel struct {
	filename string
	relID    string
}

// headerFooterEntry is an internal representation of a header/footer entry for DOCX generation.
type headerFooterEntry struct {
	Type    string // "default", "even", "first"
	Text    string
	RawText string
	Images  []int // BSE image indices
}

// formattedData is an internal representation of formatted document content for DOCX generation.
type formattedData struct {
	Paragraphs     []formattedParagraph
	Headers        []string
	Footers        []string
	HeadersRaw     []string
	FootersRaw     []string
	HeaderImages   [][]int // BSE image indices for each header
	FooterImages   [][]int // BSE image indices for each footer
	HeaderEntries  []headerFooterEntry
	FooterEntries  []headerFooterEntry
}

// formattedParagraph is an internal representation of a formatted paragraph.
type formattedParagraph struct {
	Runs            []formattedRun
	Props           doc.ParagraphFormatting
	HeadingLevel    uint8
	IsListItem      bool
	ListType        uint8
	ListLevel       uint8
	ListIlfo        uint16
	ListNfc         uint8
	ListLvlText     string
	InTable         bool
	TableRowEnd     bool
	IsTableCellEnd  bool
	PageBreakBefore bool
	HasPageBreak    bool
	IsTOC           bool
	TOCLevel        uint8
	IsSectionBreak  bool
	SectionType     uint8
	CellWidths      []int32
	DrawnImages     []int  // BSE image indices (0-based) for drawn objects
	TextBoxText     string // text from a text box shape anchored to this paragraph
}

// formattedRun is an internal representation of a formatted text run.
type formattedRun struct {
	Text     string
	Props    doc.CharacterFormatting
	ImageRef int // BSE image index (0-based) for inline image, -1 if not an image
}

// mapFormattedContent extracts formatted content from a parsed Document.
// Returns nil if the document has no formatted content available.
func mapFormattedContent(d *doc.Document) *formattedData {
	fc := d.GetFormattedContent()
	if fc == nil {
		return nil
	}
	fd := &formattedData{
		Paragraphs:   make([]formattedParagraph, len(fc.Paragraphs)),
		Headers:      fc.Headers,
		Footers:      fc.Footers,
		HeadersRaw:   fc.HeadersRaw,
		FootersRaw:   fc.FootersRaw,
		HeaderImages: fc.HeaderImages,
		FooterImages: fc.FooterImages,
	}
	// Map structured header/footer entries
	for _, he := range fc.HeaderEntries {
		fd.HeaderEntries = append(fd.HeaderEntries, headerFooterEntry{
			Type: he.Type, Text: he.Text, RawText: he.RawText, Images: he.Images,
		})
	}
	for _, fe := range fc.FooterEntries {
		fd.FooterEntries = append(fd.FooterEntries, headerFooterEntry{
			Type: fe.Type, Text: fe.Text, RawText: fe.RawText, Images: fe.Images,
		})
	}
	for i, p := range fc.Paragraphs {
		fp := formattedParagraph{
			Props:           p.Props,
			HeadingLevel:    p.HeadingLevel,
			IsListItem:      p.IsListItem,
			ListType:        p.ListType,
			ListLevel:       p.ListLevel,
			ListIlfo:        p.ListIlfo,
			ListNfc:         p.ListNfc,
			ListLvlText:     p.ListLvlText,
			InTable:         p.InTable,
			TableRowEnd:     p.TableRowEnd,
			IsTableCellEnd:  p.IsTableCellEnd,
			PageBreakBefore: p.PageBreakBefore,
			HasPageBreak:    p.HasPageBreak,
			IsTOC:           p.IsTOC,
			TOCLevel:        p.TOCLevel,
			IsSectionBreak:  p.IsSectionBreak,
			SectionType:     p.SectionType,
			CellWidths:      p.CellWidths,
			DrawnImages:     p.DrawnImages,
			TextBoxText:     p.TextBoxText,
			Runs:            make([]formattedRun, len(p.Runs)),
		}
		for j, r := range p.Runs {
			fp.Runs[j] = formattedRun{Text: r.Text, Props: r.Props, ImageRef: r.ImageRef}
		}
		fd.Paragraphs[i] = fp
	}
	return fd
}

// mapText extracts text data from a parsed Document.
func mapText(d *doc.Document) textData {
	return textData{Text: d.GetText()}
}

// mapImages extracts image data from a parsed Document.
func mapImages(d *doc.Document) []imageData {
	docImages := d.GetImages()
	result := make([]imageData, len(docImages))
	for i, img := range docImages {
		data := make([]byte, len(img.Data))
		copy(data, img.Data)
		result[i] = imageData{Format: img.Format, Data: data}
	}
	return result
}

// ConvertReader reads DOC data from reader, converts it to DOCX, and writes to writer.
func ConvertReader(reader io.ReadSeeker, writer io.Writer) error {
	document, err := doc.OpenReader(reader)
	if err != nil {
		return fmt.Errorf("docconv: failed to parse input: %w", err)
	}

	images := mapImages(&document)

	// Try formatted content first
	fd := mapFormattedContent(&document)
	if fd != nil {
		if err := writeDocxFormatted(writer, fd, images); err != nil {
			return fmt.Errorf("docconv: failed to write docx: %w", err)
		}
		return nil
	}

	// Fallback to plain text
	text := mapText(&document)
	if err := writeDocx(writer, text, images); err != nil {
		return fmt.Errorf("docconv: failed to write docx: %w", err)
	}
	return nil
}

// ConvertFile converts a DOC file at inputPath to a DOCX file at outputPath.
func ConvertFile(inputPath string, outputPath string) error {
	inFile, err := os.Open(inputPath)
	if err != nil {
		return fmt.Errorf("docconv: failed to open input file: %w", err)
	}
	defer inFile.Close()

	outFile, err := os.Create(outputPath)
	if err != nil {
		return fmt.Errorf("docconv: failed to create output file: %w", err)
	}
	defer outFile.Close()

	return ConvertReader(inFile, outFile)
}

// writeDocx generates a minimal valid DOCX (Office Open XML) zip archive.
func writeDocx(w io.Writer, text textData, images []imageData) error {
	zw := zip.NewWriter(w)
	defer zw.Close()

	// Write images into word/media/ and collect relationship info
	var imgRels []imageRel
	for i, img := range images {
		ext := (&common.Image{Format: img.Format}).Extension()
		if ext == "" {
			ext = ".bin"
		}
		filename := fmt.Sprintf("image%d%s", i+1, ext)
		fw, err := zw.Create("word/media/" + filename)
		if err != nil {
			return err
		}
		if _, err := fw.Write(img.Data); err != nil {
			return err
		}
		imgRels = append(imgRels, imageRel{
			filename: filename,
			relID:    fmt.Sprintf("rImg%d", i+1),
		})
	}

	// Write word/document.xml
	fw, err := zw.Create("word/document.xml")
	if err != nil {
		return err
	}
	if err := writeDocumentXML(fw, text, imgRels); err != nil {
		return err
	}

	// Write word/_rels/document.xml.rels
	fw, err = zw.Create("word/_rels/document.xml.rels")
	if err != nil {
		return err
	}
	if err := writeDocumentRels(fw, imgRels, false); err != nil {
		return err
	}

	// Write word/styles.xml
	fw, err = zw.Create("word/styles.xml")
	if err != nil {
		return err
	}
	if err := writeStylesXML(fw); err != nil {
		return err
	}

	// Write [Content_Types].xml
	fw, err = zw.Create("[Content_Types].xml")
	if err != nil {
		return err
	}
	if err := writeContentTypes(fw, imgRels, false, 0, 0); err != nil {
		return err
	}

	// Write _rels/.rels
	fw, err = zw.Create("_rels/.rels")
	if err != nil {
		return err
	}
	if _, err := io.WriteString(fw, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`+
		`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>`+
		`</Relationships>`); err != nil {
		return err
	}

	return nil
}

// cleanText strips DOC field codes and special characters from raw text.
// Field codes use 0x13 (begin), 0x14 (separator), 0x15 (end). Text between
// 0x13 and 0x14 is the field instruction and is removed. Text between 0x14
// and 0x15 is the field result (visible text) and is kept. If no 0x14 appears
// before 0x15, the entire field is removed. Characters 0x07 (cell/row mark)
// and 0x08 (drawn object) are also stripped. 0x01 (inline image placeholder)
// and 0x0D (\r paragraph separator) are preserved.
func cleanText(raw string) string {
	var result []rune
	depth := 0
	inInstruction := false

	for _, r := range raw {
		switch r {
		case 0x13: // field begin
			depth++
			inInstruction = true
		case 0x14: // field separator — switch from instruction to result
			inInstruction = false
		case 0x15: // field end
			depth--
			if depth <= 0 {
				depth = 0
				inInstruction = false
			}
		case 0x02, 0x03, 0x04, 0x05: // footnote ref, separator, endnote sep, annotation ref — skip
			continue
		case 0x07, 0x08: // cell mark, drawn object — skip
			continue
		case 0x0C: // page/section break — skip
			continue
		default:
			if depth > 0 && inInstruction {
				continue // skip field instruction text
			}
			result = append(result, r)
		}
	}
	return string(result)
}

// splitParagraphs splits text on \r (DOC paragraph separator).
func splitParagraphs(text string) []string {
	text = strings.ReplaceAll(text, "\r\n", "\r")
	return strings.Split(text, "\r")
}

// writeImageDrawing writes the OOXML drawing markup for an inline image.
func writeImageDrawing(buf *bytes.Buffer, rel imageRel, docPrID int, img *imageData) {
	// Try to get actual image dimensions in EMU; default to 4 inches (3657600 EMU)
	// Text area width: page width (11906 twips) - left margin (1800) - right margin (1800) = 8306 twips
	// 8306 twips * 635 EMU/twip = 5274310 EMU
	maxWidth := int64(5274310)
	cx, cy := int64(3657600), int64(3657600)
	if img != nil {
		emuW, emuH := getImageDimensionsEMU(img)
		if emuW > 0 && emuH > 0 {
			cx, cy = emuW, emuH
			// Scale to fit text area width if wider (preserve aspect ratio)
			if cx > maxWidth {
				scale := float64(maxWidth) / float64(cx)
				cx = maxWidth
				cy = int64(float64(cy) * scale)
			}
		}
	}

	buf.WriteString(`<w:drawing>`)
	buf.WriteString(`<wp:inline distT="0" distB="0" distL="0" distR="0">`)
	fmt.Fprintf(buf, `<wp:extent cx="%d" cy="%d"/>`, cx, cy)
	fmt.Fprintf(buf, `<wp:docPr id="%d" name="Image %d"/>`, docPrID, docPrID)
	buf.WriteString(`<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">`)
	buf.WriteString(`<pic:pic><pic:nvPicPr>`)
	fmt.Fprintf(buf, `<pic:cNvPr id="%d" name="Image %d"/>`, docPrID, docPrID)
	buf.WriteString(`<pic:cNvPicPr/></pic:nvPicPr>`)
	buf.WriteString(`<pic:blipFill>`)
	fmt.Fprintf(buf, `<a:blip r:embed="%s"/>`, rel.relID)
	buf.WriteString(`<a:stretch><a:fillRect/></a:stretch>`)
	buf.WriteString(`</pic:blipFill>`)
	fmt.Fprintf(buf, `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>`, cx, cy)
	buf.WriteString(`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`)
	buf.WriteString(`</pic:pic></a:graphicData></a:graphic>`)
	buf.WriteString(`</wp:inline>`)
	buf.WriteString(`</w:drawing>`)
}
// writeDrawnImageDrawing writes a drawn object image as an anchored (floating) drawing.
// Drawn objects in DOC (from PlcSpaMom/0x08 characters) are floating shapes, not inline.
// They are rendered with wrapTopAndBottom so they sit between text without overlapping.
func writeDrawnImageDrawing(buf *bytes.Buffer, rel imageRel, docPrID int, img *imageData) {
	// Use inline mode (same as writeImageDrawing) to keep images in the
	// document flow. Anchor mode can cause layout issues where large images
	// push subsequent text to a near-empty next page.
	// Text area width: page width (11906 twips) - left margin (1800) - right margin (1800) = 8306 twips
	// 8306 twips * 635 EMU/twip = 5274310 EMU
	maxWidth := int64(5274310)

	cx, cy := int64(3657600), int64(3657600)
	if img != nil {
		emuW, emuH := getImageDimensionsEMU(img)
		if emuW > 0 && emuH > 0 {
			cx, cy = emuW, emuH
			// Scale to fit text area width if wider
			if cx > maxWidth {
				scale := float64(maxWidth) / float64(cx)
				cx = maxWidth
				cy = int64(float64(cy) * scale)
			}
		}
	}

	buf.WriteString(`<w:drawing>`)
	buf.WriteString(`<wp:inline distT="0" distB="0" distL="0" distR="0">`)
	fmt.Fprintf(buf, `<wp:extent cx="%d" cy="%d"/>`, cx, cy)
	fmt.Fprintf(buf, `<wp:docPr id="%d" name="Image %d"/>`, docPrID, docPrID)
	buf.WriteString(`<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">`)
	buf.WriteString(`<pic:pic><pic:nvPicPr>`)
	fmt.Fprintf(buf, `<pic:cNvPr id="%d" name="Image %d"/>`, docPrID, docPrID)
	buf.WriteString(`<pic:cNvPicPr/></pic:nvPicPr>`)
	buf.WriteString(`<pic:blipFill>`)
	fmt.Fprintf(buf, `<a:blip r:embed="%s"/>`, rel.relID)
	buf.WriteString(`<a:stretch><a:fillRect/></a:stretch>`)
	buf.WriteString(`</pic:blipFill>`)
	fmt.Fprintf(buf, `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>`, cx, cy)
	buf.WriteString(`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`)
	buf.WriteString(`</pic:pic></a:graphicData></a:graphic>`)
	buf.WriteString(`</wp:inline>`)
	buf.WriteString(`</w:drawing>`)
}


// writeTextBoxDrawing writes a DOCX text box shape as an anchored drawing.
// The text box is rendered as a floating shape with the given text content,
// styled as a large centered title (matching the DOC title page text box).
func writeTextBoxDrawing(buf *bytes.Buffer, text string) {
	// Text box dimensions: roughly page-width centered, ~1 inch tall
	// cx/cy in EMU: page width ~6in = 5486400, height ~914400 (1in)
	cx := int64(5486400)
	cy := int64(914400)

	buf.WriteString(`<w:r>`)
	buf.WriteString(`<w:drawing>`)
	buf.WriteString(`<wp:anchor distT="0" distB="0" distL="114300" distR="114300"`)
	buf.WriteString(` simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0"`)
	buf.WriteString(` layoutInCell="1" allowOverlap="1">`)
	buf.WriteString(`<wp:simplePos x="0" y="0"/>`)
	buf.WriteString(`<wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>`)
	buf.WriteString(`<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>`)
	fmt.Fprintf(buf, `<wp:extent cx="%d" cy="%d"/>`, cx, cy)
	buf.WriteString(`<wp:effectExtent l="0" t="0" r="0" b="0"/>`)
	buf.WriteString(`<wp:wrapNone/>`)
	buf.WriteString(`<wp:docPr id="100" name="TextBox 1"/>`)
	buf.WriteString(`<a:graphic>`)
	buf.WriteString(`<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">`)
	buf.WriteString(`<wps:wsp>`)
	buf.WriteString(`<wps:cNvSpPr txBox="1"/>`)
	buf.WriteString(`<wps:spPr>`)
	fmt.Fprintf(buf, `<a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>`, cx, cy)
	buf.WriteString(`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`)
	buf.WriteString(`<a:noFill/>`)
	buf.WriteString(`<a:ln><a:noFill/></a:ln>`)
	buf.WriteString(`</wps:spPr>`)
	buf.WriteString(`<wps:txbx>`)
	buf.WriteString(`<w:txbxContent>`)
	buf.WriteString(`<w:p>`)
	buf.WriteString(`<w:pPr><w:jc w:val="center"/></w:pPr>`)
	buf.WriteString(`<w:r>`)
	buf.WriteString(`<w:rPr><w:b/><w:bCs/>`)
	buf.WriteString(`<w:sz w:val="44"/><w:szCs w:val="44"/>`)
	buf.WriteString(`<w:rFonts w:eastAsia="黑体"/>`)
	buf.WriteString(`</w:rPr>`)
	buf.WriteString(`<w:t xml:space="preserve">`)
	xml.Escape(buf, []byte(text))
	buf.WriteString(`</w:t>`)
	buf.WriteString(`</w:r>`)
	buf.WriteString(`</w:p>`)
	buf.WriteString(`</w:txbxContent>`)
	buf.WriteString(`</wps:txbx>`)
	buf.WriteString(`<wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="overflow"`)
	buf.WriteString(` horzOverflow="overflow" vert="horz" wrap="square"`)
	buf.WriteString(` lIns="91440" tIns="45720" rIns="91440" bIns="45720"`)
	buf.WriteString(` numCol="1" anchor="t" anchorCtr="0">`)
	buf.WriteString(`<a:noAutofit/>`)
	buf.WriteString(`</wps:bodyPr>`)
	buf.WriteString(`</wps:wsp>`)
	buf.WriteString(`</a:graphicData>`)
	buf.WriteString(`</a:graphic>`)
	buf.WriteString(`</wp:anchor>`)
	buf.WriteString(`</w:drawing>`)
	buf.WriteString(`</w:r>`)
}

// getImageDimensionsEMU returns width and height in EMU (English Metric Units) for supported image formats.
// 1 inch = 914400 EMU. Returns (0, 0) if dimensions cannot be determined.
func getImageDimensionsEMU(img *imageData) (int64, int64) {
	data := img.Data
	switch img.Format {
	case common.ImageFormatPNG:
		// PNG: IHDR chunk at offset 16-23 (after 8-byte signature + 8-byte chunk header)
		if len(data) > 24 && data[0] == 0x89 && data[1] == 'P' && data[2] == 'N' && data[3] == 'G' {
			w := int64(data[16])<<24 | int64(data[17])<<16 | int64(data[18])<<8 | int64(data[19])
			h := int64(data[20])<<24 | int64(data[21])<<16 | int64(data[22])<<8 | int64(data[23])
			if w > 0 && w < 20000 && h > 0 && h < 20000 {
				// Check for pHYs chunk to get DPI
				dpi := int64(96) // default
				for i := 8; i+16 < len(data); i++ {
					if data[i] == 'p' && data[i+1] == 'H' && data[i+2] == 'Y' && data[i+3] == 's' && i+16 <= len(data) {
						pxPerUnitX := int64(data[i+4])<<24 | int64(data[i+5])<<16 | int64(data[i+6])<<8 | int64(data[i+7])
						unit := data[i+12]
						if unit == 1 && pxPerUnitX > 0 { // meters
							dpi = pxPerUnitX * 254 / 10000 // convert px/m to DPI (approx)
							if dpi < 72 {
								dpi = 72
							}
						}
						break
					}
				}
				return w * 914400 / dpi, h * 914400 / dpi
			}
		}
	case common.ImageFormatJPEG:
		// JPEG: scan for SOF0 (0xFFC0) or SOF2 (0xFFC2) marker
		for i := 0; i+9 < len(data); i++ {
			if data[i] == 0xFF && (data[i+1] == 0xC0 || data[i+1] == 0xC2) {
				h := int64(data[i+5])<<8 | int64(data[i+6])
				w := int64(data[i+7])<<8 | int64(data[i+8])
				if w > 0 && w < 20000 && h > 0 && h < 20000 {
					return w * 914400 / 96, h * 914400 / 96
				}
			}
		}
	case common.ImageFormatEMF:
		// EMF header: RecordType(4) + RecordSize(4) + Bounds(16) + Frame(16)
		// Frame rectangle at offset 24-39 is in 0.01mm units
		if len(data) >= 40 {
			left := int32(data[24]) | int32(data[25])<<8 | int32(data[26])<<16 | int32(data[27])<<24
			top := int32(data[28]) | int32(data[29])<<8 | int32(data[30])<<16 | int32(data[31])<<24
			right := int32(data[32]) | int32(data[33])<<8 | int32(data[34])<<16 | int32(data[35])<<24
			bottom := int32(data[36]) | int32(data[37])<<8 | int32(data[38])<<16 | int32(data[39])<<24
			// Convert 0.01mm to EMU: 1mm = 36000 EMU, so 0.01mm = 360 EMU
			w := int64(right-left) * 360
			h := int64(bottom-top) * 360
			if w > 0 && h > 0 {
				return w, h
			}
		}
	}
	return 0, 0
}

func writeDocumentXML(w io.Writer, text textData, imgRels []imageRel) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`)
	buf.WriteString(` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`)
	buf.WriteString(` xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"`)
	buf.WriteString(` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`)
	buf.WriteString(` xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">`)
	buf.WriteString(`<w:body>`)

	cleaned := cleanText(text.Text)
	// If no images to place, strip inline image placeholders (0x01)
	if len(imgRels) == 0 {
		cleaned = strings.ReplaceAll(cleaned, "\x01", "")
	}
	paragraphs := splitParagraphs(cleaned)
	imgIdx := 0

	for _, para := range paragraphs {
		buf.WriteString(`<w:p>`)
		// Split on 0x01 to find inline image insertion points
		segments := strings.Split(para, "\x01")
		for si, seg := range segments {
			if si > 0 && imgIdx < len(imgRels) {
				// Insert inline image at this position
				buf.WriteString(`<w:r>`)
				writeImageDrawing(&buf, imgRels[imgIdx], imgIdx+1, nil)
				buf.WriteString(`</w:r>`)
				imgIdx++
			}
			if seg != "" {
				buf.WriteString(`<w:r><w:t xml:space="preserve">`)
				xml.Escape(&buf, []byte(seg))
				buf.WriteString(`</w:t></w:r>`)
			}
		}
		buf.WriteString(`</w:p>`)
	}

	// Any remaining images that weren't placed inline
	for imgIdx < len(imgRels) {
		buf.WriteString(`<w:p><w:r>`)
		writeImageDrawing(&buf, imgRels[imgIdx], imgIdx+1, nil)
		buf.WriteString(`</w:r></w:p>`)
		imgIdx++
	}

	buf.WriteString(`</w:body></w:document>`)
	_, err := w.Write(buf.Bytes())
	return err
}

func writeDocumentRels(w io.Writer, imgRels []imageRel, hasNumbering bool) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)
	buf.WriteString(`<Relationship Id="rIdStyles1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`)

	if hasNumbering {
		buf.WriteString(`<Relationship Id="rIdNumbering1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`)
	}

	for _, rel := range imgRels {
		buf.WriteString(fmt.Sprintf(`<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/%s"/>`, rel.relID, rel.filename))
	}

	buf.WriteString(`</Relationships>`)
	_, err := w.Write(buf.Bytes())
	return err
}

func writeStylesXML(w io.Writer) error {
	_, err := io.WriteString(w, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
		`<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`+
		`<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>`+
		`</w:styles>`)
	return err
}

// writeFormattedStylesXML generates word/styles.xml with heading style definitions
// for each heading level used in the formatted data.
func writeFormattedStylesXML(w io.Writer, fd *formattedData) error {
	// Scan paragraphs to find used heading levels
	usedLevels := make(map[uint8]bool)
	for _, p := range fd.Paragraphs {
		if p.HeadingLevel >= 1 && p.HeadingLevel <= 9 {
			usedLevels[p.HeadingLevel] = true
		}
	}

	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`)

	// Document defaults: 宋体 10.5pt (21 half-points)
	buf.WriteString(`<w:docDefaults>`)
	buf.WriteString(`<w:rPrDefault><w:rPr>`)
	buf.WriteString(`<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="宋体" w:cs="Times New Roman"/>`)
	buf.WriteString(`<w:sz w:val="21"/><w:szCs w:val="22"/>`)
	buf.WriteString(`<w:lang w:val="en-US" w:eastAsia="zh-CN"/>`)
	buf.WriteString(`</w:rPr></w:rPrDefault>`)
	buf.WriteString(`<w:pPrDefault><w:pPr>`)
	buf.WriteString(`<w:widowControl w:val="0"/>`)
	buf.WriteString(`<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>`)
	buf.WriteString(`</w:pPr></w:pPrDefault>`)
	buf.WriteString(`</w:docDefaults>`)

	// Normal style - justify alignment for Chinese documents, no extra spacing
	buf.WriteString(`<w:style w:type="paragraph" w:default="1" w:styleId="Normal">`)
	buf.WriteString(`<w:name w:val="Normal"/>`)
	buf.WriteString(`<w:qFormat/>`)
	buf.WriteString(`<w:pPr><w:jc w:val="both"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>`)
	buf.WriteString(`</w:style>`)

	// Header style
	buf.WriteString(`<w:style w:type="paragraph" w:styleId="Header">`)
	buf.WriteString(`<w:name w:val="header"/>`)
	buf.WriteString(`<w:basedOn w:val="Normal"/>`)
	buf.WriteString(`<w:pPr><w:tabs><w:tab w:val="center" w:pos="4153"/><w:tab w:val="right" w:pos="8306"/></w:tabs></w:pPr>`)
	buf.WriteString(`<w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>`)
	buf.WriteString(`</w:style>`)

	// Footer style
	buf.WriteString(`<w:style w:type="paragraph" w:styleId="Footer">`)
	buf.WriteString(`<w:name w:val="footer"/>`)
	buf.WriteString(`<w:basedOn w:val="Normal"/>`)
	buf.WriteString(`<w:pPr><w:tabs><w:tab w:val="center" w:pos="4153"/><w:tab w:val="right" w:pos="8306"/></w:tabs><w:jc w:val="left"/></w:pPr>`)
	buf.WriteString(`<w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>`)
	buf.WriteString(`</w:style>`)

	// TOC heading style
	buf.WriteString(`<w:style w:type="paragraph" w:styleId="TOCHeading">`)
	buf.WriteString(`<w:name w:val="TOC Heading"/>`)
	buf.WriteString(`<w:basedOn w:val="Heading1"/>`)
	buf.WriteString(`<w:qFormat/>`)
	buf.WriteString(`</w:style>`)

	// TOC entry styles
	// TOC entry styles with right-aligned tab stop and dot leader for page numbers
	for level := 1; level <= 3; level++ {
		indent := (level - 1) * 220
		fmt.Fprintf(&buf, `<w:style w:type="paragraph" w:styleId="TOC%d">`, level)
		fmt.Fprintf(&buf, `<w:name w:val="toc %d"/>`, level)
		buf.WriteString(`<w:basedOn w:val="Normal"/>`)
		buf.WriteString(`<w:pPr>`)
		buf.WriteString(`<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8306"/></w:tabs>`)
		if indent > 0 {
			fmt.Fprintf(&buf, `<w:ind w:left="%d"/>`, indent)
		}
		buf.WriteString(`</w:pPr>`)
		buf.WriteString(`</w:style>`)
	}

	// Heading style definitions with proper formatting
	// Sizes match the DOC style definitions (标题 1-9)
	headingSizes := map[uint8]int{1: 44, 2: 32, 3: 30, 4: 28, 5: 24, 6: 24, 7: 22, 8: 21, 9: 21}
	for level := uint8(1); level <= 9; level++ {
		if !usedLevels[level] {
			continue
		}
		sz := headingSizes[level]
		fmt.Fprintf(&buf, `<w:style w:type="paragraph" w:styleId="Heading%d">`, level)
		fmt.Fprintf(&buf, `<w:name w:val="heading %d"/>`, level)
		buf.WriteString(`<w:basedOn w:val="Normal"/>`)
		buf.WriteString(`<w:qFormat/>`)
		fmt.Fprintf(&buf, `<w:pPr><w:jc w:val="left"/><w:outlineLvl w:val="%d"/>`, level-1)
		buf.WriteString(`<w:keepNext/><w:keepLines/>`)
		buf.WriteString(`</w:pPr>`)
		fmt.Fprintf(&buf, `<w:rPr><w:b/><w:bCs/><w:sz w:val="%d"/><w:szCs w:val="%d"/>`, sz, sz)
		buf.WriteString(`<w:rFonts w:eastAsia="宋体"/>`)
		buf.WriteString(`</w:rPr>`)
		buf.WriteString(`</w:style>`)
	}

	buf.WriteString(`</w:styles>`)
	_, err := w.Write(buf.Bytes())
	return err
}

// writeFormattedDocumentXML generates document.xml with formatting information.
func writeFormattedDocumentXML(w io.Writer, fd *formattedData, imgRels []imageRel, images []imageData) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`)
	buf.WriteString(` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`)
	buf.WriteString(` xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"`)
	buf.WriteString(` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`)
	buf.WriteString(` xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`)
	buf.WriteString(` xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"`)
	buf.WriteString(` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`)
	buf.WriteString(` xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"`)
	buf.WriteString(` xmlns:v="urn:schemas-microsoft-com:vml"`)
	buf.WriteString(` mc:Ignorable="wp14">`)
	buf.WriteString(`<w:body>`)

	listInfo := buildListNumInfo(fd)
	imgIdx := 0

	// Build heading bookmark map for TOC hyperlinks.
	// Each heading gets a bookmark like "_Toc1", "_Toc2", etc.
	// TOC entries are matched to headings by text content.
	type headingBookmark struct {
		bookmarkName string
		bookmarkID   int
	}
	headingBookmarks := make(map[int]*headingBookmark) // paragraph index -> bookmark
	tocTargets := make(map[int]string)                 // TOC paragraph index -> bookmark name
	nextBookmarkID := 1

	// Collect headings
	type headingInfo struct {
		paraIdx      int
		text         string
		bookmarkName string
	}
	var headings []headingInfo
	for idx, p := range fd.Paragraphs {
		if p.HeadingLevel > 0 && !p.IsTOC {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			text = strings.TrimSpace(text)
			bm := fmt.Sprintf("_Toc%d", nextBookmarkID)
			headingBookmarks[idx] = &headingBookmark{bookmarkName: bm, bookmarkID: nextBookmarkID}
			headings = append(headings, headingInfo{paraIdx: idx, text: text, bookmarkName: bm})
			nextBookmarkID++
		}
	}

	// Match TOC entries to headings by text similarity
	usedHeadings := make(map[int]bool) // track which headings have been matched
	for idx, p := range fd.Paragraphs {
		if p.IsTOC {
			tocText := ""
			for _, r := range p.Runs {
				tocText += r.Text
			}
			tocText = strings.TrimSpace(tocText)
			if tocText == "" {
				continue
			}
			// Find best matching heading: prefer exact match, then shortest containing match
			bestMatch := ""
			bestDiff := int(^uint(0) >> 1) // max int
			for hi, h := range headings {
				if usedHeadings[hi] {
					continue
				}
				if h.text == tocText {
					bestMatch = h.bookmarkName
					bestDiff = 0
					usedHeadings[hi] = true
					break
				}
				diff := len(h.text) - len(tocText)
				if diff < 0 {
					diff = -diff
				}
				if (strings.Contains(h.text, tocText) || strings.Contains(tocText, h.text)) && diff < bestDiff {
					bestMatch = h.bookmarkName
					bestDiff = diff
				}
			}
			if bestMatch != "" {
				tocTargets[idx] = bestMatch
			} else {
				tocTargets[idx] = fmt.Sprintf("_Toc%d", nextBookmarkID)
				nextBookmarkID++
			}
		}
	}

	// Note: drawn image frequency filtering has been replaced by a simpler
	// approach: when a paragraph has multiple drawn images, only the largest
	// one is rendered (see writeParagraphXML).

	// Pre-filter: build a set of paragraph indices to skip.
	// Pre-filter: build a set of paragraph indices to skip, and a set of
	// paragraph indices where a page break should be inserted.
	//
	// In the cover page area (before the first heading), large runs of
	// consecutive empty paragraphs (>=10) act as visual page breaks in the
	// original DOC. We replace them with an explicit page break so that
	// "版权声明" and "修订记录" each start on their own page.
	//
	// In the body area (after the first heading), consecutive empty paragraphs
	// are collapsed to at most 1 to prevent blank pages.
	skipPara := make(map[int]bool)
	insertPageBreakBefore := make(map[int]bool) // insert a page break before this paragraph
	{
		// Find the first heading index (boundary between cover and body)
		firstHeadingIdx := len(fd.Paragraphs)
		for idx, p := range fd.Paragraphs {
			if p.HeadingLevel > 0 && !p.IsTOC {
				firstHeadingIdx = idx
				break
			}
		}

		// --- Cover page area: before first heading ---
		// The reference DOCX uses page break paragraphs (<w:br w:type="page"/>)
		// to separate cover sections, with no large empty paragraph runs.
		// We collapse all consecutive empties in the cover area to at most 1,
		// and insert page break paragraphs before key section titles
		// (版权声明, 修订记录) that should start on their own page.
		{
			// First pass: collapse consecutive empties to max 1 in cover area
			consecutiveEmpty := 0
			for idx := 0; idx < firstHeadingIdx; idx++ {
				p := fd.Paragraphs[idx]
				structural := p.IsSectionBreak || p.HasPageBreak || p.PageBreakBefore ||
					p.HeadingLevel > 0 || p.IsTOC || p.InTable || p.IsListItem
				if isEmptyParagraph(p) && !structural {
					consecutiveEmpty++
					if consecutiveEmpty > 1 {
						skipPara[idx] = true
					}
				} else {
					consecutiveEmpty = 0
				}
			}

			// Second pass: find paragraphs that need a page break before them.
			// In the DOC source, the cover area has multiple runs of consecutive
			// empty paragraphs. The first large run (>=10) is spacing within the
			// cover page (pushing address to bottom). Subsequent runs of >=4
			// empties are visual page separators between sections (cover→copyright,
			// copyright→revision history). We skip the first large run and insert
			// page breaks for subsequent ones.
			runLen := 0
			firstLargeRunSeen := false
			for idx := 0; idx < firstHeadingIdx; idx++ {
				p := fd.Paragraphs[idx]
				structural := p.IsSectionBreak || p.HasPageBreak || p.PageBreakBefore ||
					p.HeadingLevel > 0 || p.IsTOC || p.InTable || p.IsListItem
				if isEmptyParagraph(p) && !structural {
					runLen++
				} else {
					if runLen >= 10 && !firstLargeRunSeen {
						// First large run: just spacing within cover page
						firstLargeRunSeen = true
					} else if runLen >= 4 && firstLargeRunSeen && !isEmptyParagraph(p) && !p.InTable && !p.IsTOC {
						// Subsequent run after first large run: insert page break
						insertPageBreakBefore[idx] = true
					}
					runLen = 0
				}
			}

			// Third pass: skip the single remaining empty paragraph immediately
			// before each page break insertion point (the collapsed empty from
			// the original run). This matches the reference DOCX where the page
			// break paragraph directly follows the last content paragraph.
			for idx := range fd.Paragraphs {
				if insertPageBreakBefore[idx] {
					// Walk backwards to find the nearest non-skipped empty
					for j := idx - 1; j >= 0; j-- {
						if skipPara[j] {
							continue
						}
						pj := fd.Paragraphs[j]
						if isEmptyParagraph(pj) && !pj.IsSectionBreak && !pj.HasPageBreak {
							skipPara[j] = true
						}
						break
					}
				}
			}
		}

		// --- Body area: after first heading ---
		// Collapse consecutive empty paragraphs to at most 1.
		{
			consecutiveEmpty := 0
			for idx := firstHeadingIdx; idx < len(fd.Paragraphs); idx++ {
				p := fd.Paragraphs[idx]
				structural := p.IsSectionBreak || p.HasPageBreak || p.PageBreakBefore ||
					p.HeadingLevel > 0 || p.IsTOC || p.InTable || p.IsListItem
				if isEmptyParagraph(p) && !structural {
					consecutiveEmpty++
					if consecutiveEmpty > 1 {
						skipPara[idx] = true
					}
				} else {
					consecutiveEmpty = 0
				}
			}
		}

		// Skip empty paragraphs immediately after drawn-image paragraphs
		// in the body area (0x08 drawn object characters produce trailing empties).
		// In the cover area, empty paragraphs are preserved for spacing.
		for idx, p := range fd.Paragraphs {
			if idx < firstHeadingIdx {
				continue // preserve cover area empties
			}
			if len(p.DrawnImages) > 0 {
				for j := idx + 1; j < len(fd.Paragraphs); j++ {
					pj := fd.Paragraphs[j]
					if isEmptyParagraph(pj) && !pj.IsSectionBreak && !pj.HasPageBreak &&
						!pj.InTable && pj.HeadingLevel == 0 && !pj.IsTOC {
						skipPara[j] = true
					} else {
						break
					}
				}
			}
		}

		// Skip empty paragraphs immediately before headings or page breaks
		// (body area only)
		for idx := firstHeadingIdx; idx < len(fd.Paragraphs); idx++ {
			p := fd.Paragraphs[idx]
			if !isEmptyParagraph(p) || p.IsSectionBreak || p.InTable ||
				p.HasPageBreak || p.PageBreakBefore {
				continue
			}
			if idx+1 < len(fd.Paragraphs) {
				next := fd.Paragraphs[idx+1]
				if next.HeadingLevel > 0 || next.HasPageBreak || next.PageBreakBefore {
					skipPara[idx] = true
				}
			}
		}

		// Skip empty paragraphs between last TOC entry and first heading.
		lastTOCIdx := -1
		for idx := range fd.Paragraphs {
			if fd.Paragraphs[idx].IsTOC {
				lastTOCIdx = idx
			}
		}
		if lastTOCIdx >= 0 && firstHeadingIdx < len(fd.Paragraphs) {
			for idx := lastTOCIdx + 1; idx < firstHeadingIdx; idx++ {
				if isEmptyParagraph(fd.Paragraphs[idx]) && !fd.Paragraphs[idx].IsSectionBreak {
					skipPara[idx] = true
				}
			}
		}
	}

	i := 0
	inTOC := false

	// Handle leading empty section-break paragraphs.
	// In DOC format, the first paragraph often has a section break that defines
	// the first section's properties (like titlePg for first-page header).
	// We skip these empty section break paragraphs to avoid creating a blank
	// first page. Instead, we collect the section properties and attach them
	// to the first non-empty content paragraph.
	var pendingSectBreak *formattedParagraph
	for i < len(fd.Paragraphs) {
		p := fd.Paragraphs[i]
		if p.IsSectionBreak && isEmptyParagraph(p) {
			pCopy := p
			pendingSectBreak = &pCopy
			i++
			continue
		}
		break
	}

	for i < len(fd.Paragraphs) {
		p := fd.Paragraphs[i]

		// Skip paragraphs marked for removal by the pre-filter
		if skipPara[i] {
			i++
			continue
		}

		// Insert a page break before this paragraph if the pre-filter
		// determined that a large run of empty paragraphs was replaced.
		if insertPageBreakBefore[i] {
			buf.WriteString(`<w:p><w:r><w:br w:type="page"/></w:r></w:p>`)
		}

		// Emit pending section break before the first heading paragraph.
		// Instead of creating a separate section, we just clear the pending
		// flag. The document will use a single section (the final sectPr)
		// with titlePg, matching the reference file structure.
		if pendingSectBreak != nil && p.HeadingLevel > 0 {
			pendingSectBreak = nil
		}

		// Handle TOC field wrapping
		if p.IsTOC && !inTOC {
			// Start TOC field
			buf.WriteString(`<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>`)
			buf.WriteString(`<w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText></w:r>`)
			buf.WriteString(`<w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>`)
			inTOC = true
		}
		if !p.IsTOC && inTOC {
			// End TOC field
			buf.WriteString(`<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`)
			inTOC = false
		}

		if p.InTable {
			writeTableXML(&buf, fd.Paragraphs, &i, listInfo, imgRels, &imgIdx, images, fd)
			continue
		}

		// For heading paragraphs, wrap content with bookmarks for TOC linking
		if bm, ok := headingBookmarks[i]; ok {
			writeParagraphXMLWithBookmark(&buf, p, listInfo, imgRels, &imgIdx, images, i, fd, bm.bookmarkName, bm.bookmarkID)
		} else if target, ok := tocTargets[i]; ok {
			// For TOC entries, wrap content in a hyperlink
			writeParagraphXMLWithHyperlink(&buf, p, listInfo, imgRels, &imgIdx, images, i, fd, target)
		} else {
			writeParagraphXML(&buf, p, listInfo, imgRels, &imgIdx, images, i, fd)
		}
		i++
	}

	// Close TOC field if still open
	if inTOC {
		buf.WriteString(`<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`)
	}

	// Note: remaining images that weren't placed inline or as drawn objects
	// are not dumped at the end - they may be used in headers/footers or
	// other parts of the document that are handled separately.

	// Add section properties with header/footer references.
	// Single section for the entire document with titlePg so the first page
	// (cover page) uses the "first" header (background image) while all other
	// pages use the "default"/"even" headers.
	buf.WriteString(`<w:sectPr>`)
	if len(fd.HeaderEntries) > 0 || len(fd.FooterEntries) > 0 {
		for i, he := range fd.HeaderEntries {
			fmt.Fprintf(&buf, `<w:headerReference w:type="%s" r:id="rHdr%d"/>`, he.Type, i+1)
		}
		for i, fe := range fd.FooterEntries {
			fmt.Fprintf(&buf, `<w:footerReference w:type="%s" r:id="rFtr%d"/>`, fe.Type, i+1)
		}
	}
	// Add titlePg if there's a "first" header
	hasFirstHdr := false
	for _, he := range fd.HeaderEntries {
		if he.Type == "first" {
			hasFirstHdr = true
			break
		}
	}
	if hasFirstHdr {
		buf.WriteString(`<w:titlePg/>`)
	}
	buf.WriteString(`<w:pgSz w:w="11906" w:h="16838"/>`)
	buf.WriteString(`<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>`)
	buf.WriteString(`</w:sectPr>`)

	buf.WriteString(`</w:body></w:document>`)
	_, err := w.Write(buf.Bytes())
	return err
}

// writeParagraphXML writes a single <w:p> element for a non-table paragraph.
func writeParagraphXML(buf *bytes.Buffer, p formattedParagraph, listInfo *listNumInfo, imgRels []imageRel, imgIdx *int, images []imageData, paraIndex int, fd *formattedData) {
	buf.WriteString(`<w:p>`)

	writeParagraphProperties(buf, p, listInfo, paraIndex, fd)

	// Write text box if this paragraph has one (before drawn images to preserve DOC order)
	if p.TextBoxText != "" {
		writeTextBoxDrawing(buf, p.TextBoxText)
	}

	// Write drawn object images (from 0x08 characters mapped via PlcSpaMom)
	// Deduplicate: only render each unique BSE index once per paragraph.
	// When a paragraph has multiple drawn images, render only the largest one
	// (the main diagram) and skip small decorative elements (arrows, icons).
	drawnSeen := make(map[int]bool)
	uniqueDrawn := []int{}
	for _, bseIdx := range p.DrawnImages {
		if bseIdx >= 0 && bseIdx < len(imgRels) && !drawnSeen[bseIdx] {
			drawnSeen[bseIdx] = true
			uniqueDrawn = append(uniqueDrawn, bseIdx)
		}
	}
	// If multiple drawn images in one paragraph, only render the largest one
	// (the main content image). Small images are likely decorative shape elements.
	if len(uniqueDrawn) > 1 {
		bestIdx := uniqueDrawn[0]
		bestSize := 0
		if bestIdx < len(images) {
			bestSize = len(images[bestIdx].Data)
		}
		for _, bseIdx := range uniqueDrawn[1:] {
			sz := 0
			if bseIdx < len(images) {
				sz = len(images[bseIdx].Data)
			}
			if sz > bestSize {
				bestIdx = bseIdx
				bestSize = sz
			}
		}
		uniqueDrawn = []int{bestIdx}
	}
	for _, bseIdx := range uniqueDrawn {
		var imgPtr *imageData
		if bseIdx < len(images) {
			imgPtr = &images[bseIdx]
		}
		buf.WriteString(`<w:r>`)
		writeDrawnImageDrawing(buf, imgRels[bseIdx], bseIdx+1, imgPtr)
		buf.WriteString(`</w:r>`)
	}

	// Write runs, handling image placeholders (\x01)
	// If this paragraph has drawn images, skip whitespace-only text runs
	// (the 0x08 drawn object character is often represented as \t in the text stream)
	hasDrawn := len(uniqueDrawn) > 0
	for _, r := range p.Runs {
		if r.Text == "" {
			continue
		}
		// Skip whitespace-only text in paragraphs with drawn images
		// (the 0x08 drawn object character appears as \t in the text stream)
		if hasDrawn && strings.TrimSpace(r.Text) == "" {
			continue
		}
		// Split on \x01 to find inline image insertion points
		segments := strings.Split(r.Text, "\x01")
		for si, seg := range segments {
			if si > 0 {
				// Determine which BSE image this inline image references
				bseIdx := -1
				if r.ImageRef >= 0 {
					bseIdx = r.ImageRef
				} else if *imgIdx < len(imgRels) {
					// Fallback: use sequential counter
					bseIdx = *imgIdx
				}
				if bseIdx >= 0 && bseIdx < len(imgRels) {
					var imgPtr *imageData
					if bseIdx < len(images) {
						imgPtr = &images[bseIdx]
					}
					buf.WriteString(`<w:r>`)
					// Always use inline mode for images referenced by \x01.
					// Anchor mode can cause layout issues where the image
					// doesn't occupy flow space, pushing subsequent text
					// to a near-empty next page.
					writeImageDrawing(buf, imgRels[bseIdx], bseIdx+1, imgPtr)
					buf.WriteString(`</w:r>`)
				}
				*imgIdx++
			}
			if seg != "" {
				buf.WriteString(`<w:r>`)
				writeRunProperties(buf, r.Props)
				buf.WriteString(`<w:t xml:space="preserve">`)
				xml.Escape(buf, []byte(seg))
				buf.WriteString(`</w:t></w:r>`)
			}
		}
	}

	if p.HasPageBreak {
		buf.WriteString(`<w:r><w:br w:type="page"/></w:r>`)
	}

	buf.WriteString(`</w:p>`)
}

// writeParagraphXMLWithBookmark writes a paragraph with a bookmark for TOC linking.
func writeParagraphXMLWithBookmark(buf *bytes.Buffer, p formattedParagraph, listInfo *listNumInfo, imgRels []imageRel, imgIdx *int, images []imageData, paraIndex int, fd *formattedData, bookmarkName string, bookmarkID int) {
	buf.WriteString(`<w:p>`)
	writeParagraphProperties(buf, p, listInfo, paraIndex, fd)
	// Bookmark start
	fmt.Fprintf(buf, `<w:bookmarkStart w:id="%d" w:name="%s"/>`, bookmarkID, bookmarkName)

	// Write the same content as writeParagraphXML (text box, drawn images, runs)
	if p.TextBoxText != "" {
		writeTextBoxDrawing(buf, p.TextBoxText)
	}
	writeDrawnAndRuns(buf, p, imgRels, imgIdx, images)

	// Bookmark end
	fmt.Fprintf(buf, `<w:bookmarkEnd w:id="%d"/>`, bookmarkID)

	if p.HasPageBreak {
		buf.WriteString(`<w:r><w:br w:type="page"/></w:r>`)
	}
	buf.WriteString(`</w:p>`)
}

// writeParagraphXMLWithHyperlink writes a TOC paragraph with a clickable hyperlink.
func writeParagraphXMLWithHyperlink(buf *bytes.Buffer, p formattedParagraph, listInfo *listNumInfo, imgRels []imageRel, imgIdx *int, images []imageData, paraIndex int, fd *formattedData, anchor string) {
	buf.WriteString(`<w:p>`)
	writeParagraphProperties(buf, p, listInfo, paraIndex, fd)

	// Wrap all runs in a hyperlink
	fmt.Fprintf(buf, `<w:hyperlink w:anchor="%s" w:history="1">`, anchor)
	for _, r := range p.Runs {
		if r.Text == "" {
			continue
		}
		buf.WriteString(`<w:r>`)
		writeRunProperties(buf, r.Props)
		buf.WriteString(`<w:t xml:space="preserve">`)
		xml.Escape(buf, []byte(r.Text))
		buf.WriteString(`</w:t></w:r>`)
	}
	buf.WriteString(`</w:hyperlink>`)

	if p.HasPageBreak {
		buf.WriteString(`<w:r><w:br w:type="page"/></w:r>`)
	}
	buf.WriteString(`</w:p>`)
}

// writeDrawnAndRuns writes drawn images and text runs for a paragraph.
// This is the shared logic extracted from writeParagraphXML.
func writeDrawnAndRuns(buf *bytes.Buffer, p formattedParagraph, imgRels []imageRel, imgIdx *int, images []imageData) {
	drawnSeen := make(map[int]bool)
	uniqueDrawn := []int{}
	for _, bseIdx := range p.DrawnImages {
		if bseIdx >= 0 && bseIdx < len(imgRels) && !drawnSeen[bseIdx] {
			drawnSeen[bseIdx] = true
			uniqueDrawn = append(uniqueDrawn, bseIdx)
		}
	}
	if len(uniqueDrawn) > 1 {
		bestIdx := uniqueDrawn[0]
		bestSize := 0
		if bestIdx < len(images) {
			bestSize = len(images[bestIdx].Data)
		}
		for _, bseIdx := range uniqueDrawn[1:] {
			sz := 0
			if bseIdx < len(images) {
				sz = len(images[bseIdx].Data)
			}
			if sz > bestSize {
				bestIdx = bseIdx
				bestSize = sz
			}
		}
		uniqueDrawn = []int{bestIdx}
	}
	for _, bseIdx := range uniqueDrawn {
		var imgPtr *imageData
		if bseIdx < len(images) {
			imgPtr = &images[bseIdx]
		}
		buf.WriteString(`<w:r>`)
		writeDrawnImageDrawing(buf, imgRels[bseIdx], bseIdx+1, imgPtr)
		buf.WriteString(`</w:r>`)
	}

	hasDrawn := len(uniqueDrawn) > 0
	for _, r := range p.Runs {
		if r.Text == "" {
			continue
		}
		if hasDrawn && strings.TrimSpace(r.Text) == "" {
			continue
		}
		segments := strings.Split(r.Text, "\x01")
		for si, seg := range segments {
			if si > 0 {
				bseIdx := -1
				if r.ImageRef >= 0 {
					bseIdx = r.ImageRef
				} else if *imgIdx < len(imgRels) {
					bseIdx = *imgIdx
				}
				if bseIdx >= 0 && bseIdx < len(imgRels) {
					var imgPtr *imageData
					if bseIdx < len(images) {
						imgPtr = &images[bseIdx]
					}
					buf.WriteString(`<w:r>`)
					writeImageDrawing(buf, imgRels[bseIdx], bseIdx+1, imgPtr)
					buf.WriteString(`</w:r>`)
				}
				*imgIdx++
			}
			if seg != "" {
				buf.WriteString(`<w:r>`)
				writeRunProperties(buf, r.Props)
				buf.WriteString(`<w:t xml:space="preserve">`)
				xml.Escape(buf, []byte(seg))
				buf.WriteString(`</w:t></w:r>`)
			}
		}
	}
}

// writeTableXML writes a <w:tbl> element for consecutive InTable paragraphs.
// It advances *idx past all consumed table paragraphs.
func writeTableXML(buf *bytes.Buffer, paragraphs []formattedParagraph, idx *int, listInfo *listNumInfo, imgRels []imageRel, imgIdx *int, images []imageData, fd *formattedData) {
	buf.WriteString(`<w:tbl>`)

	// Write table properties with default borders
	buf.WriteString(`<w:tblPr>`)
	buf.WriteString(`<w:tblW w:w="0" w:type="auto"/>`)
	buf.WriteString(`<w:tblBorders>`)
	for _, side := range []string{"top", "left", "bottom", "right", "insideH", "insideV"} {
		fmt.Fprintf(buf, `<w:%s w:val="single" w:sz="4" w:space="0" w:color="auto"/>`, side)
	}
	buf.WriteString(`</w:tblBorders>`)
	buf.WriteString(`</w:tblPr>`)

	// Group paragraphs into rows. Each row ends with TableRowEnd=true.
	// Within a row, cells are delimited by IsTableCellEnd=true.
	// A cell can contain multiple paragraphs (multi-paragraph cells) where
	// intermediate paragraphs have IsTableCellEnd=false.
	for *idx < len(paragraphs) && paragraphs[*idx].InTable {
		// Collect all paragraphs for this row (up to and including row-end)
		type indexedPara struct {
			p        formattedParagraph
			origIdx  int
		}
		var rowParas []indexedPara
		var rowEndPara *formattedParagraph
		for *idx < len(paragraphs) && paragraphs[*idx].InTable {
			p := paragraphs[*idx]
			origIdx := *idx
			*idx++
			if p.TableRowEnd {
				rowEndPara = &p
				break
			}
			rowParas = append(rowParas, indexedPara{p: p, origIdx: origIdx})
		}

		// Group paragraphs into cells using IsTableCellEnd markers.
		// Each cell ends when IsTableCellEnd=true. Paragraphs with
		// IsTableCellEnd=false are continuation paragraphs within the same cell.
		type cellContent struct {
			paragraphs []indexedPara
		}
		var cells []cellContent
		var currentCell []indexedPara
		for _, ip := range rowParas {
			currentCell = append(currentCell, ip)
			if ip.p.IsTableCellEnd {
				cells = append(cells, cellContent{paragraphs: currentCell})
				currentCell = nil
			}
		}
		// If there are remaining paragraphs without a cell-end marker,
		// append them to the last cell or create a new cell
		if len(currentCell) > 0 {
			if len(cells) > 0 {
				cells[len(cells)-1].paragraphs = append(cells[len(cells)-1].paragraphs, currentCell...)
			} else {
				cells = append(cells, cellContent{paragraphs: currentCell})
			}
		}

		buf.WriteString(`<w:tr>`)
		for ci, cell := range cells {
			buf.WriteString(`<w:tc>`)
			// Write cell properties with width if available
			if rowEndPara != nil && ci < len(rowEndPara.CellWidths) {
				w := rowEndPara.CellWidths[ci]
				if w > 0 {
					fmt.Fprintf(buf, `<w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>`, w)
				}
			}
			if len(cell.paragraphs) > 0 {
				for _, ip := range cell.paragraphs {
					writeParagraphXML(buf, ip.p, listInfo, imgRels, imgIdx, images, ip.origIdx, fd)
				}
			} else {
				// Empty cell - write an empty paragraph (required by OOXML)
				buf.WriteString(`<w:p/>`)
			}
			buf.WriteString(`</w:tc>`)
		}
		buf.WriteString(`</w:tr>`)
	}

	buf.WriteString(`</w:tbl>`)
}

// writeParagraphProperties writes <w:pPr> element for a formatted paragraph.
func isImageOnlyParagraph(p formattedParagraph) bool {
	// A paragraph is image-only if it has images (drawn or inline) but no visible text.
	hasImage := len(p.DrawnImages) > 0
	if !hasImage {
		for _, r := range p.Runs {
			if r.ImageRef >= 0 || strings.Contains(r.Text, "\x01") {
				hasImage = true
				break
			}
		}
	}
	if !hasImage {
		return false
	}
	// Check if there's any visible text
	for _, r := range p.Runs {
		text := strings.ReplaceAll(r.Text, "\x01", "")
		text = strings.TrimSpace(text)
		if text != "" {
			return false
		}
	}
	return true
}

func writeParagraphProperties(buf *bytes.Buffer, p formattedParagraph, listInfo *listNumInfo, paraIndex int, fd *formattedData) {
	hasPPr := false
	var pprBuf bytes.Buffer

	// For image-only paragraphs, override formatting: center alignment,
	// no text indent, minimal spacing. This prevents body text formatting
	// (firstLine indent, line spacing) from creating unwanted whitespace
	// around standalone images.
	imgOnly := isImageOnlyParagraph(p)

	// Heading style reference (must be first inside <w:pPr>)
	if p.HeadingLevel >= 1 && p.HeadingLevel <= 9 {
		fmt.Fprintf(&pprBuf, `<w:pStyle w:val="Heading%d"/>`, p.HeadingLevel)
		hasPPr = true
	} else if p.IsTOC && p.TOCLevel >= 1 && p.TOCLevel <= 3 {
		fmt.Fprintf(&pprBuf, `<w:pStyle w:val="TOC%d"/>`, p.TOCLevel)
		hasPPr = true
	}

	// List numbering properties
	if p.IsListItem && listInfo != nil && paraIndex >= 0 {
		if numId, ok := listInfo.paraNumId[paraIndex]; ok {
			fmt.Fprintf(&pprBuf, `<w:numPr><w:ilvl w:val="%d"/><w:numId w:val="%d"/></w:numPr>`, p.ListLevel, numId)
			hasPPr = true
		}
	}

	// PageBreakBefore
	if p.PageBreakBefore {
		pprBuf.WriteString(`<w:pageBreakBefore/>`)
		hasPPr = true
	}

	// Alignment - image-only paragraphs are always centered
	if imgOnly {
		pprBuf.WriteString(`<w:jc w:val="center"/>`)
		hasPPr = true
	} else {
		alignVal := ""
		switch p.Props.Alignment {
		case 0:
			if p.Props.AlignmentSet && p.HeadingLevel == 0 && !p.IsTOC {
				alignVal = "left"
			}
		case 1:
			alignVal = "center"
		case 2:
			alignVal = "right"
		case 3:
			alignVal = "both"
		}
		if alignVal != "" {
			fmt.Fprintf(&pprBuf, `<w:jc w:val="%s"/>`, alignVal)
			hasPPr = true
		}
	}

	// Indentation - skip for image-only paragraphs
	if !imgOnly {
		hasIndent := p.Props.IndentLeft != 0 || p.Props.IndentRight != 0 || p.Props.IndentFirst != 0
		if hasIndent {
			pprBuf.WriteString(`<w:ind`)
			if p.Props.IndentLeft != 0 {
				fmt.Fprintf(&pprBuf, ` w:left="%d"`, p.Props.IndentLeft)
			}
			if p.Props.IndentRight != 0 {
				fmt.Fprintf(&pprBuf, ` w:right="%d"`, p.Props.IndentRight)
			}
			if p.Props.IndentFirst != 0 {
				if p.Props.IndentFirst > 0 {
					fmt.Fprintf(&pprBuf, ` w:firstLine="%d"`, p.Props.IndentFirst)
				} else {
					fmt.Fprintf(&pprBuf, ` w:hanging="%d"`, -p.Props.IndentFirst)
				}
			}
			pprBuf.WriteString(`/>`)
			hasPPr = true
		}
	}

	// Spacing - image-only paragraphs get minimal spacing
	if imgOnly {
		// Only emit spacing before if original had it, but skip line spacing
		if p.Props.SpaceBefore != 0 {
			fmt.Fprintf(&pprBuf, `<w:spacing w:before="%d"/>`, p.Props.SpaceBefore)
			hasPPr = true
		}
	} else {
		hasSpacing := p.Props.SpaceBefore != 0 || p.Props.SpaceAfter != 0 || p.Props.LineSpacing != 0
		if hasSpacing {
			pprBuf.WriteString(`<w:spacing`)
			if p.Props.SpaceBefore != 0 {
				fmt.Fprintf(&pprBuf, ` w:before="%d"`, p.Props.SpaceBefore)
			}
			if p.Props.SpaceAfter != 0 {
				fmt.Fprintf(&pprBuf, ` w:after="%d"`, p.Props.SpaceAfter)
			}
			if p.Props.LineSpacing != 0 {
				fmt.Fprintf(&pprBuf, ` w:line="%d"`, p.Props.LineSpacing)
				switch p.Props.LineRule {
				case 0:
					pprBuf.WriteString(` w:lineRule="auto"`)
				case 1:
					pprBuf.WriteString(` w:lineRule="atLeast"`)
				case 2:
					pprBuf.WriteString(` w:lineRule="exact"`)
				}
			}
			pprBuf.WriteString(`/>`)
			hasPPr = true
		}
	}

	// Section break
	if p.IsSectionBreak {
		pprBuf.WriteString(`<w:sectPr>`)
		// Include header/footer references so every section shows headers/footers
		if fd != nil {
			for hi, he := range fd.HeaderEntries {
				fmt.Fprintf(&pprBuf, `<w:headerReference w:type="%s" r:id="rHdr%d"/>`, he.Type, hi+1)
			}
			for fi, fe := range fd.FooterEntries {
				fmt.Fprintf(&pprBuf, `<w:footerReference w:type="%s" r:id="rFtr%d"/>`, fe.Type, fi+1)
			}
			// Enable titlePg if there's a "first" header or footer
			hasFirst := false
			for _, he := range fd.HeaderEntries {
				if he.Type == "first" {
					hasFirst = true
					break
				}
			}
			if !hasFirst {
				for _, fe := range fd.FooterEntries {
					if fe.Type == "first" {
						hasFirst = true
						break
					}
				}
			}
			if hasFirst {
				pprBuf.WriteString(`<w:titlePg/>`)
			}
		}
		pprBuf.WriteString(`<w:pgSz w:w="11906" w:h="16838"/>`)
		pprBuf.WriteString(`<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>`)
		switch p.SectionType {
		case 0:
			pprBuf.WriteString(`<w:type w:val="continuous"/>`)
		case 1:
			pprBuf.WriteString(`<w:type w:val="nextPage"/>`)
		case 2:
			pprBuf.WriteString(`<w:type w:val="evenPage"/>`)
		case 3:
			pprBuf.WriteString(`<w:type w:val="oddPage"/>`)
		}
		pprBuf.WriteString(`</w:sectPr>`)
		hasPPr = true
	}

	if hasPPr {
		buf.WriteString(`<w:pPr>`)
		buf.Write(pprBuf.Bytes())
		buf.WriteString(`</w:pPr>`)
	}
}


// writeRunProperties writes <w:rPr> element for character formatting.
func writeRunProperties(buf *bytes.Buffer, props doc.CharacterFormatting) {
	hasRPr := false
	var rprBuf bytes.Buffer

	if props.FontName != "" {
		// Detect if font is CJK (Chinese/Japanese/Korean) for eastAsia attribute
		isCJK := isCJKFont(props.FontName)
		if isCJK {
			fmt.Fprintf(&rprBuf, `<w:rFonts w:eastAsia="%s"/>`, xmlEscapeAttr(props.FontName))
		} else {
			fmt.Fprintf(&rprBuf, `<w:rFonts w:ascii="%s" w:hAnsi="%s"/>`, xmlEscapeAttr(props.FontName), xmlEscapeAttr(props.FontName))
		}
		hasRPr = true
	}
	if props.Bold {
		rprBuf.WriteString(`<w:b/><w:bCs/>`)
		hasRPr = true
	}
	if props.Italic {
		rprBuf.WriteString(`<w:i/><w:iCs/>`)
		hasRPr = true
	}
	if props.Underline > 0 {
		rprBuf.WriteString(`<w:u w:val="single"/>`)
		hasRPr = true
	}
	if props.FontSize > 0 {
		fmt.Fprintf(&rprBuf, `<w:sz w:val="%d"/><w:szCs w:val="%d"/>`, props.FontSize, props.FontSize)
		hasRPr = true
	}
	if props.Color != "" {
		fmt.Fprintf(&rprBuf, `<w:color w:val="%s"/>`, props.Color)
		hasRPr = true
	}

	if hasRPr {
		buf.WriteString(`<w:rPr>`)
		buf.Write(rprBuf.Bytes())
		buf.WriteString(`</w:rPr>`)
	}
}

// isEmptyParagraph returns true if the paragraph has no visible text content.
func isEmptyParagraph(p formattedParagraph) bool {
	for _, r := range p.Runs {
		text := strings.TrimSpace(r.Text)
		text = strings.ReplaceAll(text, "\r", "")
		text = strings.ReplaceAll(text, "\n", "")
		if text != "" {
			return false
		}
	}
	if p.TextBoxText != "" {
		return false
	}
	if len(p.DrawnImages) > 0 {
		return false
	}
	return true
}

// isCJKFont returns true if the font name appears to be a CJK font.
func isCJKFont(name string) bool {
	for _, r := range name {
		if r >= 0x2E80 { // CJK character range
			return true
		}
	}
	// Common CJK font names in ASCII
	switch name {
	case "SimSun", "SimHei", "FangSong", "KaiTi", "Microsoft YaHei",
		"MS Mincho", "MS Gothic", "Malgun Gothic", "Batang":
		return true
	}
	return false
}

// xmlEscapeAttr escapes special XML characters in attribute values.
func xmlEscapeAttr(s string) string {
	s = strings.ReplaceAll(s, "&", "&amp;")
	s = strings.ReplaceAll(s, "\"", "&quot;")
	s = strings.ReplaceAll(s, "<", "&lt;")
	s = strings.ReplaceAll(s, ">", "&gt;")
	return s
}

// writeDocxFormatted generates a DOCX zip archive with formatting information.
func writeDocxFormatted(w io.Writer, fd *formattedData, images []imageData) error {
	zw := zip.NewWriter(w)
	defer zw.Close()

	// Write images into word/media/ and collect relationship info
	var imgRels []imageRel
	for i, img := range images {
		ext := (&common.Image{Format: img.Format}).Extension()
		if ext == "" {
			ext = ".bin"
		}
		filename := fmt.Sprintf("image%d%s", i+1, ext)
		fw, err := zw.Create("word/media/" + filename)
		if err != nil {
			return err
		}
		if _, err := fw.Write(img.Data); err != nil {
			return err
		}
		imgRels = append(imgRels, imageRel{filename: filename, relID: fmt.Sprintf("rImg%d", i+1)})
	}

	// Write formatted document.xml
	fw, err := zw.Create("word/document.xml")
	if err != nil {
		return err
	}
	if err := writeFormattedDocumentXML(fw, fd, imgRels, images); err != nil {
		return err
	}

	// Write header/footer XML parts using structured entries
	var hdrRelIDs []string
	var hdrTypes []string
	var ftrRelIDs []string
	var ftrTypes []string

	for i, he := range fd.HeaderEntries {
		filename := fmt.Sprintf("header%d.xml", i+1)
		relID := fmt.Sprintf("rHdr%d", i+1)

		// Build header-specific image rels
		var hdrImgRels []imageRel
		for _, bseIdx := range he.Images {
			if bseIdx >= 0 && bseIdx < len(imgRels) {
				hdrImgRels = append(hdrImgRels, imgRels[bseIdx])
			}
		}

		fw, err := zw.Create("word/" + filename)
		if err != nil {
			return err
		}
		if err := writeHeaderXML(fw, he.Text, hdrImgRels, images, he.Type); err != nil {
			return err
		}

		// Write header rels file if it has images
		if len(hdrImgRels) > 0 {
			rfw, err := zw.Create(fmt.Sprintf("word/_rels/%s.rels", filename))
			if err != nil {
				return err
			}
			if err := writePartRels(rfw, hdrImgRels); err != nil {
				return err
			}
		}

		hdrRelIDs = append(hdrRelIDs, relID)
		hdrTypes = append(hdrTypes, he.Type)
	}

	for i, fe := range fd.FooterEntries {
		filename := fmt.Sprintf("footer%d.xml", i+1)
		relID := fmt.Sprintf("rFtr%d", i+1)
		fw, err := zw.Create("word/" + filename)
		if err != nil {
			return err
		}
		if err := writeFooterXML(fw, fe.Text, fe.RawText); err != nil {
			return err
		}
		ftrRelIDs = append(ftrRelIDs, relID)
		ftrTypes = append(ftrTypes, fe.Type)
	}

	// Write document.xml.rels
	hasNumbering := hasListParagraphs(fd)
	fw, err = zw.Create("word/_rels/document.xml.rels")
	if err != nil {
		return err
	}
	if err := writeDocumentRelsFormatted(fw, imgRels, hasNumbering, hdrRelIDs, hdrTypes, ftrRelIDs, ftrTypes); err != nil {
		return err
	}

	// Write styles.xml (with heading style definitions)
	fw, err = zw.Create("word/styles.xml")
	if err != nil {
		return err
	}
	if err := writeFormattedStylesXML(fw, fd); err != nil {
		return err
	}

	// Write settings.xml for document-level settings
	fw, err = zw.Create("word/settings.xml")
	if err != nil {
		return err
	}
	{
		var settingsBuf bytes.Buffer
		settingsBuf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
		settingsBuf.WriteString(`<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`)
		// Enable evenAndOddHeaders if there's an "even" header or footer
		hasEvenHdr := false
		for _, he := range fd.HeaderEntries {
			if he.Type == "even" {
				hasEvenHdr = true
				break
			}
		}
		if !hasEvenHdr {
			for _, fe := range fd.FooterEntries {
				if fe.Type == "even" {
					hasEvenHdr = true
					break
				}
			}
		}
		if hasEvenHdr {
			settingsBuf.WriteString(`<w:evenAndOddHeaders/>`)
		}
		settingsBuf.WriteString(`<w:compat><w:useFELayout/></w:compat>`)
		settingsBuf.WriteString(`<w:defaultTabStop w:val="420"/>`)
		settingsBuf.WriteString(`</w:settings>`)
		if _, err := fw.Write(settingsBuf.Bytes()); err != nil {
			return err
		}
	}

	// Write numbering.xml if there are list paragraphs
	if hasNumbering {
		fw, err = zw.Create("word/numbering.xml")
		if err != nil {
			return err
		}
		if err := writeNumberingXML(fw, fd); err != nil {
			return err
		}
	}

	// Write [Content_Types].xml
	fw, err = zw.Create("[Content_Types].xml")
	if err != nil {
		return err
	}
	if err := writeContentTypes(fw, imgRels, hasNumbering, len(fd.HeaderEntries), len(fd.FooterEntries)); err != nil {
		return err
	}

	// Write _rels/.rels
	fw, err = zw.Create("_rels/.rels")
	if err != nil {
		return err
	}
	if _, err := io.WriteString(fw, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`+
		`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>`+
		`</Relationships>`); err != nil {
		return err
	}

	return nil
}

func writeContentTypes(w io.Writer, imgRels []imageRel, hasNumbering bool, numHeaders int, numFooters int) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`)
	buf.WriteString(`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`)
	buf.WriteString(`<Default Extension="xml" ContentType="application/xml"/>`)
	buf.WriteString(`<Default Extension="jpeg" ContentType="image/jpeg"/>`)
	buf.WriteString(`<Default Extension="png" ContentType="image/png"/>`)
	buf.WriteString(`<Default Extension="emf" ContentType="image/x-emf"/>`)
	buf.WriteString(`<Default Extension="wmf" ContentType="image/x-wmf"/>`)
	buf.WriteString(`<Default Extension="tiff" ContentType="image/tiff"/>`)
	buf.WriteString(`<Default Extension="bmp" ContentType="image/bmp"/>`)
	buf.WriteString(`<Default Extension="pict" ContentType="image/pict"/>`)
	buf.WriteString(`<Default Extension="bin" ContentType="application/octet-stream"/>`)
	buf.WriteString(`<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`)
	buf.WriteString(`<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`)
	buf.WriteString(`<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>`)
	if hasNumbering {
		buf.WriteString(`<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`)
	}
	for i := 0; i < numHeaders; i++ {
		fmt.Fprintf(&buf, `<Override PartName="/word/header%d.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`, i+1)
	}
	for i := 0; i < numFooters; i++ {
		fmt.Fprintf(&buf, `<Override PartName="/word/footer%d.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`, i+1)
	}
	buf.WriteString(`</Types>`)
	_, err := w.Write(buf.Bytes())
	return err
}

// hasListParagraphs returns true if any paragraph in the formatted data is a list item.
func hasListParagraphs(fd *formattedData) bool {
	for _, p := range fd.Paragraphs {
		if p.IsListItem {
			return true
		}
	}
	return false
}

// listNumInfo holds the numbering assignment for the document.
// It maps each list paragraph index to a numId, and tracks the abstract
// numbering definitions needed.
type listNumInfo struct {
	// paraNumId maps paragraph index -> numId (1-based)
	paraNumId map[int]int
	// numDefs maps numId -> list type: 100=heading, 0=bullet, 1=ordered
	numDefs map[int]uint8
	// numNfc maps numId -> number format code (nfc) from the DOC list definition
	numNfc map[int]uint8
	// numLvlText maps numId -> level text template (e.g. "(%1)" or "%1.")
	numLvlText map[int]string
	// totalNums is the total number of numId values assigned
	totalNums int
}

// buildListNumInfo scans paragraphs and assigns numIds.
// Heading list items share one continuous numId.
// Non-heading list items get a new numId each time the ilfo changes or
// there's a gap (non-list paragraph) between groups.
func buildListNumInfo(fd *formattedData) *listNumInfo {
	info := &listNumInfo{
		paraNumId:  make(map[int]int),
		numDefs:    make(map[int]uint8),
		numNfc:     make(map[int]uint8),
		numLvlText: make(map[int]string),
	}

	nextId := 1
	headingNumId := 0

	// First pass: assign heading numId
	for _, p := range fd.Paragraphs {
		if p.IsListItem && p.HeadingLevel > 0 {
			if headingNumId == 0 {
				headingNumId = nextId
				info.numDefs[headingNumId] = 100
				nextId++
			}
			break
		}
	}

	// Second pass: assign numIds for non-heading lists.
	// Track the current ilfo/type and whether we're in a list group.
	var currentIlfo uint16
	var currentType uint8
	currentNumId := 0
	inListGroup := false

	for i, p := range fd.Paragraphs {
		if !p.IsListItem {
			inListGroup = false
			continue
		}

		if p.HeadingLevel > 0 {
			// Heading list items use the shared heading numId
			info.paraNumId[i] = headingNumId
			inListGroup = false // heading breaks non-heading list groups
			continue
		}

		// Non-heading list item: check if we need a new numId
		needNewNum := false
		if !inListGroup {
			needNewNum = true
		} else if p.ListIlfo != currentIlfo {
			needNewNum = true
		} else if p.ListType != currentType {
			needNewNum = true
		}

		if needNewNum {
			currentNumId = nextId
			info.numDefs[currentNumId] = p.ListType
			info.numNfc[currentNumId] = p.ListNfc
			info.numLvlText[currentNumId] = p.ListLvlText
			nextId++
			currentIlfo = p.ListIlfo
			currentType = p.ListType
			inListGroup = true
		}

		info.paraNumId[i] = currentNumId
	}

	info.totalNums = nextId - 1
	return info
}

// writeNumberingXML generates word/numbering.xml with abstract numbering definitions
// and numbering instances for each list group used in the document.
// Each independent list group (detected by ilfo changes or gaps) gets its own
// abstractNum and num definition, so numbering restarts correctly.
func writeNumberingXML(w io.Writer, fd *formattedData) error {
	info := buildListNumInfo(fd)
	if info.totalNums == 0 {
		return nil
	}

	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`)

	// Sort numIds for deterministic output
	numIds := make([]int, 0, len(info.numDefs))
	for nid := range info.numDefs {
		numIds = append(numIds, nid)
	}
	for i := 1; i < len(numIds); i++ {
		for j := i; j > 0 && numIds[j] < numIds[j-1]; j-- {
			numIds[j], numIds[j-1] = numIds[j-1], numIds[j]
		}
	}

	// Write abstractNum definitions
	for _, nid := range numIds {
		listType := info.numDefs[nid]
		absId := nid - 1
		fmt.Fprintf(&buf, `<w:abstractNum w:abstractNumId="%d">`, absId)

		if listType == 100 {
			// Heading numbering: multi-level decimal (1., 1.1., 1.1.1.)
			buf.WriteString(`<w:multiLevelType w:val="multilevel"/>`)
			for lvl := 0; lvl < 9; lvl++ {
				fmt.Fprintf(&buf, `<w:lvl w:ilvl="%d">`, lvl)
				buf.WriteString(`<w:start w:val="1"/>`)
				buf.WriteString(`<w:numFmt w:val="decimal"/>`)
				lvlText := ""
				for i := 0; i <= lvl; i++ {
					lvlText += fmt.Sprintf("%%%d.", i+1)
				}
				fmt.Fprintf(&buf, `<w:lvlText w:val="%s"/>`, lvlText)
				buf.WriteString(`<w:lvlJc w:val="left"/>`)
				buf.WriteString(`<w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr>`)
				buf.WriteString(`</w:lvl>`)
			}
		} else if listType == 1 {
			// Ordered list: use nfc to determine format
			nfc := info.numNfc[nid]
			numFmt := "decimal"
			switch nfc {
			case 1:
				numFmt = "upperRoman"
			case 2:
				numFmt = "lowerRoman"
			case 3:
				numFmt = "upperLetter"
			case 4:
				numFmt = "lowerLetter"
			case 5:
				numFmt = "ordinal"
			case 22:
				numFmt = "chineseCounting"
			case 38:
				numFmt = "chineseCountingThousand"
			default:
				numFmt = "decimal"
			}
			buf.WriteString(`<w:multiLevelType w:val="hybridMultilevel"/>`)
			for lvl := 0; lvl < 9; lvl++ {
				fmt.Fprintf(&buf, `<w:lvl w:ilvl="%d">`, lvl)
				buf.WriteString(`<w:start w:val="1"/>`)
				fmt.Fprintf(&buf, `<w:numFmt w:val="%s"/>`, numFmt)
				// Use lvlText from DOC if available for level 0, otherwise default
				if lvl == 0 && info.numLvlText[nid] != "" {
					fmt.Fprintf(&buf, `<w:lvlText w:val="%s"/>`, xmlEscapeAttr(info.numLvlText[nid]))
				} else {
					fmt.Fprintf(&buf, `<w:lvlText w:val="%%%d."/>`, lvl+1)
				}
				indent := 420 * (lvl + 1)
				hanging := 420
				fmt.Fprintf(&buf, `<w:pPr><w:ind w:left="%d" w:hanging="%d"/></w:pPr>`, indent, hanging)
				buf.WriteString(`</w:lvl>`)
			}
		} else {
			// Bullet list: unordered with bullet symbols
			buf.WriteString(`<w:multiLevelType w:val="hybridMultilevel"/>`)
			for lvl := 0; lvl < 9; lvl++ {
				fmt.Fprintf(&buf, `<w:lvl w:ilvl="%d">`, lvl)
				buf.WriteString(`<w:numFmt w:val="bullet"/>`)
				if lvl == 0 {
					buf.WriteString("<w:lvlText w:val=\"&#xF0B7;\"/>")
					buf.WriteString(`<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>`)
				} else if lvl == 1 {
					buf.WriteString(`<w:lvlText w:val="o"/>`)
					buf.WriteString(`<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:hint="default"/></w:rPr>`)
				} else {
					buf.WriteString("<w:lvlText w:val=\"&#xF0A7;\"/>")
					buf.WriteString(`<w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/></w:rPr>`)
				}
				indent := 420 * (lvl + 1)
				hanging := 420
				fmt.Fprintf(&buf, `<w:pPr><w:ind w:left="%d" w:hanging="%d"/></w:pPr>`, indent, hanging)
				buf.WriteString(`</w:lvl>`)
			}
		}
		buf.WriteString(`</w:abstractNum>`)
	}

	// Write num definitions
	for _, nid := range numIds {
		fmt.Fprintf(&buf, `<w:num w:numId="%d">`, nid)
		fmt.Fprintf(&buf, `<w:abstractNumId w:val="%d"/>`, nid-1)
		buf.WriteString(`</w:num>`)
	}

	buf.WriteString(`</w:numbering>`)
	_, err := w.Write(buf.Bytes())
	return err
}

// writeDocumentRelsFormatted writes document.xml.rels with header/footer relationships.
func writeDocumentRelsFormatted(w io.Writer, imgRels []imageRel, hasNumbering bool, hdrRelIDs []string, hdrTypes []string, ftrRelIDs []string, ftrTypes []string) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)
	buf.WriteString(`<Relationship Id="rIdStyles1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`)
	buf.WriteString(`<Relationship Id="rIdSettings1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>`)

	if hasNumbering {
		buf.WriteString(`<Relationship Id="rIdNumbering1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`)
	}

	for i, relID := range hdrRelIDs {
		fmt.Fprintf(&buf, `<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header%d.xml"/>`, relID, i+1)
	}

	for i, relID := range ftrRelIDs {
		fmt.Fprintf(&buf, `<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer%d.xml"/>`, relID, i+1)
	}

	for _, rel := range imgRels {
		fmt.Fprintf(&buf, `<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/%s"/>`, rel.relID, rel.filename)
	}

	buf.WriteString(`</Relationships>`)
	_, err := w.Write(buf.Bytes())
	return err
}

// writeHeaderXML writes a header XML part with optional image support.
func writeHeaderXML(w io.Writer, text string, hdrImgRels []imageRel, images []imageData, hdrType string) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`)
	buf.WriteString(` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`)
	buf.WriteString(` xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"`)
	buf.WriteString(` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`)
	buf.WriteString(` xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">`)
	buf.WriteString(`<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>`)

	// Write drawn images in header
	for i, rel := range hdrImgRels {
		// Find the image data by matching the rel filename to the global image list
		var imgPtr *imageData
		for j := range images {
			ext := (&common.Image{Format: images[j].Format}).Extension()
			if ext == "" {
				ext = ".bin"
			}
			fname := fmt.Sprintf("image%d%s", j+1, ext)
			if fname == rel.filename {
				imgPtr = &images[j]
				break
			}
		}

		// For "first" page header with a tall image (cover background),
		// render as an anchor behind document text instead of inline.
		// The image is positioned at the top-left of the page and scaled
		// to fill the full page width (A4: 11906 twips = 7562850 EMU).
		if hdrType == "first" && imgPtr != nil {
			emuW, emuH := getImageDimensionsEMU(imgPtr)
			// Full page width: 11906 twips * 635 EMU/twip = 7560310 EMU
			pageWidth := int64(7560310)
			if emuW > 0 && emuH > 0 {
				// Scale to full page width
				scale := float64(pageWidth) / float64(emuW)
				emuW = pageWidth
				emuH = int64(float64(emuH) * scale)
			}
			// If the image is taller than ~half the page, treat as background
			if emuH > 5000000 {
				buf.WriteString(`<w:r><w:drawing>`)
				buf.WriteString(`<wp:anchor distT="0" distB="0" distL="0" distR="0"`)
				buf.WriteString(` simplePos="0" relativeHeight="0" behindDoc="1" locked="1"`)
				buf.WriteString(` layoutInCell="1" allowOverlap="1">`)
				buf.WriteString(`<wp:simplePos x="0" y="0"/>`)
				buf.WriteString(`<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>`)
				buf.WriteString(`<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>`)
				fmt.Fprintf(&buf, `<wp:extent cx="%d" cy="%d"/>`, emuW, emuH)
				buf.WriteString(`<wp:effectExtent l="0" t="0" r="0" b="0"/>`)
				buf.WriteString(`<wp:wrapNone/>`)
				fmt.Fprintf(&buf, `<wp:docPr id="%d" name="Image %d"/>`, 100+i+1, 100+i+1)
				buf.WriteString(`<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">`)
				buf.WriteString(`<pic:pic><pic:nvPicPr>`)
				fmt.Fprintf(&buf, `<pic:cNvPr id="%d" name="Image %d"/>`, 100+i+1, 100+i+1)
				buf.WriteString(`<pic:cNvPicPr/></pic:nvPicPr>`)
				buf.WriteString(`<pic:blipFill>`)
				fmt.Fprintf(&buf, `<a:blip r:embed="%s"/>`, rel.relID)
				buf.WriteString(`<a:stretch><a:fillRect/></a:stretch>`)
				buf.WriteString(`</pic:blipFill>`)
				fmt.Fprintf(&buf, `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>`, emuW, emuH)
				buf.WriteString(`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`)
				buf.WriteString(`</pic:pic></a:graphicData></a:graphic>`)
				buf.WriteString(`</wp:anchor></w:drawing></w:r>`)
				continue
			}
		}

		buf.WriteString(`<w:r>`)
		writeHeaderImageDrawing(&buf, rel, 100+i+1, imgPtr)
		buf.WriteString(`</w:r>`)
	}

	// Write text if any
	if text != "" {
		buf.WriteString(`<w:r><w:t xml:space="preserve">`)
		xml.Escape(&buf, []byte(text))
		buf.WriteString(`</w:t></w:r>`)
	}

	buf.WriteString(`</w:p>`)
	buf.WriteString(`</w:hdr>`)
	_, err := w.Write(buf.Bytes())
	return err
}

// writeHeaderImageDrawing writes an image drawing element for a header.
// Header images are scaled to fit the text area width (page width - margins).
func writeHeaderImageDrawing(buf *bytes.Buffer, rel imageRel, docPrID int, img *imageData) {
	// Header area width: page width (11906 twips) - left margin (1800) - right margin (1800) = 8306 twips
	// 1 twip = 914400/1440 = 635 EMU, so 8306 twips = 8306 * 635 = 5274310 EMU
	maxWidth := int64(5274310)

	cx, cy := int64(3657600), int64(3657600)
	if img != nil {
		emuW, emuH := getImageDimensionsEMU(img)
		if emuW > 0 && emuH > 0 {
			cx, cy = emuW, emuH
			// Scale to fit header area width if wider
			if cx > maxWidth {
				scale := float64(maxWidth) / float64(cx)
				cx = maxWidth
				cy = int64(float64(cy) * scale)
			}
		}
	}

	buf.WriteString(`<w:drawing>`)
	buf.WriteString(`<wp:inline distT="0" distB="0" distL="0" distR="0">`)
	fmt.Fprintf(buf, `<wp:extent cx="%d" cy="%d"/>`, cx, cy)
	fmt.Fprintf(buf, `<wp:docPr id="%d" name="Image %d"/>`, docPrID, docPrID)
	buf.WriteString(`<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">`)
	buf.WriteString(`<pic:pic><pic:nvPicPr>`)
	fmt.Fprintf(buf, `<pic:cNvPr id="%d" name="Image %d"/>`, docPrID, docPrID)
	buf.WriteString(`<pic:cNvPicPr/></pic:nvPicPr>`)
	buf.WriteString(`<pic:blipFill>`)
	fmt.Fprintf(buf, `<a:blip r:embed="%s"/>`, rel.relID)
	buf.WriteString(`<a:stretch><a:fillRect/></a:stretch>`)
	buf.WriteString(`</pic:blipFill>`)
	fmt.Fprintf(buf, `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>`, cx, cy)
	buf.WriteString(`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`)
	buf.WriteString(`</pic:pic></a:graphicData></a:graphic>`)
	buf.WriteString(`</wp:inline>`)
	buf.WriteString(`</w:drawing>`)
}

// writeFooterXML writes a footer XML part with page number field support.
// rawText contains the original text with DOC field codes (0x13/0x14/0x15).
// If rawText is empty, falls back to the cleaned text.
func writeFooterXML(w io.Writer, text string, rawText string) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`)

	// Use raw text if available to preserve field codes
	source := text
	if rawText != "" {
		source = rawText
	}

	// Split by tabs to create separate alignment zones
	parts := strings.Split(source, "\t")
	buf.WriteString(`<w:p>`)
	buf.WriteString(`<w:pPr><w:pStyle w:val="Footer"/></w:pPr>`)
	for i, part := range parts {
		if i > 0 {
			buf.WriteString(`<w:r><w:tab/></w:r>`)
		}
		writeFooterSegment(&buf, part)
	}
	buf.WriteString(`</w:p>`)
	buf.WriteString(`</w:ftr>`)
	_, err := w.Write(buf.Bytes())
	return err
}

// writePartRels writes a relationship file for a header/footer part that references images.
func writePartRels(w io.Writer, imgRels []imageRel) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)
	for _, rel := range imgRels {
		fmt.Fprintf(&buf, `<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/%s"/>`, rel.relID, rel.filename)
	}
	buf.WriteString(`</Relationships>`)
	_, err := w.Write(buf.Bytes())
	return err
}

// writeFooterSegment writes a footer text segment, converting DOC field codes
// (0x13=begin, 0x14=separator, 0x15=end) to OOXML field elements.
func writeFooterSegment(buf *bytes.Buffer, segment string) {
	runes := []rune(segment)
	i := 0
	for i < len(runes) {
		if runes[i] == 0x13 {
			// Field begin - collect instruction and result
			i++
			// Collect field instruction (between 0x13 and 0x14)
			var instruction []rune
			for i < len(runes) && runes[i] != 0x14 && runes[i] != 0x15 {
				instruction = append(instruction, runes[i])
				i++
			}
			// Skip 0x14 separator
			if i < len(runes) && runes[i] == 0x14 {
				i++
			}
			// Skip field result (between 0x14 and 0x15) - DOCX will compute it
			for i < len(runes) && runes[i] != 0x15 {
				i++
			}
			// Skip 0x15 end
			if i < len(runes) && runes[i] == 0x15 {
				i++
			}

			// Convert known field instructions to OOXML
			instrStr := strings.TrimSpace(string(instruction))
			if strings.HasPrefix(instrStr, "PAGE") {
				buf.WriteString(`<w:r><w:fldChar w:fldCharType="begin"/></w:r>`)
				buf.WriteString(`<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>`)
				buf.WriteString(`<w:r><w:fldChar w:fldCharType="separate"/></w:r>`)
				buf.WriteString(`<w:r><w:t>1</w:t></w:r>`)
				buf.WriteString(`<w:r><w:fldChar w:fldCharType="end"/></w:r>`)
			} else if strings.HasPrefix(instrStr, "NUMPAGES") {
				buf.WriteString(`<w:r><w:fldChar w:fldCharType="begin"/></w:r>`)
				buf.WriteString(`<w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>`)
				buf.WriteString(`<w:r><w:fldChar w:fldCharType="separate"/></w:r>`)
				buf.WriteString(`<w:r><w:t>1</w:t></w:r>`)
				buf.WriteString(`<w:r><w:fldChar w:fldCharType="end"/></w:r>`)
			} else {
				// Unknown field - emit as fldSimple
				buf.WriteString(`<w:r><w:fldChar w:fldCharType="begin"/></w:r>`)
				buf.WriteString(`<w:r><w:instrText xml:space="preserve"> `)
				xml.Escape(buf, []byte(instrStr))
				buf.WriteString(` </w:instrText></w:r>`)
				buf.WriteString(`<w:r><w:fldChar w:fldCharType="separate"/></w:r>`)
				buf.WriteString(`<w:r><w:t> </w:t></w:r>`)
				buf.WriteString(`<w:r><w:fldChar w:fldCharType="end"/></w:r>`)
			}
		} else {
			// Regular text - collect until next field or end
			var textRunes []rune
			for i < len(runes) && runes[i] != 0x13 {
				textRunes = append(textRunes, runes[i])
				i++
			}
			if len(textRunes) > 0 {
				buf.WriteString(`<w:r><w:t xml:space="preserve">`)
				xml.Escape(buf, []byte(string(textRunes)))
				buf.WriteString(`</w:t></w:r>`)
			}
		}
	}
}
