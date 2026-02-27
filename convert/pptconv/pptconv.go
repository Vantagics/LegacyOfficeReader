package pptconv

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/common"
	"github.com/shakinm/xlsReader/ppt"
)

// slideData is an internal representation of a slide for PPTX generation.
type slideData struct {
	Texts []string
}

// imageData is an internal representation of an image for PPTX generation.
type imageData struct {
	Format common.ImageFormat
	Data   []byte
}

// imageRel tracks an image file and its relationship ID inside the PPTX archive.
type imageRel struct {
	filename string
	relID    string
}

// formattedSlideData holds formatted slide content for PPTX generation.
type formattedSlideData struct {
	Shapes            []formattedShape
	Background        formattedBackground
	LayoutIdx         int // index into the layouts slice
	DefaultTextStyles [5]ppt.MasterTextStyle // from master
}

// formattedLayoutData holds a slide layout (from PPT master) for PPTX generation.
type formattedLayoutData struct {
	Shapes     []formattedShape
	Background formattedBackground
	// ColorScheme from the PPT master (8 RGB hex strings)
	ColorScheme []string
	// TitleBgColor is a gradient fill color for the title area (above the connector line).
	// Detected from slides that have white title text with embedded dark background hints.
	// Empty string means no title background fill is needed.
	TitleBgColor string
	// TitleBgBottom is the Y coordinate (EMU) of the bottom of the title background area.
	TitleBgBottom int32
	// WatermarkShapes are shapes (typically logo images) positioned near the bottom
	// of the slide that should render ON TOP of slide content. These are separated
	// from Shapes and written into each slide's shape tree after the slide's own shapes.
	WatermarkShapes []formattedShape
}

// formattedBackground holds background fill info for a slide.
type formattedBackground struct {
	HasBackground bool
	FillColor     string // 6-digit hex RGB
	ImageIdx      int    // -1 = no image
}

type formattedShape struct {
	ShapeType     uint16
	Left, Top     int32
	Width, Height int32
	IsText        bool
	IsImage       bool
	ImageIdx      int
	FillColor     string
	LineColor     string
	NoFill        bool
	NoLine        bool
	LineWidth     int32
	LineDash      int32
	Rotation      int32
	FlipH         bool
	FlipV         bool
	// Text body properties
	TextMarginLeft   int32
	TextMarginTop    int32
	TextMarginRight  int32
	TextMarginBottom int32
	TextAnchor       int32
	TextWordWrap     int32
	FillOpacity      int32
	FillColorRaw     uint32
	LineColorRaw     uint32
	// Image cropping (1/65536 = 100%)
	CropFromTop    int32
	CropFromBottom int32
	CropFromLeft   int32
	CropFromRight  int32
	// Freeform geometry
	GeoVertices  []ppt.GeoVertex
	GeoSegments  []ppt.GeoSegment
	GeoLeft      int32
	GeoTop       int32
	GeoRight     int32
	GeoBottom    int32
	// Line arrow properties
	LineStartArrowHead   int32
	LineEndArrowHead     int32
	LineStartArrowWidth  int32
	LineStartArrowLength int32
	LineEndArrowWidth    int32
	LineEndArrowLength   int32
	TextType             int // from TextHeaderAtom: 0=title, 1=body, 4=other, -1=unknown
	Paragraphs       []formattedSlideParagraph
}

type formattedSlideParagraph struct {
	Alignment   uint8
	IndentLevel uint8
	SpaceBefore int32
	SpaceAfter  int32
	LineSpacing int32
	HasBullet   bool
	BulletChar  string
	BulletColor string
	BulletSize  int16
	BulletFont  string
	LeftMargin  int32
	Indent      int32
	Runs        []formattedSlideRun
}

type formattedSlideRun struct {
	Text      string
	FontName  string
	FontSize  uint16
	Bold      bool
	Italic    bool
	Underline bool
	Color     string
	ColorRaw  uint32 // Raw color value from PPT (for scheme color detection)
	fontSizeExplicit bool // true if font size was explicitly set in PPT (not estimated)
}

// mapSlides extracts slide data from a parsed Presentation.
func mapSlides(p *ppt.Presentation) []slideData {
	pptSlides := p.GetSlides()
	result := make([]slideData, len(pptSlides))
	for i, s := range pptSlides {
		texts := s.GetTexts()
		if texts == nil {
			texts = []string{}
		}
		result[i] = slideData{Texts: make([]string, len(texts))}
		copy(result[i].Texts, texts)
	}
	return result
}

// mapImages extracts image data from a parsed Presentation.
func mapImages(p *ppt.Presentation) []imageData {
	pptImages := p.GetImages()
	result := make([]imageData, len(pptImages))
	for i, img := range pptImages {
		data := make([]byte, len(img.Data))
		copy(data, img.Data)
		result[i] = imageData{Format: img.Format, Data: data}
	}
	return result
}

// mapFormattedSlides extracts formatted slide data from a Presentation.
func mapFormattedSlides(p *ppt.Presentation) ([]formattedSlideData, []formattedLayoutData) {
	pptSlides := p.GetSlides()
	masters := p.GetMasters()

	// Build unique layout list from master refs
	var layouts []formattedLayoutData
	masterRefToLayoutIdx := make(map[uint32]int)

	for _, s := range pptSlides {
		ref := s.GetMasterRef()
		if _, ok := masterRefToLayoutIdx[ref]; ok {
			continue
		}
		idx := len(layouts)
		masterRefToLayoutIdx[ref] = idx

		layout := formattedLayoutData{}
		if m, ok := masters[ref]; ok {
			layout.ColorScheme = m.ColorScheme
			layout.Background = formattedBackground{
				HasBackground: m.Background.HasBackground,
				FillColor:     m.Background.FillColor,
				ImageIdx:      m.Background.ImageIdx,
			}
			layout.Shapes = mapShapesToFormatted(m.Shapes)
			// Resolve inherited text props for layout shapes (no master defaults)
			for si := range layout.Shapes {
				resolveInheritedTextProps(&layout.Shapes[si], layout.ColorScheme)
			}
		}
		layouts = append(layouts, layout)
	}

	// If no layouts found, create a default one
	if len(layouts) == 0 {
		layouts = append(layouts, formattedLayoutData{})
	}

	result := make([]formattedSlideData, len(pptSlides))
	for i, s := range pptSlides {
		shapes := s.GetShapes()
		bg := s.GetBackground()
		fShapes := mapShapesToFormatted(shapes)
		layoutIdx := 0
		if idx, ok := masterRefToLayoutIdx[s.GetMasterRef()]; ok {
			layoutIdx = idx
		}
		defaultStyles := s.GetDefaultTextStyles()
		// Get color scheme from the layout for this slide
		var slideColorScheme []string
		if layoutIdx < len(layouts) {
			slideColorScheme = layouts[layoutIdx].ColorScheme
		}
		// Determine if the layout/master has a dark background
		masterHasDarkBg := false
		if layoutIdx < len(layouts) {
			lb := layouts[layoutIdx].Background
			if lb.ImageIdx >= 0 {
				masterHasDarkBg = true // background image is assumed dark
			} else if lb.FillColor != "" && isDarkFillColor(lb.FillColor) {
				masterHasDarkBg = true
			}
		}
		// Collect layout image shapes for overlap checking
		var layoutImageShapes []formattedShape
		var connectorY int32 // Y position of horizontal connector line (title/content boundary)
		if layoutIdx < len(layouts) {
			for _, ls := range layouts[layoutIdx].Shapes {
				if ls.IsImage && ls.ImageIdx >= 0 {
					layoutImageShapes = append(layoutImageShapes, ls)
				}
				// Detect horizontal connector line spanning most of slide width
				if isConnectorShape(ls.ShapeType) && ls.Height == 0 && ls.Width > 6000000 {
					connectorY = ls.Top
				}
			}
		}
		// Apply master default text styles to runs with missing font sizes,
		// then resolve remaining inherited props (sibling inheritance, estimation, dark fill text).
		applyMasterTextDefaults(fShapes, defaultStyles)
		// Apply text-type-specific master defaults (color, font size) from TextHeaderAtom.
		// This uses the per-text-type TextMasterStyleAtom data to provide more accurate
		// defaults than the generic body text defaults.
		textTypeStyles := s.GetTextTypeStyles()
		if textTypeStyles != nil {
			applyTextTypeDefaults(fShapes, textTypeStyles, slideColorScheme)
		}
		// Pre-compute: for each transparent text shape, check if an earlier
		// slide-level shape with a colored (non-near-white) fill overlaps it (z-order bg).
		// This covers both dark fills AND distinctly colored fills (e.g., BDD7EE light blue,
		// 003296 dark blue) where white text should be preserved.
		slideBgDark := make([]bool, len(fShapes))
		for si := range fShapes {
			sh := &fShapes[si]
			if len(sh.Paragraphs) == 0 {
				continue
			}
			hasOwnFill := sh.FillColor != "" && !sh.NoFill
			if hasOwnFill {
				continue
			}
			// Check earlier shapes for colored fill overlap
			shCX := int64(sh.Left) + int64(sh.Width)/2
			shCY := int64(sh.Top) + int64(sh.Height)/2
			for bi := 0; bi < si; bi++ {
				bg := &fShapes[bi]
				if bg.NoFill || bg.FillColor == "" {
					continue
				}
				if bg.FillOpacity >= 0 && bg.FillOpacity < 13107 {
					continue // very transparent
				}
				// Skip near-white fills — they don't provide a colored background
				if isNearWhite(bg.FillColor) {
					continue
				}
				bgR := int64(bg.Left) + int64(bg.Width)
				bgB := int64(bg.Top) + int64(bg.Height)
				if shCX >= int64(bg.Left) && shCX <= bgR && shCY >= int64(bg.Top) && shCY <= bgB {
					slideBgDark[si] = true
					break
				}
			}
		}
		for si := range fShapes {
			resolveInheritedTextPropsWithBg(&fShapes[si], slideColorScheme, masterHasDarkBg || slideBgDark[si], layoutImageShapes, connectorY)
		}
		result[i] = formattedSlideData{
			Shapes: fShapes,
			Background: formattedBackground{
				HasBackground: bg.HasBackground,
				FillColor:     bg.FillColor,
				ImageIdx:      bg.ImageIdx,
			},
			LayoutIdx:         layoutIdx,
			DefaultTextStyles: defaultStyles,
		}
	}
	return result, layouts
}

// detectTitleBackgrounds scans slides for each layout to detect if they need
// a gradient title background fill. This handles PPT masters with gradient fills
// in the title area that we can't parse directly.
func detectTitleBackgrounds(slides []formattedSlideData, layouts []formattedLayoutData, slideW int32) {
	for li := range layouts {
		layout := &layouts[li]

		// Find the horizontal connector line in this layout (marks title/content boundary)
		var connectorY int32 = -1
		for _, sh := range layout.Shapes {
			if isConnectorShape(sh.ShapeType) && sh.Height == 0 && sh.Width > slideW/2 {
				// Horizontal line spanning most of the slide width
				connectorY = sh.Top
				break
			}
		}
		if connectorY <= 0 {
			continue // no connector line found
		}

		// Check slides using this layout for white title text with dark bg hints
		var titleBgColor string
		for _, slide := range slides {
			if slide.LayoutIdx != li {
				continue
			}
			for _, sh := range slide.Shapes {
				if len(sh.Paragraphs) == 0 {
					continue
				}
				// Title shape: above the connector line, no fill, has white bold text
				shapeCenterY := sh.Top + sh.Height/2
				if shapeCenterY >= connectorY {
					continue // below the line
				}
				if sh.FillColor != "" && !sh.NoFill {
					continue // has its own fill
				}
				for _, para := range sh.Paragraphs {
					for _, run := range para.Runs {
						if run.Color != "FFFFFF" || !run.Bold || run.FontSize < 2400 {
							continue
						}
						// Check for embedded dark background hint
						if run.ColorRaw&0xFF000000 == 0xFE000000 {
							embG := uint8((run.ColorRaw >> 8) & 0xFF)
							embB := uint8((run.ColorRaw >> 16) & 0xFF)
							if embG != 0 || embB != 0 {
								// Extract the embedded background color
								embR := uint8(run.ColorRaw & 0xFF)
								titleBgColor = fmt.Sprintf("%02X%02X%02X", embR, embG, embB)
							}
						}
						if titleBgColor != "" {
							break
						}
					}
					if titleBgColor != "" {
						break
					}
				}
				if titleBgColor != "" {
					break
				}
			}
			if titleBgColor != "" {
				break
			}
		}

		if titleBgColor != "" {
			layout.TitleBgColor = titleBgColor
			layout.TitleBgBottom = connectorY
		}
	}
}
func mapShapesToFormatted(shapes []ppt.ShapeFormatting) []formattedShape {
	fShapes := make([]formattedShape, len(shapes))
	for j, sh := range shapes {
		fShapes[j] = formattedShape{
			ShapeType: sh.ShapeType,
			Left:      sh.Left,
			Top:       sh.Top,
			Width:     sh.Width,
			Height:    sh.Height,
			IsText:    sh.IsText,
			IsImage:   sh.IsImage,
			ImageIdx:  sh.ImageIdx,
			FillColor: sh.FillColor,
			LineColor: sh.LineColor,
			NoFill:    sh.NoFill,
			NoLine:    sh.NoLine,
			LineWidth: sh.LineWidth,
			LineDash:  sh.LineDash,
			Rotation:  sh.Rotation,
			FlipH:     sh.FlipH,
			FlipV:     sh.FlipV,
			TextMarginLeft:   sh.TextMarginLeft,
			TextMarginTop:    sh.TextMarginTop,
			TextMarginRight:  sh.TextMarginRight,
			TextMarginBottom: sh.TextMarginBottom,
			TextAnchor:       sh.TextAnchor,
			TextWordWrap:     sh.TextWordWrap,
			FillOpacity:      sh.FillOpacity,
			FillColorRaw:     sh.FillColorRaw,
			LineColorRaw:     sh.LineColorRaw,
			CropFromTop:      sh.CropFromTop,
			CropFromBottom:   sh.CropFromBottom,
			CropFromLeft:     sh.CropFromLeft,
			CropFromRight:    sh.CropFromRight,
			GeoVertices:      sh.GeoVertices,
			GeoSegments:      sh.GeoSegments,
			GeoLeft:          sh.GeoLeft,
			GeoTop:           sh.GeoTop,
			GeoRight:         sh.GeoRight,
			GeoBottom:        sh.GeoBottom,
			LineStartArrowHead:   sh.LineStartArrowHead,
			LineEndArrowHead:     sh.LineEndArrowHead,
			LineStartArrowWidth:  sh.LineStartArrowWidth,
			LineStartArrowLength: sh.LineStartArrowLength,
			LineEndArrowWidth:    sh.LineEndArrowWidth,
			LineEndArrowLength:   sh.LineEndArrowLength,
			TextType:             sh.TextType,
		}
		fParas := make([]formattedSlideParagraph, len(sh.Paragraphs))
		for k, p := range sh.Paragraphs {
			fParas[k] = formattedSlideParagraph{
				Alignment:   p.Alignment,
				IndentLevel: p.IndentLevel,
				SpaceBefore: p.SpaceBefore,
				SpaceAfter:  p.SpaceAfter,
				LineSpacing: p.LineSpacing,
				HasBullet:   p.HasBullet,
				BulletChar:  p.BulletChar,
				BulletColor: p.BulletColor,
				BulletSize:  p.BulletSize,
				BulletFont:  p.BulletFont,
				LeftMargin:  p.LeftMargin,
				Indent:      p.Indent,
			}
			fRuns := make([]formattedSlideRun, len(p.Runs))
			for l, r := range p.Runs {
				fRuns[l] = formattedSlideRun{
					Text:             r.Text,
					FontName:         r.FontName,
					FontSize:         r.FontSize,
					Bold:             r.Bold,
					Italic:           r.Italic,
					Underline:        r.Underline,
					Color:            r.Color,
					ColorRaw:         r.ColorRaw,
					fontSizeExplicit: r.FontSize > 0,
				}
			}
			fParas[k].Runs = fRuns
		}
		fShapes[j].Paragraphs = fParas
	}
	return fShapes
}

// resolveInheritedTextProps fills in missing font sizes, font names, and colors
// for text runs by inheriting from sibling runs within the same shape, and
// ensures text is visible on dark-filled shapes.
// colorScheme is the master's color scheme (8 RGB hex strings) used to resolve
// the default text color when a run has no explicit color.
func resolveInheritedTextProps(shape *formattedShape, colorScheme []string) {
	resolveInheritedTextPropsWithBg(shape, colorScheme, false, nil, 0)
}

// resolveInheritedTextPropsWithBg resolves inherited text properties.
// masterHasDarkBg indicates whether the slide's layout/master has a dark
// background (image or dark solid fill), which affects default text color
// for shapes with no fill.
// layoutImages contains image shapes from the layout that may overlap with
// this shape, providing a visual background behind transparent shapes.
// titleBgBottom is the Y coordinate of the title gradient bottom (0 = no title bg).
func resolveInheritedTextPropsWithBg(shape *formattedShape, colorScheme []string, masterHasDarkBg bool, layoutImages []formattedShape, titleBgBottom int32) {
	if len(shape.Paragraphs) == 0 {
		return
	}

	// Step 1: Find the first explicit font size and font name in the shape
	var defaultFontSize uint16
	var defaultFontName string
	for _, para := range shape.Paragraphs {
		for _, run := range para.Runs {
			if run.FontSize > 0 && defaultFontSize == 0 {
				defaultFontSize = run.FontSize
			}
			if run.FontName != "" && defaultFontName == "" {
				defaultFontName = run.FontName
			}
			if defaultFontSize > 0 && defaultFontName != "" {
				break
			}
		}
		if defaultFontSize > 0 && defaultFontName != "" {
			break
		}
	}

	// If no explicit font size found anywhere in the shape, estimate from shape height
	if defaultFontSize == 0 {
		defaultFontSize = estimateDefaultFontSize(shape)
	}

	// Default font name fallback: use "微软雅黑" as the presentation default
	if defaultFontName == "" {
		defaultFontName = "微软雅黑"
	}

	// Step 2: Resolve fontSize=0 and empty fontName runs
	for pi := range shape.Paragraphs {
		para := &shape.Paragraphs[pi]
		// Find first explicit size in this paragraph
		var paraFontSize uint16
		for _, run := range para.Runs {
			if run.FontSize > 0 {
				paraFontSize = run.FontSize
				break
			}
		}
		inheritSize := paraFontSize
		if inheritSize == 0 {
			inheritSize = defaultFontSize
		}
		// Apply inherited size to runs with fontSize=0
		if inheritSize > 0 {
			for ri := range para.Runs {
				if para.Runs[ri].FontSize == 0 {
					para.Runs[ri].FontSize = inheritSize
				}
			}
		}
		// Apply inherited font name to runs with empty font name
		for ri := range para.Runs {
			if para.Runs[ri].FontName == "" {
				para.Runs[ri].FontName = defaultFontName
			}
		}
	}

	// Step 3: Determine the effective background color for this shape.
	shapeFillColor := shape.FillColor
	hasOpaqueColoredFill := shapeFillColor != "" && !shape.NoFill
	if hasOpaqueColoredFill && shape.FillOpacity >= 0 && shape.FillOpacity < 13107 {
		// Very transparent fill (<20% opacity) - treat as no fill for color decisions
		hasOpaqueColoredFill = false
	}
	shapeFillIsDark := hasOpaqueColoredFill && isDarkFillColor(shapeFillColor)
	shapeIsTransparent := !hasOpaqueColoredFill

	// Check if a layout image overlaps with this shape's position.
	// If so, the shape effectively sits on an image background (assumed dark/colored),
	// and we should preserve the original text color.
	// Skip decorative/watermark images and logos that don't provide a real background.
	layoutImageBehind := false
	if shapeIsTransparent && len(layoutImages) > 0 {
		for _, img := range layoutImages {
			// Skip small images (logos) — they don't provide a background
			if img.Width < 3000000 || img.Height < 2000000 {
				continue
			}
			// Skip decorative/watermark images: images that are horizontally offset
			// from the left edge are partial overlays (watermarks), not backgrounds.
			// A real background image starts at or near x=0.
			// Images starting at x=0 with a y offset are still valid backgrounds
			// (e.g., an image covering the lower 80% of the slide).
			if img.Left > 500000 {
				continue
			}
			// Check if the image overlaps with the shape's bounding box
			imgRight := int64(img.Left) + int64(img.Width)
			imgBottom := int64(img.Top) + int64(img.Height)
			// Overlap if the image covers at least the center of the shape
			shapeCenterX := int64(shape.Left) + int64(shape.Width)/2
			shapeCenterY := int64(shape.Top) + int64(shape.Height)/2
			if shapeCenterX >= int64(img.Left) && shapeCenterX <= imgRight &&
				shapeCenterY >= int64(img.Top) && shapeCenterY <= imgBottom {
				layoutImageBehind = true
				break
			}
		}
	}

	// Default text color from scheme
	defaultTextColor := ""
	if len(colorScheme) >= 2 {
		defaultTextColor = colorScheme[1] // dk1 = text/lines color
	}

	// Background color from scheme (index 0)
	bgSchemeColor := ""
	if len(colorScheme) >= 1 {
		bgSchemeColor = colorScheme[0] // lt1 = background color
	}

	// Determine if the shape fill is "distinctly colored" (not white/near-white).
	// When a shape has a distinct colored fill (yellow, blue, orange, etc.),
	// white text on it is intentional and should be preserved.
	fillIsDistinctColor := false
	if hasOpaqueColoredFill {
		fillIsDistinctColor = shapeFillColor != bgSchemeColor && !isNearWhite(shapeFillColor)
	}

	// Step 4: Fix text colors to ensure visibility.
	for pi := range shape.Paragraphs {
		for ri := range shape.Paragraphs[pi].Runs {
			run := &shape.Paragraphs[pi].Runs[ri]

			// Detect if this color came from a scheme reference.
			// Per [MS-PPT] ColorIndexStruct, the high byte (index field) determines:
			//   0x00: scheme index 0 (background) — but 0x00000000 is treated as "no color"
			//   0x01-0x07: scheme color index (the high byte IS the index)
			//   0x08: scheme color with index in low byte (legacy)
			//   0xFE: direct sRGB or scheme[0] reference with embedded bg hint
			colorFlag := run.ColorRaw >> 24
			isSchemeColor := run.ColorRaw != 0 && (colorFlag == 0xFE || colorFlag == 0x08)
			schemeIdx := int(run.ColorRaw & 0xFF)
			// For 0xFE colors, the low bytes are RGB data (or cached bg hint),
			// NOT a scheme index. 0xFE always references scheme[0] (background).
			if colorFlag == 0xFE {
				// 0xFE means direct sRGB color. The low 3 bytes are R, G, B.
				// Special case: 0xFE000000 (all-zero RGB) is a scheme[0] reference,
				// already resolved by ResolveSchemeColor.
				// Non-zero low bytes: the parsed color IS the intended text color.
				// Do NOT override it — the earlier assumption that non-zero G/B
				// bytes encode a "cached background hint" was incorrect and caused
				// direct sRGB colors (e.g., 003296 dark blue) to be replaced with
				// scheme[0] (white), making title text invisible on light backgrounds
				// and triggering false title-background detection.
				if run.ColorRaw&0x00FFFFFF == 0 {
					schemeIdx = 0
				} else {
					// Direct sRGB color — not a scheme reference.
					// Skip the scheme[0] special-case logic below.
					isSchemeColor = false
				}
			}

			// 0x01-0x07 are direct scheme index references where the high byte IS the index.
			// These have already been resolved to the correct scheme color by ResolveSchemeColor,
			// so they are treated as direct colors and skip the scheme[0] special case below.

			// Treat colorRaw=0x00000000 as "no explicit color" — the PPT binary
			// sets the color flag with all zeros as a default/inherit marker.
			noExplicitColor := run.Color == "" || run.ColorRaw == 0x00000000

			if noExplicitColor {
				// No color set: use dk1 for light backgrounds, white for dark
				shapeCenterY := shape.Top + shape.Height/2
				inTitleArea := titleBgBottom > 0 && shapeCenterY < titleBgBottom
				if shapeFillIsDark || (masterHasDarkBg && shapeIsTransparent) || (layoutImageBehind && shapeIsTransparent) || (inTitleArea && shapeIsTransparent) {
					run.Color = "FFFFFF"
				} else if defaultTextColor != "" {
					run.Color = defaultTextColor
				}
				continue
			}

			// Scheme color index 0 (background/lt1) used for text.
			// PPT uses this as a "default text color" marker in many cases.
			// This only triggers for 0xFE000000 (scheme[0] reference with zero RGB).
			// - On transparent shapes or white/near-white fills: replace with dk1
			// - On distinctly colored fills (yellow, blue, etc.): keep as-is (white text is intentional)
			// - On dark fills: keep as white
			if isSchemeColor && schemeIdx == 0 && bgSchemeColor != "" && run.Color == bgSchemeColor {
				// This block handles 0xFE000000 (scheme[0] reference) and 0x08 scheme
				// references that resolved to scheme[0]. The text color is white/lt1.
				if shapeFillIsDark || fillIsDistinctColor {
					// Keep the resolved color (usually FFFFFF) - white text on colored/dark fill is intentional
				} else if hasOpaqueColoredFill && isNearWhite(shapeFillColor) {
					// Near-white fill (E9EBF5, CFD5EA, etc.): white text would be invisible.
					// Convert to dk1 regardless of scheme[0] reference.
					if defaultTextColor != "" {
						run.Color = defaultTextColor
					} else {
						run.Color = "000000"
					}
				} else if shapeIsTransparent && layoutImageBehind {
					// Transparent shape on layout image: keep white (image provides dark/colored bg)
				} else if shapeIsTransparent && !masterHasDarkBg {
					// Transparent shape on light master bg: use dk1
					if defaultTextColor != "" {
						run.Color = defaultTextColor
					} else {
						run.Color = "000000"
					}
				} else if shapeIsTransparent && masterHasDarkBg {
					// Transparent shape on dark master bg: keep white
				} else {
					// White/near-white fill: use dk1
					if defaultTextColor != "" {
						run.Color = defaultTextColor
					} else {
						run.Color = "000000"
					}
				}
				continue
			}

			// Check if the text color would be invisible against the shape fill
			textIsDark := isDarkFillColor(run.Color)

			if shapeFillIsDark && textIsDark && colorLowContrast(run.Color, shapeFillColor) {
				// Dark text on dark fill with low contrast - make text white
				run.Color = "FFFFFF"
			} else if hasOpaqueColoredFill && !shapeFillIsDark && run.Color == shapeFillColor {
				// Text color exactly matches fill color - use dk1
				if defaultTextColor != "" {
					run.Color = defaultTextColor
				} else {
					run.Color = "000000"
				}
			} else if shapeIsTransparent && (masterHasDarkBg || layoutImageBehind) && textIsDark {
				// Transparent shape on dark background with dark text.
				// Only convert if the text color is actually black/near-black (low contrast
				// with the background). Distinctive dark colors like red (FF0000) should
				// be preserved as they're intentional and visible on dark backgrounds.
				if run.Color == "000000" || run.ColorRaw == 0x00000000 {
					run.Color = "FFFFFF"
				}
			}

			// Final safety check: white/near-white text on transparent shape
			// over light background with no layout image = invisible text.
			// This catches cases like 0xFEFFFFFF (direct white) and other
			// white colors that weren't caught by the scheme[0] logic above.
			if isNearWhite(run.Color) && shapeIsTransparent && !masterHasDarkBg && !layoutImageBehind {
				// Check if the shape is in the title gradient area
				shapeCenterY := shape.Top + shape.Height/2
				inTitleArea := titleBgBottom > 0 && shapeCenterY < titleBgBottom
				if !inTitleArea {
					if defaultTextColor != "" {
						run.Color = defaultTextColor
					} else {
						run.Color = "000000"
					}
				}
			}

			// Safety check: white/near-white text on near-white fill = invisible.
			// This catches non-scheme white text on light fills like E9EBF5, CFD5EA, etc.
			// BUT: skip this check when the text color was explicitly set via scheme[0]
			// reference — the PPT author intentionally chose white text on these fills.
			isExplicitScheme0 := isSchemeColor && schemeIdx == 0 && bgSchemeColor != "" && run.Color == bgSchemeColor
			if isNearWhite(run.Color) && hasOpaqueColoredFill && !shapeFillIsDark && isNearWhite(shapeFillColor) && !isExplicitScheme0 {
				if defaultTextColor != "" {
					run.Color = defaultTextColor
				} else {
					run.Color = "000000"
				}
			}
		}
	}
}

// estimateDefaultFontSize estimates a reasonable font size for a shape with no
// explicit font size, based on the shape's dimensions, text length, and paragraph count.
// Returns size in centipoints (hundredths of a point).
func estimateDefaultFontSize(shape *formattedShape) uint16 {
	paraCount := len(shape.Paragraphs)
	if paraCount == 0 {
		return 1400
	}

	heightEMU := int64(shape.Height)
	widthEMU := int64(shape.Width)
	if heightEMU <= 0 {
		return 1400
	}

	// Account for text body margins
	// For small shapes (height < 500000 EMU ≈ 0.55"), use minimal margins
	// since PPT auto-shrinks text to fit.
	marginTop := int64(45720)
	marginBottom := int64(45720)
	marginLeft := int64(91440)
	marginRight := int64(91440)
	if heightEMU < 500000 {
		// Small shape: use minimal margins
		marginTop = int64(12700)
		marginBottom = int64(12700)
		marginLeft = int64(25400)
		marginRight = int64(25400)
	}
	if shape.TextMarginTop >= 0 {
		marginTop = int64(shape.TextMarginTop)
	}
	if shape.TextMarginBottom >= 0 {
		marginBottom = int64(shape.TextMarginBottom)
	}
	if shape.TextMarginLeft >= 0 {
		marginLeft = int64(shape.TextMarginLeft)
	}
	if shape.TextMarginRight >= 0 {
		marginRight = int64(shape.TextMarginRight)
	}
	availHeight := heightEMU - marginTop - marginBottom
	availWidth := widthEMU - marginLeft - marginRight
	if availHeight <= 0 {
		availHeight = heightEMU
	}
	if availWidth <= 0 {
		availWidth = widthEMU
	}

	// Try candidate font sizes from large to small, pick the largest that fits.
	// Cap the maximum candidate based on shape characteristics:
	// - Small shapes (height < 500000 EMU) use smaller candidates
	// - Shapes with short text (labels) should not get oversized fonts
	maxCandidate := uint16(2000)
	totalChars := 0
	for _, para := range shape.Paragraphs {
		for _, run := range para.Runs {
			totalChars += len([]rune(run.Text))
		}
	}
	// For short text labels (< 15 chars), cap at 1400 to avoid oversized labels
	if totalChars > 0 && totalChars < 15 && paraCount <= 1 {
		maxCandidate = 1400
	}
	// For medium text (15-40 chars), cap at 1600 — but only for multi-paragraph
	// shapes. Single-paragraph banners/headers should be allowed to use larger sizes.
	if totalChars >= 15 && totalChars <= 40 && paraCount >= 2 {
		if maxCandidate > 1600 {
			maxCandidate = 1600
		}
	}
	// For shapes with own fill color (content boxes, banners, callouts) and
	// substantial text (>40 chars), cap at 1400. These are typically body-text
	// content areas where large fonts cause overflow.
	hasOwnFill := shape.FillColor != "" && !shape.NoFill
	if hasOwnFill && totalChars > 40 {
		if maxCandidate > 1400 {
			maxCandidate = 1400
		}
	}
	// For very wide shapes (aspect ratio > 4:1), the estimation algorithm tends
	// to pick oversized fonts because text wraps into few lines on wide shapes.
	// Cap more aggressively based on aspect ratio and text length.
	// Exception: single-paragraph shapes with short text (≤40 chars) are likely
	// banners/headers where the text won't wrap regardless — skip the cap.
	if heightEMU > 0 && widthEMU > heightEMU*4 && totalChars > 20 {
		if paraCount > 1 || totalChars > 40 {
			if maxCandidate > 1400 {
				maxCandidate = 1400
			}
		}
	}
	allCandidates := []uint16{2000, 1800, 1600, 1400, 1200, 1100, 1000, 900, 800, 700, 600}
	var candidates []uint16
	for _, c := range allCandidates {
		if c <= maxCandidate {
			candidates = append(candidates, c)
		}
	}
	for _, szCp := range candidates {
		szPt := float64(szCp) / 100.0 // points
		szEMU := szPt * 12700.0        // EMU per character height
		baseLineHeight := szEMU * 1.2  // default line height with minimal spacing

		// Determine effective line spacing multiplier from paragraph properties.
		// PPT lineSpacing > 0 is a percentage (e.g., 200 = 200% = double spacing).
		// If paragraphs have explicit line spacing, use the maximum across paragraphs.
		lineSpacingMult := 1.0
		for _, para := range shape.Paragraphs {
			if para.LineSpacing > 0 {
				mult := float64(para.LineSpacing) / 100.0
				if mult > lineSpacingMult {
					lineSpacingMult = mult
				}
			}
		}
		lineHeightEMU := baseLineHeight
		if lineSpacingMult > 1.0 {
			lineHeightEMU = szEMU * lineSpacingMult
		}

		// Estimate characters per line: CJK char width ≈ font size, Latin ≈ 0.6×
		charWidthEMU := szEMU * 0.85
		if availWidth > 0 && charWidthEMU > 0 {
			charsPerLine := float64(availWidth) / charWidthEMU
			if charsPerLine < 1 {
				charsPerLine = 1
			}
			// Estimate visual line count from text wrapping
			visualLines := 0
			for _, para := range shape.Paragraphs {
				paraChars := 0
				for _, run := range para.Runs {
					for _, r := range run.Text {
						if r == '\v' || r == '\n' {
							visualLines++
							paraChars = 0
						} else {
							paraChars++
						}
					}
				}
				if paraChars > 0 {
					wrapLines := (paraChars + int(charsPerLine) - 1) / int(charsPerLine)
					visualLines += wrapLines
				} else {
					visualLines++
				}
			}
			if visualLines == 0 {
				visualLines = 1
			}

			totalHeight := float64(visualLines) * lineHeightEMU
			if totalHeight <= float64(availHeight) {
				return szCp
			}
		} else {
			// Fallback: just use explicit line count
			explicitLines := paraCount
			totalHeight := float64(explicitLines) * lineHeightEMU
			if totalHeight <= float64(availHeight) {
				return szCp
			}
		}
	}
	return 600
}

// isDarkFillColor returns true if the hex RGB color is dark (low luminance).
func isDarkFillColor(hex string) bool {
	if len(hex) != 6 {
		return false
	}
	r := hexDigit(hex[0])*16 + hexDigit(hex[1])
	g := hexDigit(hex[2])*16 + hexDigit(hex[3])
	b := hexDigit(hex[4])*16 + hexDigit(hex[5])
	lum := 0.299*float64(r) + 0.587*float64(g) + 0.114*float64(b)
	return lum < 128
}

// isNearWhite returns true if the color is very close to white - high luminance
// AND low saturation. This distinguishes near-white fills (like CFD5EA light purple)
// from distinctly colored fills (like FFD966 yellow) where white text is intentional.
func isNearWhite(hex string) bool {
	if len(hex) != 6 {
		return false
	}
	r := hexDigit(hex[0])*16 + hexDigit(hex[1])
	g := hexDigit(hex[2])*16 + hexDigit(hex[3])
	b := hexDigit(hex[4])*16 + hexDigit(hex[5])
	lum := 0.299*float64(r) + 0.587*float64(g) + 0.114*float64(b)
	// Very high luminance (>230) is always near-white
	if lum > 230 {
		return true
	}
	// For luminance 180-230, check saturation
	if lum > 180 {
		maxC := r
		if g > maxC { maxC = g }
		if b > maxC { maxC = b }
		minC := r
		if g < minC { minC = g }
		if b < minC { minC = b }
		spread := maxC - minC
		// Low color spread means desaturated (grayish/whitish)
		// CFD5EA: spread=27, FFD966: spread=153, C5E0B4: spread=44
		if spread < 80 {
			return true
		}
	}
	return false
}


// colorLowContrast returns true if two hex colors are too similar to be
// distinguishable (e.g., black text on dark blue fill). Distinctive colors
// like red on blue have enough contrast and return false.
func colorLowContrast(hex1, hex2 string) bool {
	if len(hex1) != 6 || len(hex2) != 6 {
		return true
	}
	r1 := hexDigit(hex1[0])*16 + hexDigit(hex1[1])
	g1 := hexDigit(hex1[2])*16 + hexDigit(hex1[3])
	b1 := hexDigit(hex1[4])*16 + hexDigit(hex1[5])
	r2 := hexDigit(hex2[0])*16 + hexDigit(hex2[1])
	g2 := hexDigit(hex2[2])*16 + hexDigit(hex2[3])
	b2 := hexDigit(hex2[4])*16 + hexDigit(hex2[5])
	dr := r1 - r2
	dg := g1 - g2
	db := b1 - b2
	dist := dr*dr + dg*dg + db*db
	return dist < 15000
}

func hexDigit(c byte) int {
	switch {
	case c >= '0' && c <= '9':
		return int(c - '0')
	case c >= 'A' && c <= 'F':
		return int(c-'A') + 10
	case c >= 'a' && c <= 'f':
		return int(c-'a') + 10
	}
	return 0
}

// applyMasterTextDefaults applies master default text styles to shapes that have
// runs with non-explicit font sizes (estimated by resolveInheritedTextProps).
// Master defaults are more accurate than shape-dimension-based estimation.
// Only applies to shapes where NO run has an explicit font size (i.e., the entire
// shape inherited its font size from the master in the original PPT).
func applyMasterTextDefaults(shapes []formattedShape, defaults [5]ppt.MasterTextStyle) {
	// Check if we have any useful defaults
	hasDefaults := false
	for _, d := range defaults {
		if d.FontSize > 0 {
			hasDefaults = true
			break
		}
	}
	if !hasDefaults {
		return
	}

	for si := range shapes {
		shape := &shapes[si]

		// Skip shapes that are unlikely to be body text placeholders.
		// Shapes with their own fill color (non-transparent) are typically
		// content shapes (banners, labels, callouts) that should use
		// estimateDefaultFontSize instead of master defaults.
		// Also skip textbox shapes (type 202) which are free-form text boxes.
		hasOwnFill := shape.FillColor != "" && !shape.NoFill
		isTextBox := shape.ShapeType == 202
		if hasOwnFill || isTextBox {
			// For filled single-paragraph banner shapes with short text,
			// also apply master default font size (not just font name).
			// These are typically header banners where the master default
			// is more appropriate than estimateDefaultFontSize.
			isBanner := hasOwnFill && !isTextBox && len(shape.Paragraphs) == 1
			bannerChars := 0
			if isBanner {
				for _, run := range shape.Paragraphs[0].Runs {
					bannerChars += len([]rune(run.Text))
				}
				// Only treat as banner if text is short enough (≤40 chars)
				if bannerChars > 40 {
					isBanner = false
				}
			}

			// Still apply font name defaults (and font size for banners)
			for pi := range shape.Paragraphs {
				para := &shape.Paragraphs[pi]
				level := int(para.IndentLevel)
				if level > 4 {
					level = 4
				}
				def := defaults[level]
				if def.FontName == "" && level > 0 {
					def = defaults[0]
				}
				if def.FontSize == 0 && level > 0 {
					def = defaults[0]
				}
				for ri := range para.Runs {
					if para.Runs[ri].FontName == "" && def.FontName != "" {
						para.Runs[ri].FontName = def.FontName
					}
					// Apply master font size to banner shapes if it fits
					if isBanner && !para.Runs[ri].fontSizeExplicit && para.Runs[ri].FontSize == 0 && def.FontSize > 0 {
						if masterDefaultFitsShape(shape, def.FontSize) {
							para.Runs[ri].FontSize = def.FontSize
						}
					}
				}
			}
			continue
		}

		// Count total text characters to detect multi-line content shapes.
		// Shapes with substantial text content (long paragraphs) are more likely
		// to be content boxes than title placeholders, so we should be conservative
		// about applying large master defaults.
		totalChars := 0
		for _, para := range shape.Paragraphs {
			for _, run := range para.Runs {
				totalChars += len([]rune(run.Text))
			}
		}

		for pi := range shape.Paragraphs {
			para := &shape.Paragraphs[pi]
			level := int(para.IndentLevel)
			if level > 4 {
				level = 4
			}
			def := defaults[level]
			if def.FontSize == 0 && level > 0 {
				def = defaults[0]
			}
			for ri := range para.Runs {
				run := &para.Runs[ri]
				// Apply master default font size if the run doesn't have an explicit one.
				// Only apply to shapes large enough to be body text placeholders.
				// For small shapes, the estimateDefaultFontSize will handle sizing.
				if !run.fontSizeExplicit && run.FontSize == 0 && def.FontSize > 0 {
					targetSize := def.FontSize
					// For shapes with substantial text content (>40 chars), cap the
					// master default to avoid oversized text in content boxes.
					// These shapes are more likely content areas than title placeholders.
					if totalChars > 40 && targetSize > 1800 {
						targetSize = 1800
					}
					// Estimate visual line count considering text wrapping at the
					// candidate font size, not just paragraph count.
					if masterDefaultFitsShape(shape, targetSize) {
						run.FontSize = targetSize
					}
				}
				if run.FontName == "" && def.FontName != "" {
					run.FontName = def.FontName
				}
			}
		}
	}
}

// applyTextTypeDefaults applies text-type-specific master defaults to shapes.
// When a shape has a TextType from TextHeaderAtom, we look up the corresponding
// TextMasterStyleAtom to get the correct default color and font size for that text type.
// This is especially important for "other" text (type 4) which may have different
// default colors than body text (type 1).
func applyTextTypeDefaults(shapes []formattedShape, textTypeStyles map[int][5]ppt.MasterTextStyle, colorScheme []string) {
	for si := range shapes {
		shape := &shapes[si]
		if shape.TextType < 0 || len(shape.Paragraphs) == 0 {
			continue
		}
		styles, ok := textTypeStyles[shape.TextType]
		if !ok {
			continue
		}
		for pi := range shape.Paragraphs {
			para := &shape.Paragraphs[pi]
			level := int(para.IndentLevel)
			if level > 4 {
				level = 4
			}
			def := styles[level]
			if def.FontSize == 0 && def.Color == "" && level > 0 {
				def = styles[0]
			}
			for ri := range para.Runs {
				run := &para.Runs[ri]
				// Apply text-type-specific default color when the run has no explicit color.
				// ColorRaw==0 means no color was set in the PPT binary.
				if run.ColorRaw == 0 && run.Color == "" && def.Color != "" {
					resolvedColor := def.Color
					if def.ColorRaw != 0 {
						resolvedColor = ppt.ResolveSchemeColor(def.Color, def.ColorRaw, colorScheme)
					}
					run.Color = resolvedColor
					run.ColorRaw = def.ColorRaw
				}
				// Apply text-type-specific font size if not already set
				if !run.fontSizeExplicit && run.FontSize == 0 && def.FontSize > 0 {
					run.FontSize = def.FontSize
				}
				if run.FontName == "" && def.FontName != "" {
					run.FontName = def.FontName
				}
			}
		}
	}
}

// masterDefaultFitsShape checks whether applying the given font size (centipoints)
// to all sz=0 runs in the shape would fit within the shape's height, accounting
// for text wrapping.
func masterDefaultFitsShape(shape *formattedShape, fontSizeCp uint16) bool {
	heightEMU := int64(shape.Height)
	widthEMU := int64(shape.Width)
	if heightEMU <= 0 {
		return false
	}

	// Account for text body margins
	marginTop := int64(45720)
	marginBottom := int64(45720)
	marginLeft := int64(91440)
	marginRight := int64(91440)
	if shape.TextMarginTop >= 0 {
		marginTop = int64(shape.TextMarginTop)
	}
	if shape.TextMarginBottom >= 0 {
		marginBottom = int64(shape.TextMarginBottom)
	}
	if shape.TextMarginLeft >= 0 {
		marginLeft = int64(shape.TextMarginLeft)
	}
	if shape.TextMarginRight >= 0 {
		marginRight = int64(shape.TextMarginRight)
	}
	availHeight := heightEMU - marginTop - marginBottom
	availWidth := widthEMU - marginLeft - marginRight
	if availHeight <= 0 {
		availHeight = heightEMU
	}
	if availWidth <= 0 {
		availWidth = widthEMU
	}

	szPt := float64(fontSizeCp) / 100.0
	szEMU := szPt * 12700.0
	lineHeightEMU := szEMU * 1.2
	charWidthEMU := szEMU * 0.85 // CJK char width ≈ font size, Latin ≈ 0.6×, blend

	charsPerLine := float64(availWidth) / charWidthEMU
	if charsPerLine < 1 {
		charsPerLine = 1
	}

	visualLines := 0
	for _, para := range shape.Paragraphs {
		paraChars := 0
		for _, run := range para.Runs {
			for _, r := range run.Text {
				if r == '\v' || r == '\n' {
					visualLines++
					paraChars = 0
				} else {
					paraChars++
				}
			}
		}
		if paraChars > 0 {
			wrapLines := (paraChars + int(charsPerLine) - 1) / int(charsPerLine)
			visualLines += wrapLines
		} else {
			visualLines++
		}
	}
	if visualLines == 0 {
		visualLines = 1
	}

	totalHeight := float64(visualLines) * lineHeightEMU
	return totalHeight <= float64(availHeight)
}



// hasFormattedContent checks if any slide has non-empty shapes.
func hasFormattedContent(slides []formattedSlideData) bool {
	for _, s := range slides {
		if len(s.Shapes) > 0 {
			return true
		}
	}
	return false
}

// isEmptySlide returns true if a formatted slide has no meaningful content
// (no shapes, only empty text boxes, or only master template placeholders).
func isEmptySlide(slide formattedSlideData) bool {
	if len(slide.Shapes) == 0 {
		return true
	}
	var allTexts []string

	for _, sh := range slide.Shapes {
		for _, para := range sh.Paragraphs {
			for _, run := range para.Runs {
				t := strings.TrimSpace(run.Text)
				if t != "" {
					allTexts = append(allTexts, t)
				}
			}
		}
	}

	// No text at all — this is a template/decoration slide (images, lines only)
	if len(allTexts) == 0 {
		return true
	}

	// Check if ALL texts are trivial (decorative symbols, URLs, master placeholders)
	for _, t := range allTexts {
		if !isTrivialOrPlaceholderText(t) {
			return false
		}
	}
	return true
}

// masterPlaceholderPrefixes are text prefixes that indicate master slide placeholders.
var masterPlaceholderPrefixes = []string{
	"单击此处编辑母版",
	"点击此处编辑母版",
	"Click to edit Master",
}

// masterPlaceholderTexts are exact texts that appear in master slide level indicators.
var masterPlaceholderTexts = []string{
	"二级", "三级", "四级", "五级",
	"第二级", "第三级", "第四级", "第五级",
	"Second level", "Third level", "Fourth level", "Fifth level",
}

// isTrivialOrPlaceholderText returns true if the text is decorative, a placeholder,
// or otherwise not substantive content.
func isTrivialOrPlaceholderText(t string) bool {
	// Master template prefixes
	for _, prefix := range masterPlaceholderPrefixes {
		if strings.Contains(t, prefix) {
			return true
		}
	}
	// Level indicators
	for _, lvl := range masterPlaceholderTexts {
		if t == lvl {
			return true
		}
	}
	// Decorative symbols
	if t == "*" || t == "•" || t == "-" || t == "—" || t == "·" {
		return true
	}
	// Short URLs / footers
	if len(t) <= 30 && (strings.Contains(t, "www.") || strings.Contains(t, ".com") || strings.Contains(t, ".cn") || strings.Contains(t, ".org")) {
		return true
	}
	return false
}

// filterEmptySlides removes slides that have no meaningful content.
func filterEmptySlides(slides []formattedSlideData) []formattedSlideData {
	var result []formattedSlideData
	for _, s := range slides {
		if !isEmptySlide(s) {
			result = append(result, s)
		}
	}
	if result == nil {
		return []formattedSlideData{}
	}
	return result
}

// isEmptyTextSlide returns true if a plain text slide has no content.
func isEmptyTextSlide(slide slideData) bool {
	for _, t := range slide.Texts {
		for _, r := range t {
			if r != ' ' && r != '\t' && r != '\n' && r != '\r' {
				return false
			}
		}
	}
	return true
}

// filterEmptyTextSlides removes plain text slides with no content.
func filterEmptyTextSlides(slides []slideData) []slideData {
	var result []slideData
	for _, s := range slides {
		if !isEmptyTextSlide(s) {
			result = append(result, s)
		}
	}
	if result == nil {
		return []slideData{}
	}
	return result
}

// ConvertReader reads PPT data from reader, converts it to PPTX, and writes to writer.
func ConvertReader(reader io.ReadSeeker, writer io.Writer) error {
	presentation, err := ppt.OpenReader(reader)
	if err != nil {
		return fmt.Errorf("pptconv: failed to parse input: %w", err)
	}

	images := mapImages(&presentation)

	// Try formatted path first
	fmtSlides, layouts := mapFormattedSlides(&presentation)
	if hasFormattedContent(fmtSlides) {
		slideWidth, slideHeight := presentation.GetSlideSize()
		// Detect title backgrounds from gradient fills we can't parse
		detectTitleBackgrounds(fmtSlides, layouts, slideWidth)
		// Separate watermark shapes from layouts so they render on top of slide content
		separateWatermarkShapes(layouts, slideWidth, slideHeight)
		if err := writePptxFormatted(writer, fmtSlides, layouts, images, slideWidth, slideHeight); err != nil {
			return fmt.Errorf("pptconv: failed to write pptx: %w", err)
		}
		return nil
	}

	// Fallback to plain text mode
	slides := mapSlides(&presentation)
	if err := writePptx(writer, slides, images); err != nil {
		return fmt.Errorf("pptconv: failed to write pptx: %w", err)
	}
	return nil
}


// ConvertFile converts a PPT file at inputPath to a PPTX file at outputPath.
func ConvertFile(inputPath string, outputPath string) error {
	inFile, err := os.Open(inputPath)
	if err != nil {
		return fmt.Errorf("pptconv: failed to open input file: %w", err)
	}
	defer inFile.Close()

	outFile, err := os.Create(outputPath)
	if err != nil {
		return fmt.Errorf("pptconv: failed to create output file: %w", err)
	}
	defer outFile.Close()

	return ConvertReader(inFile, outFile)
}

// writePptx generates a minimal valid PPTX (Office Open XML) zip archive.
func writePptx(w io.Writer, slides []slideData, images []imageData) error {
	zw := zip.NewWriter(w)
	defer zw.Close()

	// Write images into ppt/media/ and collect relationship info
	var imgRels []imageRel
	for i, img := range images {
		ext := (&common.Image{Format: img.Format}).Extension()
		if ext == "" {
			ext = ".bin"
		}
		filename := fmt.Sprintf("image%d%s", i+1, ext)
		fw, err := zw.Create("ppt/media/" + filename)
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

	// Write slide XML files
	for i, slide := range slides {
		slideNum := i + 1
		fw, err := zw.Create(fmt.Sprintf("ppt/slides/slide%d.xml", slideNum))
		if err != nil {
			return err
		}
		if err := writeSlideXML(fw, slide); err != nil {
			return err
		}
		fw, err = zw.Create(fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", slideNum))
		if err != nil {
			return err
		}
		if _, err := io.WriteString(fw, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`+
			`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`+
			`</Relationships>`); err != nil {
			return err
		}
	}

	// If no slides, create at least one empty slide for a valid PPTX
	if len(slides) == 0 {
		fw, err := zw.Create("ppt/slides/slide1.xml")
		if err != nil {
			return err
		}
		if err := writeSlideXML(fw, slideData{}); err != nil {
			return err
		}
		fw, err = zw.Create("ppt/slides/_rels/slide1.xml.rels")
		if err != nil {
			return err
		}
		if _, err := io.WriteString(fw, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`+
			`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`+
			`</Relationships>`); err != nil {
			return err
		}
	}

	// Write minimal slideLayout
	fw, err := zw.Create("ppt/slideLayouts/slideLayout1.xml")
	if err != nil {
		return err
	}
	if _, err := io.WriteString(fw, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
		`<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" type="blank">`+
		`<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>`); err != nil {
		return err
	}

	fw, err = zw.Create("ppt/slideLayouts/_rels/slideLayout1.xml.rels")
	if err != nil {
		return err
	}
	if _, err := io.WriteString(fw, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`+
		`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>`+
		`</Relationships>`); err != nil {
		return err
	}

	// Write minimal slideMaster with color map
	fw, err = zw.Create("ppt/slideMasters/slideMaster1.xml")
	if err != nil {
		return err
	}
	if _, err := io.WriteString(fw, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
		`<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`+
		`<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>`+
		`<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>`+
		`<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>`+
		`</p:sldMaster>`); err != nil {
		return err
	}

	fw, err = zw.Create("ppt/slideMasters/_rels/slideMaster1.xml.rels")
	if err != nil {
		return err
	}
	if _, err := io.WriteString(fw, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`+
		`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`+
		`<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>`+
		`</Relationships>`); err != nil {
		return err
	}

	// Write theme
	fw, err = zw.Create("ppt/theme/theme1.xml")
	if err != nil {
		return err
	}
	if _, err := io.WriteString(fw, themeXML); err != nil {
		return err
	}

	// Write presProps, viewProps, tableStyles
	fw, err = zw.Create("ppt/presProps.xml")
	if err != nil {
		return err
	}
	if _, err := io.WriteString(fw, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
		`<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`); err != nil {
		return err
	}
	fw, err = zw.Create("ppt/viewProps.xml")
	if err != nil {
		return err
	}
	if _, err := io.WriteString(fw, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
		`<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:slideViewPr><p:cSldViewPr><p:cViewPr varScale="1"><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:cSldViewPr></p:slideViewPr></p:viewPr>`); err != nil {
		return err
	}
	fw, err = zw.Create("ppt/tableStyles.xml")
	if err != nil {
		return err
	}
	if _, err := io.WriteString(fw, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
		`<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`); err != nil {
		return err
	}

	// Write presentation.xml
	fw, err = zw.Create("ppt/presentation.xml")
	if err != nil {
		return err
	}
	if err := writePresentationXML(fw, slides); err != nil {
		return err
	}

	// Write presentation.xml.rels
	fw, err = zw.Create("ppt/_rels/presentation.xml.rels")
	if err != nil {
		return err
	}
	if err := writePresentationRels(fw, slides, imgRels); err != nil {
		return err
	}

	// Write [Content_Types].xml
	fw, err = zw.Create("[Content_Types].xml")
	if err != nil {
		return err
	}
	if err := writeContentTypes(fw, slides, imgRels); err != nil {
		return err
	}

	// Write _rels/.rels
	fw, err = zw.Create("_rels/.rels")
	if err != nil {
		return err
	}
	if _, err := io.WriteString(fw, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`+
		`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>`+
		`</Relationships>`); err != nil {
		return err
	}

	return nil
}

func writeSlideXML(w io.Writer, slide slideData) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`)
	buf.WriteString(`<p:cSld><p:spTree>`)
	buf.WriteString(`<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`)
	buf.WriteString(`<p:grpSpPr/>`)

	for i, text := range slide.Texts {
		spID := i + 2
		buf.WriteString(fmt.Sprintf(`<p:sp><p:nvSpPr><p:cNvPr id="%d" name="TextBox %d"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`, spID, i+1))
		buf.WriteString(`<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`)
		buf.WriteString(`<p:txBody><a:bodyPr/><a:lstStyle/>`)
		buf.WriteString(`<a:p><a:r><a:t>`)
		xml.Escape(&buf, []byte(filterInvalidXMLChars(text)))
		buf.WriteString(`</a:t></a:r></a:p>`)
		buf.WriteString(`</p:txBody></p:sp>`)
	}

	buf.WriteString(`</p:spTree></p:cSld></p:sld>`)
	_, err := w.Write(buf.Bytes())
	return err
}

func writePresentationXML(w io.Writer, slides []slideData) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`)
	buf.WriteString(`<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster1"/></p:sldMasterIdLst>`)
	buf.WriteString(`<p:sldIdLst>`)

	slideCount := len(slides)
	if slideCount == 0 {
		slideCount = 1
	}
	for i := 0; i < slideCount; i++ {
		buf.WriteString(fmt.Sprintf(`<p:sldId id="%d" r:id="rIdSlide%d"/>`, 256+i, i+1))
	}

	buf.WriteString(`</p:sldIdLst>`)
	buf.WriteString(`<p:sldSz cx="9144000" cy="6858000"/>`)
	buf.WriteString(`<p:notesSz cx="6858000" cy="9144000"/>`)
	buf.WriteString(`</p:presentation>`)
	_, err := w.Write(buf.Bytes())
	return err
}

func writePresentationRels(w io.Writer, slides []slideData, imgRels []imageRel) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)
	buf.WriteString(`<Relationship Id="rIdMaster1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>`)

	slideCount := len(slides)
	if slideCount == 0 {
		slideCount = 1
	}
	for i := 0; i < slideCount; i++ {
		buf.WriteString(fmt.Sprintf(`<Relationship Id="rIdSlide%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide%d.xml"/>`, i+1, i+1))
	}

	for _, rel := range imgRels {
		buf.WriteString(fmt.Sprintf(`<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/%s"/>`, rel.relID, rel.filename))
	}

	buf.WriteString(`<Relationship Id="rIdTheme1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`)
	buf.WriteString(`<Relationship Id="rIdPresProps" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>`)
	buf.WriteString(`<Relationship Id="rIdViewProps" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>`)
	buf.WriteString(`<Relationship Id="rIdTableStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>`)
	buf.WriteString(`</Relationships>`)
	_, err := w.Write(buf.Bytes())
	return err
}

func writeContentTypes(w io.Writer, slides []slideData, imgRels []imageRel) error {
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
	buf.WriteString(`<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>`)
	buf.WriteString(`<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>`)
	buf.WriteString(`<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>`)
	buf.WriteString(`<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>`)
	buf.WriteString(`<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`)

	slideCount := len(slides)
	if slideCount == 0 {
		slideCount = 1
	}
	for i := 0; i < slideCount; i++ {
		buf.WriteString(fmt.Sprintf(`<Override PartName="/ppt/slides/slide%d.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`, i+1))
	}

	buf.WriteString(`<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`)
	buf.WriteString(`<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>`)
	buf.WriteString(`</Types>`)
	_, err := w.Write(buf.Bytes())
	return err
}

// themeXML is a minimal Office theme required for PowerPoint to open the file.
const themeXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>`

// writePptxFormatted generates a PPTX with formatting, shape positions, and images.
func writePptxFormatted(w io.Writer, slides []formattedSlideData, layouts []formattedLayoutData, images []imageData, slideWidth, slideHeight int32) error {
	zw := zip.NewWriter(w)
	defer zw.Close()

	// Write images into ppt/media/
	var imgRels []imageRel
	for i, img := range images {
		ext := (&common.Image{Format: img.Format}).Extension()
		if ext == "" {
			ext = ".bin"
		}
		filename := fmt.Sprintf("image%d%s", i+1, ext)
		fw, err := zw.Create("ppt/media/" + filename)
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

	slideCount := len(slides)
	if slideCount == 0 {
		slideCount = 1
	}

	// Write slide XML files
	for i := 0; i < slideCount; i++ {
		slideNum := i + 1
		fw, err := zw.Create(fmt.Sprintf("ppt/slides/slide%d.xml", slideNum))
		if err != nil {
			return err
		}

		var slide formattedSlideData
		if i < len(slides) {
			slide = slides[i]
		}

		// Get watermark shapes from this slide's layout
		var watermarkShapes []formattedShape
		if slide.LayoutIdx >= 0 && slide.LayoutIdx < len(layouts) {
			watermarkShapes = layouts[slide.LayoutIdx].WatermarkShapes
		}

		// Collect unique image rels for this slide (deduplicate by relID)
		var slideImgRels []imageRel
		seenRels := make(map[string]bool)
		for _, sh := range slide.Shapes {
			if sh.IsImage && sh.ImageIdx >= 0 && sh.ImageIdx < len(imgRels) {
				rel := imgRels[sh.ImageIdx]
				if !seenRels[rel.relID] {
					seenRels[rel.relID] = true
					slideImgRels = append(slideImgRels, rel)
				}
			}
		}
		// Include watermark image rels
		for _, sh := range watermarkShapes {
			if sh.IsImage && sh.ImageIdx >= 0 && sh.ImageIdx < len(imgRels) {
				rel := imgRels[sh.ImageIdx]
				if !seenRels[rel.relID] {
					seenRels[rel.relID] = true
					slideImgRels = append(slideImgRels, rel)
				}
			}
		}
		// Include background image rel if present
		if slide.Background.HasBackground && slide.Background.ImageIdx >= 0 && slide.Background.ImageIdx < len(imgRels) {
			rel := imgRels[slide.Background.ImageIdx]
			if !seenRels[rel.relID] {
				seenRels[rel.relID] = true
				slideImgRels = append(slideImgRels, rel)
			}
		}

		if err := writeFormattedSlideXML(fw, slide, slideImgRels, watermarkShapes); err != nil {
			return err
		}

		// Write slide rels (pointing to correct layout)
		fw, err = zw.Create(fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", slideNum))
		if err != nil {
			return err
		}
		layoutNum := slide.LayoutIdx + 1
		if err := writeFormattedSlideRels(fw, slideImgRels, layoutNum); err != nil {
			return err
		}
	}

	layoutCount := len(layouts)

	// Write slideLayout files with master shapes and background
	for li, layout := range layouts {
		layoutNum := li + 1

		// Collect image rels needed by this layout
		var layoutImgRels []imageRel
		seenRels := make(map[string]bool)
		for _, sh := range layout.Shapes {
			if sh.IsImage && sh.ImageIdx >= 0 && sh.ImageIdx < len(imgRels) {
				rel := imgRels[sh.ImageIdx]
				if !seenRels[rel.relID] {
					seenRels[rel.relID] = true
					layoutImgRels = append(layoutImgRels, rel)
				}
			}
		}
		if layout.Background.HasBackground && layout.Background.ImageIdx >= 0 && layout.Background.ImageIdx < len(imgRels) {
			rel := imgRels[layout.Background.ImageIdx]
			if !seenRels[rel.relID] {
				seenRels[rel.relID] = true
				layoutImgRels = append(layoutImgRels, rel)
			}
		}

		fw, err := zw.Create(fmt.Sprintf("ppt/slideLayouts/slideLayout%d.xml", layoutNum))
		if err != nil {
			return err
		}
		if err := writeLayoutXML(fw, layout, layoutImgRels, slideWidth, slideHeight); err != nil {
			return err
		}

		// Write layout rels
		fw, err = zw.Create(fmt.Sprintf("ppt/slideLayouts/_rels/slideLayout%d.xml.rels", layoutNum))
		if err != nil {
			return err
		}
		if err := writeLayoutRels(fw, layoutImgRels); err != nil {
			return err
		}
	}

	// Write slideMaster referencing all layouts
	{
		fw, err := zw.Create("ppt/slideMasters/slideMaster1.xml")
		if err != nil {
			return err
		}
		if err := writeSlideMasterXML(fw, layoutCount); err != nil {
			return err
		}
	}

	// Write slideMaster rels
	{
		fw, err := zw.Create("ppt/slideMasters/_rels/slideMaster1.xml.rels")
		if err != nil {
			return err
		}
		if err := writeSlideMasterRels(fw, layoutCount); err != nil {
			return err
		}
	}

	// Write theme with color scheme from the most commonly used layout
	var colorScheme []string
	// Count how many slides use each layout
	layoutUsage := make(map[int]int)
	for _, slide := range slides {
		layoutUsage[slide.LayoutIdx]++
	}
	bestLayout := -1
	bestCount := 0
	for idx, count := range layoutUsage {
		if count > bestCount {
			bestCount = count
			bestLayout = idx
		}
	}
	if bestLayout >= 0 && bestLayout < len(layouts) && len(layouts[bestLayout].ColorScheme) >= 8 {
		colorScheme = layouts[bestLayout].ColorScheme
	} else {
		for _, layout := range layouts {
			if len(layout.ColorScheme) >= 8 {
				colorScheme = layout.ColorScheme
				break
			}
		}
	}
	if err := writeThemeXML(zw, colorScheme); err != nil {
		return err
	}

	// Write presProps, viewProps, tableStyles (required for proper rendering)
	if err := writeStaticPart(zw, "ppt/presProps.xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
			`<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`); err != nil {
		return err
	}
	if err := writeStaticPart(zw, "ppt/viewProps.xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
			`<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:slideViewPr><p:cSldViewPr><p:cViewPr varScale="1"><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:cSldViewPr></p:slideViewPr></p:viewPr>`); err != nil {
		return err
	}
	if err := writeStaticPart(zw, "ppt/tableStyles.xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
			`<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`); err != nil {
		return err
	}

	// Write presentation.xml with slide size
	fw, err := zw.Create("ppt/presentation.xml")
	if err != nil {
		return err
	}
	if err := writeFormattedPresentationXML(fw, slideCount, slideWidth, slideHeight); err != nil {
		return err
	}

	// Write presentation.xml.rels
	fw, err = zw.Create("ppt/_rels/presentation.xml.rels")
	if err != nil {
		return err
	}
	if err := writeFormattedPresentationRels(fw, slideCount); err != nil {
		return err
	}

	// Write [Content_Types].xml
	fw, err = zw.Create("[Content_Types].xml")
	if err != nil {
		return err
	}
	if err := writeFormattedContentTypes(fw, slideCount, imgRels, layoutCount); err != nil {
		return err
	}

	// Write _rels/.rels
	if err := writeStaticPart(zw, "_rels/.rels",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`+
			`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>`+
			`</Relationships>`); err != nil {
		return err
	}

	return nil
}

// writeStaticPart writes a static string as a zip entry.
func writeStaticPart(zw *zip.Writer, path, content string) error {
	fw, err := zw.Create(path)
	if err != nil {
		return err
	}
	_, err = io.WriteString(fw, content)
	return err
}

func writeFormattedSlideXML(w io.Writer, slide formattedSlideData, imgRels []imageRel, watermarkShapes []formattedShape) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" showMasterSp="1">`)
	buf.WriteString(`<p:cSld>`)

	// Write background if present
	if slide.Background.HasBackground {
		writeSlideBgXML(&buf, slide.Background, imgRels)
	}

	buf.WriteString(`<p:spTree>`)
	buf.WriteString(`<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`)
	buf.WriteString(`<p:grpSpPr/>`)

	spID := 2

	// Render watermark shapes (from layout) BEFORE slide shapes so they appear
	// behind slide content but on top of the layout background, matching PPT behavior.
	for _, shape := range watermarkShapes {
		writeShapeXML(&buf, shape, spID, imgRels)
		spID++
	}

	for _, shape := range slide.Shapes {
		writeShapeXML(&buf, shape, spID, imgRels)
		spID++
	}

	buf.WriteString(`</p:spTree></p:cSld>`)
	buf.WriteString(`<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>`)
	buf.WriteString(`</p:sld>`)
	_, err := w.Write(buf.Bytes())
	return err
}

// writeSlideBgXML writes the <p:bg> element for a slide background.
func writeSlideBgXML(buf *bytes.Buffer, bg formattedBackground, imgRels []imageRel) {
	buf.WriteString(`<p:bg><p:bgPr>`)

	if bg.ImageIdx >= 0 {
		// Image/blip fill background
		relIDStr := fmt.Sprintf("rImg%d", bg.ImageIdx+1)
		var relID string
		for _, rel := range imgRels {
			if rel.relID == relIDStr {
				relID = rel.relID
				break
			}
		}
		if relID != "" {
			buf.WriteString(fmt.Sprintf(`<a:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`, relID))
		} else if bg.FillColor != "" {
			buf.WriteString(fmt.Sprintf(`<a:solidFill><a:srgbClr val="%s"/></a:solidFill>`, bg.FillColor))
		} else {
			buf.WriteString(`<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>`)
		}
	} else if bg.FillColor != "" {
		buf.WriteString(fmt.Sprintf(`<a:solidFill><a:srgbClr val="%s"/></a:solidFill>`, bg.FillColor))
	} else {
		buf.WriteString(`<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>`)
	}

	buf.WriteString(`<a:effectLst/></p:bgPr></p:bg>`)
}

func writeShapeXML(buf *bytes.Buffer, shape formattedShape, spID int, imgRels []imageRel) {
	if shape.IsImage && shape.ImageIdx >= 0 {
		// Bounds check: skip shapes with out-of-range image index
		relIDStr := fmt.Sprintf("rImg%d", shape.ImageIdx+1)
		var relID string
		for _, rel := range imgRels {
			if rel.relID == relIDStr {
				relID = rel.relID
				break
			}
		}
		if relID == "" {
			return
		}

		buf.WriteString(fmt.Sprintf(`<p:pic><p:nvPicPr><p:cNvPr id="%d" name="Image %d"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>`, spID, spID))
		// Write blipFill with optional cropping
		hasCrop := shape.CropFromTop != 0 || shape.CropFromBottom != 0 || shape.CropFromLeft != 0 || shape.CropFromRight != 0
		if hasCrop {
			// PPT crop values are in 1/65536 units; PPTX srcRect uses 1/1000th of a percent (100000 = 100%)
			t := int64(shape.CropFromTop) * 100000 / 65536
			b := int64(shape.CropFromBottom) * 100000 / 65536
			l := int64(shape.CropFromLeft) * 100000 / 65536
			r := int64(shape.CropFromRight) * 100000 / 65536
			buf.WriteString(fmt.Sprintf(`<p:blipFill><a:blip r:embed="%s"/><a:srcRect l="%d" t="%d" r="%d" b="%d"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`, relID, l, t, r, b))
		} else {
			buf.WriteString(fmt.Sprintf(`<p:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`, relID))
		}
		buf.WriteString(`<p:spPr>`)
		writeXfrm(buf, shape)
		buf.WriteString(`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`)
		// Suppress default line on images to prevent unwanted borders
		if shape.NoLine || shape.LineColor == "" {
			buf.WriteString(`<a:ln><a:noFill/></a:ln>`)
		} else {
			writeLineXML(buf, shape)
		}
		buf.WriteString(`</p:spPr></p:pic>`)
		return
	}

	// Text/shape - check if this is a connector/line
	if isConnectorShape(shape.ShapeType) {
		writeConnectorXML(buf, shape, spID)
		return
	}

	buf.WriteString(fmt.Sprintf(`<p:sp><p:nvSpPr><p:cNvPr id="%d" name="Shape %d"/>`, spID, spID))
	if shape.IsText {
		buf.WriteString(`<p:cNvSpPr txBox="1"/>`)
	} else {
		buf.WriteString(`<p:cNvSpPr/>`)
	}
	buf.WriteString(`<p:nvPr/></p:nvSpPr>`)
	buf.WriteString(`<p:spPr>`)
	writeXfrm(buf, shape)

	// Geometry
	if len(shape.GeoVertices) > 0 {
		writeCustGeomXML(buf, shape)
	} else {
		prst := mapShapeGeometry(shape.ShapeType)
		buf.WriteString(fmt.Sprintf(`<a:prstGeom prst="%s"><a:avLst/></a:prstGeom>`, prst))
	}

	// Fill
	if shape.NoFill {
		buf.WriteString(`<a:noFill/>`)
	} else if shape.FillColor != "" {
		if shape.FillOpacity >= 0 && shape.FillOpacity < 65536 {
			alpha := int64(shape.FillOpacity) * 100000 / 65536
			buf.WriteString(fmt.Sprintf(`<a:solidFill><a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr></a:solidFill>`, shape.FillColor, alpha))
		} else {
			buf.WriteString(fmt.Sprintf(`<a:solidFill><a:srgbClr val="%s"/></a:solidFill>`, shape.FillColor))
		}
	} else {
		// No explicit fill color and no noFill flag: write noFill to prevent
		// PowerPoint from applying a default theme fill to the shape.
		buf.WriteString(`<a:noFill/>`)
	}

	// Line
	if shape.NoLine {
		buf.WriteString(`<a:ln><a:noFill/></a:ln>`)
	} else if shape.LineColor != "" {
		writeLineXML(buf, shape)
	} else if shape.LineWidth > 0 {
		writeLineXML(buf, shape)
	} else if shape.LineStartArrowHead > 0 || shape.LineEndArrowHead > 0 {
		writeLineXML(buf, shape)
	} else {
		// No explicit line: suppress default line to match PPT rendering
		buf.WriteString(`<a:ln><a:noFill/></a:ln>`)
	}

	buf.WriteString(`</p:spPr>`)

	if len(shape.Paragraphs) > 0 {
		buf.WriteString(`<p:txBody>`)
		writeBodyPr(buf, shape)
		buf.WriteString(`<a:lstStyle/>`)
		writeTextBodyXML(buf, shape.Paragraphs)
		buf.WriteString(`</p:txBody>`)
	} else if shape.IsText {
		// Empty textbox still needs a minimal txBody
		buf.WriteString(`<p:txBody>`)
		writeBodyPr(buf, shape)
		buf.WriteString(`<a:lstStyle/><a:p><a:endParaRPr lang="zh-CN" dirty="0" sz="1800"/></a:p></p:txBody>`)
	}

	buf.WriteString(`</p:sp>`)
}



// writeXfrm writes the transform element with position, size, rotation, and flip.
func writeXfrm(buf *bytes.Buffer, shape formattedShape) {
	hasXfrmAttrs := shape.Rotation != 0 || shape.FlipH || shape.FlipV

	// PPT stores the bounding box AFTER rotation (visual bounds on screen).
	// OOXML stores the shape dimensions BEFORE rotation.
	// For rotations near 90° or 270°, we need to swap width/height and
	// adjust the offset so the center stays in the same position.
	x := int64(shape.Left)
	y := int64(shape.Top)
	cx := int64(shape.Width)
	cy := int64(shape.Height)
	if cx < 0 {
		cx = 0
	}
	if cy < 0 {
		cy = 0
	}

	if shape.Rotation != 0 {
		// Convert rotation to degrees (PPT uses fixedPoint 16.16)
		rotDeg := float64(shape.Rotation) / 65536.0
		// Normalize to 0-360
		for rotDeg < 0 {
			rotDeg += 360
		}
		for rotDeg >= 360 {
			rotDeg -= 360
		}
		// Check if rotation is near 90° or 270° (within 5° tolerance)
		needSwap := (rotDeg > 45 && rotDeg < 135) || (rotDeg > 225 && rotDeg < 315)
		if needSwap {
			// Swap width/height and adjust offset to keep center in place
			centerX := x + cx/2
			centerY := y + cy/2
			cx, cy = cy, cx
			x = centerX - cx/2
			y = centerY - cy/2
		}
	}

	if hasXfrmAttrs {
		buf.WriteString(`<a:xfrm`)
		if shape.Rotation != 0 {
			// PPT rotation is in fixedPoint 16.16 format (1/65536 of a degree)
			// OOXML expects 1/60000 of a degree
			rot := int64(shape.Rotation) * 60000 / 65536
			buf.WriteString(fmt.Sprintf(` rot="%d"`, rot))
		}
		if shape.FlipH {
			buf.WriteString(` flipH="1"`)
		}
		if shape.FlipV {
			buf.WriteString(` flipV="1"`)
		}
		buf.WriteString(`>`)
	} else {
		buf.WriteString(`<a:xfrm>`)
	}
	buf.WriteString(fmt.Sprintf(`<a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/>`, x, y, cx, cy))
	buf.WriteString(`</a:xfrm>`)
}

// writeCustGeomXML writes a custom geometry element for freeform shapes.
// Converts PPT pVertices + pSegmentInfo to OOXML <a:custGeom>.
func writeCustGeomXML(buf *bytes.Buffer, shape formattedShape) {
	geoW := shape.GeoRight - shape.GeoLeft
	geoH := shape.GeoBottom - shape.GeoTop
	if geoW <= 0 {
		geoW = 21600 // default PPT coordinate space
	}
	if geoH <= 0 {
		geoH = 21600
	}

	buf.WriteString(`<a:custGeom><a:avLst/>`)
	buf.WriteString(fmt.Sprintf(`<a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="%d" b="%d"/>`, geoW, geoH))
	buf.WriteString(fmt.Sprintf(`<a:pathLst><a:path w="%d" h="%d">`, geoW, geoH))

	verts := shape.GeoVertices
	segs := shape.GeoSegments
	vi := 0 // vertex index

	if len(segs) > 0 {
		for _, seg := range segs {
			switch seg.SegType {
			case 0: // lineTo
				count := int(seg.Count)
				if count == 0 {
					count = 1
				}
				for j := 0; j < count && vi < len(verts); j++ {
					v := verts[vi]
					buf.WriteString(fmt.Sprintf(`<a:lnTo><a:pt x="%d" y="%d"/></a:lnTo>`, v.X-shape.GeoLeft, v.Y-shape.GeoTop))
					vi++
				}
			case 1: // curveTo (cubic bezier, 3 points per curve)
				count := int(seg.Count)
				if count == 0 {
					count = 1
				}
				for j := 0; j < count && vi+2 < len(verts); j++ {
					v1 := verts[vi]
					v2 := verts[vi+1]
					v3 := verts[vi+2]
					buf.WriteString(fmt.Sprintf(`<a:cubicBezTo><a:pt x="%d" y="%d"/><a:pt x="%d" y="%d"/><a:pt x="%d" y="%d"/></a:cubicBezTo>`,
						v1.X-shape.GeoLeft, v1.Y-shape.GeoTop,
						v2.X-shape.GeoLeft, v2.Y-shape.GeoTop,
						v3.X-shape.GeoLeft, v3.Y-shape.GeoTop))
					vi += 3
				}
			case 2: // moveTo
				if vi < len(verts) {
					v := verts[vi]
					buf.WriteString(fmt.Sprintf(`<a:moveTo><a:pt x="%d" y="%d"/></a:moveTo>`, v.X-shape.GeoLeft, v.Y-shape.GeoTop))
					vi++
				}
			case 3: // close
				buf.WriteString(`<a:close/>`)
			case 4: // end
				// end of path, do nothing
			case 5: // escape (special commands, skip)
				// escape commands may consume vertices
				count := int(seg.Count)
				vi += count
			}
		}
	} else {
		// No segment info: treat all vertices as a polygon (moveTo first, lineTo rest, close)
		if len(verts) > 0 {
			v := verts[0]
			buf.WriteString(fmt.Sprintf(`<a:moveTo><a:pt x="%d" y="%d"/></a:moveTo>`, v.X-shape.GeoLeft, v.Y-shape.GeoTop))
			for i := 1; i < len(verts); i++ {
				v = verts[i]
				buf.WriteString(fmt.Sprintf(`<a:lnTo><a:pt x="%d" y="%d"/></a:lnTo>`, v.X-shape.GeoLeft, v.Y-shape.GeoTop))
			}
			buf.WriteString(`<a:close/>`)
		}
	}

	buf.WriteString(`</a:path></a:pathLst></a:custGeom>`)
}

// mapShapeGeometry maps PPT shape type IDs to OOXML preset geometry names.
func mapShapeGeometry(shapeType uint16) string {
	switch shapeType {
	case 0: // msosptNotPrimitive (freeform/group placeholder)
		return "rect"
	case 1: // msosptRectangle
		return "rect"
	case 2: // msosptRoundRectangle
		return "roundRect"
	case 3: // msosptEllipse
		return "ellipse"
	case 4: // msosptDiamond
		return "diamond"
	case 5: // msosptIsocelesTriangle
		return "triangle"
	case 6: // msosptRightTriangle
		return "rtTriangle"
	case 7: // msosptParallelogram
		return "parallelogram"
	case 8: // msosptTrapezoid
		return "trapezoid"
	case 9: // msosptHexagon
		return "hexagon"
	case 10: // msosptOctagon
		return "octagon"
	case 11: // msosptPlus
		return "mathPlus"
	case 12: // msosptStar
		return "star5"
	case 13: // msosptArrow (right arrow)
		return "rightArrow"
	case 14: // msosptThickArrow (not in OOXML, use rightArrow)
		return "rightArrow"
	case 15: // msosptHomePlate
		return "homePlate"
	case 16: // msosptCube
		return "cube"
	case 17: // msosptBalloon
		return "wedgeRoundRectCallout"
	case 19: // msosptArc
		return "arc"
	case 20: // msosptLine
		return "line"
	case 21: // msosptPlaque
		return "plaque"
	case 22: // msosptCan
		return "can"
	case 23: // msosptDonut
		return "donut"
	case 32: // msosptStraightConnector1
		return "straightConnector1"
	case 33: // msosptBentConnector2
		return "bentConnector2"
	case 34: // msosptBentConnector3
		return "bentConnector3"
	case 35: // msosptBentConnector4
		return "bentConnector4"
	case 36: // msosptBentConnector5
		return "bentConnector5"
	case 37: // msosptCurvedConnector2
		return "curvedConnector2"
	case 38: // msosptCurvedConnector3
		return "curvedConnector3"
	case 39: // msosptCurvedConnector4
		return "curvedConnector4"
	case 40: // msosptCurvedConnector5
		return "curvedConnector5"
	case 55: // msosptNoSmoking
		return "noSmoking"
	case 56: // msosptRightArrow
		return "rightArrow"
	case 57: // msosptLeftArrow
		return "leftArrow"
	case 58: // msosptUpArrow
		return "upArrow"
	case 59: // msosptDownArrow
		return "downArrow"
	case 60: // msosptLeftRightArrow
		return "leftRightArrow"
	case 61: // msosptUpDownArrow
		return "upDownArrow"
	case 66: // msosptIrregularSeal1
		return "irregularSeal1"
	case 67: // msosptWedgeRectCallout
		return "wedgeRectCallout"
	case 68: // msosptWedgeRRectCallout
		return "wedgeRoundRectCallout"
	case 69: // msosptWedgeEllipseCallout
		return "wedgeEllipseCallout"
	case 72: // msosptCloudCallout
		return "cloudCallout"
	case 75: // msosptPictureFrame
		return "rect"
	case 84: // msosptBevel
		return "bevel"
	case 85: // msosptFoldedCorner
		return "foldedCorner"
	case 87: // msosptSmileyFace
		return "smileyFace"
	case 91: // msosptHeart
		return "heart"
	case 92: // msosptLightningBolt
		return "lightningBolt"
	case 93: // msosptSun
		return "sun"
	case 94: // msosptFlowChartProcess
		return "flowChartProcess"
	case 95: // msosptFlowChartDecision
		return "flowChartDecision"
	case 96: // msosptFlowChartInputOutput
		return "flowChartInputOutput"
	case 97: // msosptFlowChartPredefinedProcess
		return "flowChartPredefinedProcess"
	case 98: // msosptFlowChartInternalStorage
		return "flowChartInternalStorage"
	case 99: // msosptFlowChartDocument
		return "flowChartDocument"
	case 100: // msosptFlowChartMultidocument
		return "flowChartMultidocument"
	case 101: // msosptFlowChartMerge
		return "flowChartMerge"
	case 102: // msosptFlowChartOnlineStorage
		return "flowChartOnlineStorage"
	case 103: // msosptFlowChartDelay
		return "flowChartDelay"
	case 104: // msosptFlowChartSequentialAccessStorage
		return "flowChartMagneticTape"
	case 105: // msosptFlowChartMagneticDisk
		return "flowChartMagneticDisk"
	case 106: // msosptFlowChartTerminator
		return "flowChartTerminator"
	case 107: // msosptFlowChartPunchedTape
		return "flowChartPunchedTape"
	case 108: // msosptFlowChartSummingJunction
		return "flowChartSummingJunction"
	case 109: // msosptFlowChartOr
		return "flowChartOr"
	case 110: // msosptCallout1
		return "callout1"
	case 111: // msosptCallout2
		return "callout2"
	case 112: // msosptCallout3
		return "callout3"
	case 114: // msosptRibbon2
		return "ribbon2"
	case 115: // msosptRibbon
		return "ribbon"
	case 116: // msosptCallout2 (accentCallout2)
		return "accentCallout2"
	case 120: // msosptChevron
		return "chevron"
	case 121: // msosptPentagon
		return "pentagon"
	case 122: // msosptBlockArc
		return "blockArc"
	case 183: // msosptLeftBrace
		return "leftBrace"
	case 184: // msosptRightBrace
		return "rightBrace"
	case 185: // msosptLeftBracket
		return "leftBracket"
	case 186: // msosptRightBracket
		return "rightBracket"
	case 202: // msosptTextBox
		return "rect"
	default:
		return "rect"
	}
}

// writeLineXML writes the <a:ln> element with optional width and color.
func writeLineXML(buf *bytes.Buffer, shape formattedShape) {
	if shape.LineWidth > 0 {
		buf.WriteString(fmt.Sprintf(`<a:ln w="%d">`, shape.LineWidth))
	} else {
		buf.WriteString(`<a:ln>`)
	}
	if shape.LineColor != "" {
		buf.WriteString(fmt.Sprintf(`<a:solidFill><a:srgbClr val="%s"/></a:solidFill>`, shape.LineColor))
	} else {
		// Default to black line when width is set but no color specified
		buf.WriteString(`<a:solidFill><a:srgbClr val="000000"/></a:solidFill>`)
	}
	// Line dash style
	if shape.LineDash > 0 {
		dashMap := []string{"solid", "dash", "dot", "dashDot", "dashDotDot",
			"solid", "lgDash", "lgDashDot", "lgDashDotDot", "sysDash", "sysDot", "sysDashDot", "sysDashDotDot"}
		if int(shape.LineDash) < len(dashMap) {
			buf.WriteString(fmt.Sprintf(`<a:prstDash val="%s"/>`, dashMap[shape.LineDash]))
		}
	}
	// Arrow head at line start
	if shape.LineStartArrowHead > 0 {
		writeArrowEndXML(buf, "headEnd", shape.LineStartArrowHead, shape.LineStartArrowWidth, shape.LineStartArrowLength)
	}
	// Arrow head at line end
	if shape.LineEndArrowHead > 0 {
		writeArrowEndXML(buf, "tailEnd", shape.LineEndArrowHead, shape.LineEndArrowWidth, shape.LineEndArrowLength)
	}
	buf.WriteString(`</a:ln>`)
}

// writeArrowEndXML writes an <a:headEnd> or <a:tailEnd> element for line arrows.
// arrowType: 1=triangle, 2=stealth, 3=diamond, 4=oval, 5=open (per MS-ODRAW)
// arrowWidth: 0=narrow(sm), 1=medium(med), 2=wide(lg), -1=not set
// arrowLength: 0=short(sm), 1=medium(med), 2=long(lg), -1=not set
func writeArrowEndXML(buf *bytes.Buffer, elemName string, arrowType, arrowWidth, arrowLength int32) {
	typeMap := []string{"none", "triangle", "stealth", "diamond", "oval", "arrow"}
	typeName := "triangle"
	if int(arrowType) < len(typeMap) {
		typeName = typeMap[arrowType]
	}
	widthMap := []string{"sm", "med", "lg"}
	lengthMap := []string{"sm", "med", "lg"}
	attrs := fmt.Sprintf(` type="%s"`, typeName)
	if arrowWidth >= 0 && int(arrowWidth) < len(widthMap) {
		attrs += fmt.Sprintf(` w="%s"`, widthMap[arrowWidth])
	}
	if arrowLength >= 0 && int(arrowLength) < len(lengthMap) {
		attrs += fmt.Sprintf(` len="%s"`, lengthMap[arrowLength])
	}
	buf.WriteString(fmt.Sprintf(`<a:%s%s/>`, elemName, attrs))
}


// writeBodyPr writes the <a:bodyPr> element with text margins and anchor.
func writeBodyPr(buf *bytes.Buffer, shape formattedShape) {
	hasAttrs := shape.TextMarginLeft >= 0 || shape.TextMarginTop >= 0 ||
		shape.TextMarginRight >= 0 || shape.TextMarginBottom >= 0 ||
		shape.TextAnchor >= 0 || shape.TextWordWrap >= 0
	if !hasAttrs {
		buf.WriteString(`<a:bodyPr><a:normAutofit/></a:bodyPr>`)
		return
	}
	buf.WriteString(`<a:bodyPr`)
	if shape.TextMarginLeft >= 0 {
		buf.WriteString(fmt.Sprintf(` lIns="%d"`, shape.TextMarginLeft))
	}
	if shape.TextMarginTop >= 0 {
		buf.WriteString(fmt.Sprintf(` tIns="%d"`, shape.TextMarginTop))
	}
	if shape.TextMarginRight >= 0 {
		buf.WriteString(fmt.Sprintf(` rIns="%d"`, shape.TextMarginRight))
	}
	if shape.TextMarginBottom >= 0 {
		buf.WriteString(fmt.Sprintf(` bIns="%d"`, shape.TextMarginBottom))
	}
	if shape.TextAnchor >= 0 {
		// PPT TextAnchor: 0=top, 1=middle, 2=bottom, 3=topCentered, 4=middleCentered, 5=bottomCentered
		anchorMap := []string{"t", "ctr", "b", "t", "ctr", "b"}
		if int(shape.TextAnchor) < len(anchorMap) {
			buf.WriteString(fmt.Sprintf(` anchor="%s"`, anchorMap[shape.TextAnchor]))
		}
		// Centered variants (3, 4, 5) need anchorCtr="1"
		if shape.TextAnchor >= 3 && shape.TextAnchor <= 5 {
			buf.WriteString(` anchorCtr="1"`)
		}
	}
	if shape.TextWordWrap == 0 {
		buf.WriteString(` wrap="none"`)
	}
	buf.WriteString(`>`)
	// Use normAutofit to allow PowerPoint to shrink text to fit the shape.
	// For shapes with very small height relative to text content, allow more
	// aggressive shrinking by setting a lower fontScale.
	buf.WriteString(`<a:normAutofit/>`)
	buf.WriteString(`</a:bodyPr>`)
}

// isConnectorShape returns true if the shape type is a line or connector.
func isConnectorShape(shapeType uint16) bool {
	switch shapeType {
	case 20: // msosptLine
		return true
	case 32, 33, 34, 35, 36, 37, 38, 39, 40: // all connector types
		return true
	}
	return false
}

// writeConnectorXML writes a connector shape (<p:cxnSp>) element.
func writeConnectorXML(buf *bytes.Buffer, shape formattedShape, spID int) {
	buf.WriteString(fmt.Sprintf(`<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="%d" name="Connector %d"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>`, spID, spID))
	buf.WriteString(`<p:spPr>`)
	writeXfrm(buf, shape)

	prst := mapShapeGeometry(shape.ShapeType)
	buf.WriteString(fmt.Sprintf(`<a:prstGeom prst="%s"><a:avLst/></a:prstGeom>`, prst))

	// Connectors don't need fill
	buf.WriteString(`<a:noFill/>`)

	// Line
	if shape.NoLine {
		buf.WriteString(`<a:ln><a:noFill/></a:ln>`)
	} else if shape.LineColor != "" || shape.LineWidth > 0 || shape.LineStartArrowHead > 0 || shape.LineEndArrowHead > 0 {
		writeLineXML(buf, shape)
	} else {
		// Default connector line: thin black
		buf.WriteString(`<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`)
	}

	buf.WriteString(`</p:spPr></p:cxnSp>`)
}

func writeTextBodyXML(buf *bytes.Buffer, paragraphs []formattedSlideParagraph) {
	for _, para := range paragraphs {
		buf.WriteString(`<a:p>`)

		// Paragraph properties
		hasPPr := para.Alignment > 0 || para.IndentLevel > 0 ||
			para.SpaceBefore != 0 || para.SpaceAfter != 0 || para.LineSpacing != 0 ||
			para.HasBullet || para.LeftMargin != 0 || para.Indent != 0
		if hasPPr {
			buf.WriteString(`<a:pPr`)
			if para.Alignment > 0 {
				algnMap := []string{"l", "ctr", "r", "just"}
				if int(para.Alignment) < len(algnMap) {
					buf.WriteString(fmt.Sprintf(` algn="%s"`, algnMap[para.Alignment]))
				}
			}
			if para.IndentLevel > 0 {
				buf.WriteString(fmt.Sprintf(` lvl="%d"`, para.IndentLevel))
			}
			// Left margin: PPT master units → EMU (multiply by 12700/8)
			if para.LeftMargin != 0 {
				marL := int64(para.LeftMargin) * 12700 / 8
				buf.WriteString(fmt.Sprintf(` marL="%d"`, marL))
			}
			// First line indent: PPT master units → EMU
			if para.Indent != 0 {
				indent := int64(para.Indent) * 12700 / 8
				buf.WriteString(fmt.Sprintf(` indent="%d"`, indent))
			}
			buf.WriteString(`>`)

			// Line spacing: positive = percentage*100, negative = centipoints
			if para.LineSpacing != 0 {
				if para.LineSpacing > 0 {
					// Percentage: PPT stores as percentage (e.g., 100 = single space)
					// OOXML spcPct val is in 1/1000 of a percent
					pct := int64(para.LineSpacing) * 1000
					buf.WriteString(fmt.Sprintf(`<a:lnSpc><a:spcPct val="%d"/></a:lnSpc>`, pct))
				} else {
					// Absolute: PPT stores as negative centipoints
					// OOXML spcPts val is in hundredths of a point
					pts := -para.LineSpacing
					buf.WriteString(fmt.Sprintf(`<a:lnSpc><a:spcPts val="%d"/></a:lnSpc>`, pts))
				}
			}
			// Space before: positive = percentage, negative = centipoints
			if para.SpaceBefore != 0 {
				if para.SpaceBefore > 0 {
					pct := int64(para.SpaceBefore) * 1000
					buf.WriteString(fmt.Sprintf(`<a:spcBef><a:spcPct val="%d"/></a:spcBef>`, pct))
				} else {
					pts := -para.SpaceBefore
					buf.WriteString(fmt.Sprintf(`<a:spcBef><a:spcPts val="%d"/></a:spcBef>`, pts))
				}
			}
			// Space after
			if para.SpaceAfter != 0 {
				if para.SpaceAfter > 0 {
					pct := int64(para.SpaceAfter) * 1000
					buf.WriteString(fmt.Sprintf(`<a:spcAft><a:spcPct val="%d"/></a:spcAft>`, pct))
				} else {
					pts := -para.SpaceAfter
					buf.WriteString(fmt.Sprintf(`<a:spcAft><a:spcPts val="%d"/></a:spcAft>`, pts))
				}
			}
			if para.HasBullet && para.BulletChar != "" {
				if para.BulletSize != 0 {
					buf.WriteString(fmt.Sprintf(`<a:buSzPct val="%d"/>`, int32(para.BulletSize)*1000))
				}
				if para.BulletColor != "" {
					buf.WriteString(fmt.Sprintf(`<a:buClr><a:srgbClr val="%s"/></a:buClr>`, para.BulletColor))
				}
				if para.BulletFont != "" {
					buf.WriteString(fmt.Sprintf(`<a:buFont typeface="%s"/>`, xmlEscapeAttr(para.BulletFont)))
				}
				buf.WriteString(fmt.Sprintf(`<a:buChar char="%s"/>`, xmlEscapeAttr(para.BulletChar)))
			} else if para.HasBullet {
				// HasBullet with empty BulletChar means the PPT set the bullet flag
				// but didn't specify a character. Suppress the bullet (buNone) to avoid
				// adding unwanted bullet points to non-bullet paragraphs.
				buf.WriteString(`<a:buNone/>`)
			}

			buf.WriteString(`</a:pPr>`)
		}

		// Text runs
		hasContent := false
		for _, run := range para.Runs {
			// Handle vertical tab (U+000B) as line break <a:br/>
			text := run.Text
			parts := splitByVerticalTab(text)
			for pi, part := range parts {
				if pi > 0 {
					// Insert line break with same run properties
					buf.WriteString(`<a:br>`)
					writeSlideRunProperties(buf, run)
					buf.WriteString(`</a:br>`)
					hasContent = true
				}
				if part != "" {
					buf.WriteString(`<a:r>`)
					writeSlideRunProperties(buf, run)
					filtered := filterInvalidXMLChars(part)
					// Add xml:space="preserve" if text has leading/trailing whitespace
					if len(filtered) > 0 && (filtered[0] == ' ' || filtered[len(filtered)-1] == ' ' || strings.ContainsAny(filtered, "\t")) {
						buf.WriteString(`<a:t xml:space="preserve">`)
					} else {
						buf.WriteString(`<a:t>`)
					}
					xml.Escape(buf, []byte(filtered))
					buf.WriteString(`</a:t></a:r>`)
					hasContent = true
				}
			}
		}

		// OOXML requires at least endParaRPr if paragraph has no runs
		if !hasContent {
			// Use the paragraph's first run font size if available
			endSz := uint16(1800)
			for _, run := range para.Runs {
				if run.FontSize > 0 {
					endSz = run.FontSize
					break
				}
			}
			buf.WriteString(fmt.Sprintf(`<a:endParaRPr lang="zh-CN" dirty="0" sz="%d"/>`, endSz))
		}

		buf.WriteString(`</a:p>`)
	}
}

// filterInvalidXMLChars removes characters that are invalid in XML 1.0
func filterInvalidXMLChars(s string) string {
	var buf bytes.Buffer
	for _, r := range s {
		// Valid XML 1.0 characters: #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
		if r == 0x09 || r == 0x0A || r == 0x0D || (r >= 0x20 && r <= 0xD7FF) || (r >= 0xE000 && r <= 0xFFFD) || (r >= 0x10000 && r <= 0x10FFFF) {
			buf.WriteRune(r)
		}
	}
	return buf.String()
}

// splitByVerticalTab splits text by vertical tab (U+000B) which represents
// a soft return / line break in PPT. Returns at least one element.
func splitByVerticalTab(s string) []string {
	var parts []string
	start := 0
	for i, r := range s {
		if r == 0x0B {
			parts = append(parts, s[start:i])
			start = i + 1 // \v is 1 byte in UTF-8
		}
	}
	parts = append(parts, s[start:])
	return parts
}

func writeSlideRunProperties(buf *bytes.Buffer, run formattedSlideRun) {
	hasProps := run.FontName != "" || run.FontSize > 0 || run.Bold || run.Italic || run.Underline || run.Color != ""
	if !hasProps {
		// Even with no explicit props, set a default font size so text renders properly
		buf.WriteString(`<a:rPr lang="zh-CN" altLang="en-US" dirty="0" sz="1800"/>`)
		return
	}

	buf.WriteString(`<a:rPr lang="zh-CN" altLang="en-US" dirty="0"`)
	// Always emit font size; default to 1800 (18pt) if not specified
	fontSize := run.FontSize
	if fontSize == 0 {
		fontSize = 1800
	}
	buf.WriteString(fmt.Sprintf(` sz="%d"`, fontSize))
	if run.Bold {
		buf.WriteString(` b="1"`)
	}
	if run.Italic {
		buf.WriteString(` i="1"`)
	}
	if run.Underline {
		buf.WriteString(` u="sng"`)
	}
	buf.WriteString(`>`)

	if run.Color != "" {
		buf.WriteString(fmt.Sprintf(`<a:solidFill><a:srgbClr val="%s"/></a:solidFill>`, run.Color))
	}
	if run.FontName != "" {
		escaped := xmlEscapeAttr(run.FontName)
		buf.WriteString(fmt.Sprintf(`<a:latin typeface="%s"/><a:ea typeface="%s"/><a:cs typeface="%s"/>`, escaped, escaped, escaped))
	}

	buf.WriteString(`</a:rPr>`)
}


func xmlEscapeAttr(s string) string {
	var buf bytes.Buffer
	for _, r := range s {
		// Filter out invalid XML characters (control chars except tab, newline, carriage return)
		if r < 0x09 || (r > 0x0D && r < 0x20) || r == 0xFFFE || r == 0xFFFF {
			continue // skip invalid XML characters
		}
		switch r {
		case '"':
			buf.WriteString("&quot;")
		case '&':
			buf.WriteString("&amp;")
		case '<':
			buf.WriteString("&lt;")
		case '>':
			buf.WriteString("&gt;")
		default:
			buf.WriteRune(r)
		}
	}
	return buf.String()
}

// isLayoutPlaceholderShape returns true if a shape in a layout is a master
// placeholder (e.g., "单击此处编辑母版") that should not appear in the output.
// These are editing hints in the PPT master that have no value in the PPTX.
func isLayoutPlaceholderShape(shape formattedShape) bool {
	if !shape.IsText || len(shape.Paragraphs) == 0 {
		return false
	}
	for _, para := range shape.Paragraphs {
		for _, run := range para.Runs {
			t := strings.TrimSpace(run.Text)
			if t == "" {
				continue
			}
			for _, prefix := range masterPlaceholderPrefixes {
				if strings.Contains(t, prefix) {
					return true
				}
			}
		}
	}
	return false
}

// isFullPageImage returns true if a shape is an image covering most of the slide area.
// These should be rendered as background fills, not as shapes.
func isFullPageImage(shape formattedShape, slideW, slideH int32) bool {
	if !shape.IsImage || shape.ImageIdx < 0 {
		return false
	}
	// Check if the image covers >70% of the slide in both dimensions
	return shape.Width > int32(float64(slideW)*0.7) && shape.Height > int32(float64(slideH)*0.7)
}

// isWatermarkShape returns true if a shape looks like a watermark/logo that should
// render on top of slide content. Watermarks are typically images positioned in the
// bottom portion of the slide, often extending beyond the slide boundary.
func isWatermarkShape(shape formattedShape, slideW, slideH int32) bool {
	if !shape.IsImage || shape.ImageIdx < 0 {
		return false
	}
	// A watermark image is a small logo/icon positioned in the lower portion
	// of the slide (top position > 50% of slide height).
	// Large images (width > 40% of slide) are decorative layout backgrounds,
	// not watermarks — they should stay in the layout.
	if int64(shape.Width) > int64(slideW)*40/100 {
		return false
	}
	// Images with significant cropping are decorative elements, not watermarks
	if shape.CropFromLeft > 1000 || shape.CropFromTop > 1000 ||
		shape.CropFromRight > 1000 || shape.CropFromBottom > 1000 {
		return false
	}
	bottom := int64(shape.Top) + int64(shape.Height)
	return int64(shape.Top) > int64(slideH)/2 && bottom > int64(slideH)*3/4
}


// separateWatermarkShapes moves watermark-like shapes from layout.Shapes into
// layout.WatermarkShapes so they can be rendered on top of slide content instead
// of behind it. This ensures bottom logos/watermarks are not obscured by slide shapes.
func separateWatermarkShapes(layouts []formattedLayoutData, slideW, slideH int32) {
	for i := range layouts {
		var regular []formattedShape
		var watermarks []formattedShape
		for _, shape := range layouts[i].Shapes {
			if isLayoutPlaceholderShape(shape) {
				regular = append(regular, shape)
				continue
			}
			if isFullPageImage(shape, slideW, slideH) {
				regular = append(regular, shape)
				continue
			}
			if isWatermarkShape(shape, slideW, slideH) {
				watermarks = append(watermarks, shape)
			} else {
				regular = append(regular, shape)
			}
		}
		if len(watermarks) > 0 {
			layouts[i].Shapes = regular
			layouts[i].WatermarkShapes = watermarks
		}
	}
}

// writeLayoutXML writes a slideLayout XML file with the master's shapes and background.
func writeLayoutXML(w io.Writer, layout formattedLayoutData, imgRels []imageRel, slideW, slideH int32) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" type="blank" preserve="1" showMasterSp="0">`)

	buf.WriteString(`<p:cSld>`)

	// Check if any shape is a full-page image that should become the background
	var bgImageIdx int = -1
	var bgCropT, bgCropB, bgCropL, bgCropR int32
	for _, shape := range layout.Shapes {
		if isLayoutPlaceholderShape(shape) {
			continue
		}
		if isFullPageImage(shape, slideW, slideH) {
			bgImageIdx = shape.ImageIdx
			bgCropT = shape.CropFromTop
			bgCropB = shape.CropFromBottom
			bgCropL = shape.CropFromLeft
			bgCropR = shape.CropFromRight
			break // use the first full-page image as background
		}
	}

	// Write background: prefer full-page image as blipFill background
	if bgImageIdx >= 0 {
		relIDStr := fmt.Sprintf("rImg%d", bgImageIdx+1)
		var relID string
		for _, rel := range imgRels {
			if rel.relID == relIDStr {
				relID = rel.relID
				break
			}
		}
		if relID != "" {
			buf.WriteString(`<p:bg><p:bgPr>`)
			hasCrop := bgCropT != 0 || bgCropB != 0 || bgCropL != 0 || bgCropR != 0
			if hasCrop {
				t := int64(bgCropT) * 100000 / 65536
				b := int64(bgCropB) * 100000 / 65536
				l := int64(bgCropL) * 100000 / 65536
				r := int64(bgCropR) * 100000 / 65536
				buf.WriteString(fmt.Sprintf(`<a:blipFill><a:blip r:embed="%s"/><a:srcRect l="%d" t="%d" r="%d" b="%d"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`, relID, l, t, r, b))
			} else {
				buf.WriteString(fmt.Sprintf(`<a:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`, relID))
			}
			buf.WriteString(`<a:effectLst/></p:bgPr></p:bg>`)
		} else if layout.Background.HasBackground {
			writeSlideBgXML(&buf, layout.Background, imgRels)
		}
	} else if layout.Background.HasBackground {
		writeSlideBgXML(&buf, layout.Background, imgRels)
	}

	buf.WriteString(`<p:spTree>`)
	buf.WriteString(`<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`)
	buf.WriteString(`<p:grpSpPr/>`)

	// Write master shapes into the layout, skipping placeholder text shapes
	// and full-page images (already used as background)
	spID := 2

	// If this layout has a detected title background gradient, add it as the first shape
	// (behind all other shapes) so the title text is visible.
	if layout.TitleBgColor != "" && layout.TitleBgBottom > 0 {
		// Create a gradient fill rectangle from top of slide to the connector line
		buf.WriteString(fmt.Sprintf(`<p:sp><p:nvSpPr><p:cNvPr id="%d" name="Title Background"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>`, spID))
		buf.WriteString(`<p:spPr>`)
		buf.WriteString(fmt.Sprintf(`<a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>`, slideW, layout.TitleBgBottom))
		buf.WriteString(`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`)
		// Gradient fill: dark color on left fading to slightly lighter on right
		buf.WriteString(`<a:gradFill>`)
		buf.WriteString(`<a:gsLst>`)
		buf.WriteString(fmt.Sprintf(`<a:gs pos="0"><a:srgbClr val="%s"/></a:gs>`, layout.TitleBgColor))
		// Lighten the color slightly for the right side
		buf.WriteString(fmt.Sprintf(`<a:gs pos="100000"><a:srgbClr val="%s"><a:tint val="60000"/></a:srgbClr></a:gs>`, layout.TitleBgColor))
		buf.WriteString(`</a:gsLst>`)
		buf.WriteString(`<a:lin ang="0" scaled="1"/>`)
		buf.WriteString(`</a:gradFill>`)
		buf.WriteString(`<a:ln><a:noFill/></a:ln>`)
		buf.WriteString(`</p:spPr></p:sp>`)
		spID++
	}

	for _, shape := range layout.Shapes {
		if isLayoutPlaceholderShape(shape) {
			continue
		}
		if isFullPageImage(shape, slideW, slideH) {
			continue // already rendered as background
		}
		if isConnectorShape(shape.ShapeType) {
			writeConnectorXML(&buf, shape, spID)
		} else {
			writeShapeXML(&buf, shape, spID, imgRels)
		}
		spID++
	}

	buf.WriteString(`</p:spTree>`)
	buf.WriteString(`</p:cSld>`)
	buf.WriteString(`</p:sldLayout>`)
	_, err := w.Write(buf.Bytes())
	return err
}


// writeLayoutRels writes the rels file for a slideLayout.
func writeLayoutRels(w io.Writer, imgRels []imageRel) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)
	buf.WriteString(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>`)
	for _, rel := range imgRels {
		buf.WriteString(fmt.Sprintf(`<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/%s"/>`, rel.relID, rel.filename))
	}
	buf.WriteString(`</Relationships>`)
	_, err := w.Write(buf.Bytes())
	return err
}

// writeSlideMasterXML writes the slideMaster XML referencing all layouts.
func writeSlideMasterXML(w io.Writer, layoutCount int) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`)
	buf.WriteString(`<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>`)
	buf.WriteString(`<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>`)
	buf.WriteString(`<p:sldLayoutIdLst>`)
	for i := 0; i < layoutCount; i++ {
		buf.WriteString(fmt.Sprintf(`<p:sldLayoutId id="%d" r:id="rIdLayout%d"/>`, 2147483649+i, i+1))
	}
	buf.WriteString(`</p:sldLayoutIdLst>`)
	// Default text styles for body text (inherited by all slides)
	buf.WriteString(`<p:txStyles>`)
	buf.WriteString(`<p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="3200" kern="1200"><a:latin typeface="微软雅黑"/><a:ea typeface="微软雅黑"/></a:defRPr></a:lvl1pPr></p:titleStyle>`)
	buf.WriteString(`<p:bodyStyle><a:lvl1pPr><a:defRPr sz="1800" kern="1200"><a:latin typeface="微软雅黑"/><a:ea typeface="微软雅黑"/></a:defRPr></a:lvl1pPr></p:bodyStyle>`)
	buf.WriteString(`<p:otherStyle><a:lvl1pPr><a:defRPr sz="1800" kern="1200"><a:latin typeface="微软雅黑"/><a:ea typeface="微软雅黑"/></a:defRPr></a:lvl1pPr></p:otherStyle>`)
	buf.WriteString(`</p:txStyles>`)
	buf.WriteString(`</p:sldMaster>`)
	_, err := w.Write(buf.Bytes())
	return err
}

// writeSlideMasterRels writes the rels file for the slideMaster.
func writeSlideMasterRels(w io.Writer, layoutCount int) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)
	for i := 0; i < layoutCount; i++ {
		buf.WriteString(fmt.Sprintf(`<Relationship Id="rIdLayout%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout%d.xml"/>`, i+1, i+1))
	}
	buf.WriteString(`<Relationship Id="rIdTheme1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>`)
	buf.WriteString(`</Relationships>`)
	_, err := w.Write(buf.Bytes())
	return err
}

// writeThemeXML writes the theme XML with colors from the PPT color scheme.
func writeThemeXML(zw *zip.Writer, colorScheme []string) error {
	fw, err := zw.Create("ppt/theme/theme1.xml")
	if err != nil {
		return err
	}

	// PPT ColorSchemeAtom indices:
	// 0=background, 1=text/lines, 2=shadow, 3=title text,
	// 4=fill, 5=accent, 6=hyperlink, 7=followed hyperlink
	// Map to OOXML theme:
	// lt1=background(0), dk1=text(1), lt2=shadow(2), dk2=title(3),
	// accent1=fill(4), accent2=accent(5), hlink=hyperlink(6), folHlink=followed(7)

	lt1 := "FFFFFF"
	dk1 := "000000"
	lt2 := "E7E6E6"
	dk2 := "44546A"
	accent1 := "4472C4"
	accent2 := "ED7D31"
	hlink := "0563C1"
	folHlink := "954F72"

	if len(colorScheme) >= 8 {
		lt1 = colorScheme[0]
		dk1 = colorScheme[1]
		lt2 = colorScheme[2]
		dk2 = colorScheme[3]
		accent1 = colorScheme[4]
		accent2 = colorScheme[5]
		hlink = colorScheme[6]
		folHlink = colorScheme[7]
	}

	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">`)
	buf.WriteString(`<a:themeElements>`)
	buf.WriteString(`<a:clrScheme name="PPT Theme">`)
	buf.WriteString(fmt.Sprintf(`<a:dk1><a:srgbClr val="%s"/></a:dk1>`, dk1))
	buf.WriteString(fmt.Sprintf(`<a:lt1><a:srgbClr val="%s"/></a:lt1>`, lt1))
	buf.WriteString(fmt.Sprintf(`<a:dk2><a:srgbClr val="%s"/></a:dk2>`, dk2))
	buf.WriteString(fmt.Sprintf(`<a:lt2><a:srgbClr val="%s"/></a:lt2>`, lt2))
	buf.WriteString(fmt.Sprintf(`<a:accent1><a:srgbClr val="%s"/></a:accent1>`, accent1))
	buf.WriteString(fmt.Sprintf(`<a:accent2><a:srgbClr val="%s"/></a:accent2>`, accent2))
	buf.WriteString(`<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>`)
	buf.WriteString(`<a:accent4><a:srgbClr val="FFC000"/></a:accent4>`)
	buf.WriteString(`<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>`)
	buf.WriteString(`<a:accent6><a:srgbClr val="70AD47"/></a:accent6>`)
	buf.WriteString(fmt.Sprintf(`<a:hlink><a:srgbClr val="%s"/></a:hlink>`, hlink))
	buf.WriteString(fmt.Sprintf(`<a:folHlink><a:srgbClr val="%s"/></a:folHlink>`, folHlink))
	buf.WriteString(`</a:clrScheme>`)
	buf.WriteString(`<a:fontScheme name="Office">`)
	buf.WriteString(`<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface="微软雅黑"/><a:cs typeface=""/></a:majorFont>`)
	buf.WriteString(`<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="微软雅黑"/><a:cs typeface=""/></a:minorFont>`)
	buf.WriteString(`</a:fontScheme>`)
	buf.WriteString(`<a:fmtScheme name="Office">`)
	buf.WriteString(`<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>`)
	buf.WriteString(`<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>`)
	buf.WriteString(`<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>`)
	buf.WriteString(`<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>`)
	buf.WriteString(`</a:fmtScheme>`)
	buf.WriteString(`</a:themeElements>`)
	buf.WriteString(`<a:objectDefaults/>`)
	buf.WriteString(`<a:extraClrSchemeLst/>`)
	buf.WriteString(`</a:theme>`)

	_, err = fw.Write(buf.Bytes())
	return err
}

func writeFormattedSlideRels(w io.Writer, imgRels []imageRel, layoutNum int) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)
	buf.WriteString(fmt.Sprintf(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout%d.xml"/>`, layoutNum))
	for _, rel := range imgRels {
		buf.WriteString(fmt.Sprintf(`<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/%s"/>`, rel.relID, rel.filename))
	}
	buf.WriteString(`</Relationships>`)
	_, err := w.Write(buf.Bytes())
	return err
}

func writeFormattedPresentationXML(w io.Writer, slideCount int, slideWidth, slideHeight int32) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" saveSubsetFonts="1">`)
	buf.WriteString(`<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster1"/></p:sldMasterIdLst>`)
	buf.WriteString(`<p:sldIdLst>`)
	for i := 0; i < slideCount; i++ {
		buf.WriteString(fmt.Sprintf(`<p:sldId id="%d" r:id="rIdSlide%d"/>`, 256+i, i+1))
	}
	buf.WriteString(`</p:sldIdLst>`)
	buf.WriteString(fmt.Sprintf(`<p:sldSz cx="%d" cy="%d"/>`, slideWidth, slideHeight))
	buf.WriteString(`<p:notesSz cx="6858000" cy="9144000"/>`)
	// Default text style for consistent text rendering
	buf.WriteString(`<p:defaultTextStyle>`)
	buf.WriteString(`<a:defPPr><a:defRPr lang="zh-CN"/></a:defPPr>`)
	for lvl := 1; lvl <= 5; lvl++ {
		marL := (lvl - 1) * 457200
		buf.WriteString(fmt.Sprintf(`<a:lvl%dpPr marL="%d" algn="l" defTabSz="457200" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">`, lvl, marL))
		buf.WriteString(`<a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>`)
		buf.WriteString(`<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/>`)
		buf.WriteString(fmt.Sprintf(`</a:defRPr></a:lvl%dpPr>`, lvl))
	}
	buf.WriteString(`</p:defaultTextStyle>`)
	buf.WriteString(`</p:presentation>`)
	_, err := w.Write(buf.Bytes())
	return err
}

func writeFormattedPresentationRels(w io.Writer, slideCount int) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)
	buf.WriteString(`<Relationship Id="rIdMaster1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>`)
	for i := 0; i < slideCount; i++ {
		buf.WriteString(fmt.Sprintf(`<Relationship Id="rIdSlide%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide%d.xml"/>`, i+1, i+1))
	}
	buf.WriteString(`<Relationship Id="rIdTheme1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`)
	buf.WriteString(`<Relationship Id="rIdPresProps" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>`)
	buf.WriteString(`<Relationship Id="rIdViewProps" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>`)
	buf.WriteString(`<Relationship Id="rIdTableStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>`)
	buf.WriteString(`</Relationships>`)
	_, err := w.Write(buf.Bytes())
	return err
}

func writeFormattedContentTypes(w io.Writer, slideCount int, imgRels []imageRel, layoutCount int) error {
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
	buf.WriteString(`<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>`)
	buf.WriteString(`<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>`)
	buf.WriteString(`<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>`)
	buf.WriteString(`<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>`)
	buf.WriteString(`<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`)
	for i := 0; i < slideCount; i++ {
		buf.WriteString(fmt.Sprintf(`<Override PartName="/ppt/slides/slide%d.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`, i+1))
	}
	for i := 0; i < layoutCount; i++ {
		buf.WriteString(fmt.Sprintf(`<Override PartName="/ppt/slideLayouts/slideLayout%d.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`, i+1))
	}
	buf.WriteString(`<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>`)
	buf.WriteString(`</Types>`)
	_, err := w.Write(buf.Bytes())
	return err
}
