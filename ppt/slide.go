package ppt

// Slide represents a single slide in a PPT presentation.
type Slide struct {
	texts             []string
	shapes            []ShapeFormatting
	layoutType        int
	background        SlideBackground
	masterRef         uint32   // masterIdRef from SlideAtom (persist ref to MainMasterContainer)
	colorScheme       []string // resolved from master
	defaultTextStyles [5]MasterTextStyle // from master
	textTypeStyles    map[int][5]MasterTextStyle // from master, keyed by text type
}

// GetTexts returns all text blocks in this slide.
func (s *Slide) GetTexts() []string {
	return s.texts
}

// GetShapes returns all shapes with formatting information in this slide.
func (s *Slide) GetShapes() []ShapeFormatting {
	return s.shapes
}

// GetLayoutType returns the slide layout type identifier.
func (s *Slide) GetLayoutType() int {
	return s.layoutType
}

// GetBackground returns the slide background fill information.
func (s *Slide) GetBackground() SlideBackground {
	return s.background
}

// GetMasterRef returns the masterIdRef from the SlideAtom.
func (s *Slide) GetMasterRef() uint32 {
	return s.masterRef
}

// GetColorScheme returns the color scheme inherited from the master.
func (s *Slide) GetColorScheme() []string {
	return s.colorScheme
}


// GetDefaultTextStyles returns the master's default text styles per indent level.
func (s *Slide) GetDefaultTextStyles() [5]MasterTextStyle {
	return s.defaultTextStyles
}

// GetTextTypeStyles returns the master's text styles per text type.
func (s *Slide) GetTextTypeStyles() map[int][5]MasterTextStyle {
	return s.textTypeStyles
}
