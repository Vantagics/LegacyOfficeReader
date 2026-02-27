package xlsconv

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strconv"

	"github.com/shakinm/xlsReader/xls"
	"github.com/shakinm/xlsReader/xls/structure"
)

// cellValue represents a mapped cell value for XLSX generation.
type cellValue struct {
	IsNumeric bool
	NumVal    float64
	StrVal    string
}

// sheetData is an internal representation of a worksheet for XLSX generation.
type sheetData struct {
	Name string
	Rows [][]cellValue // rows[rowIdx][colIdx]; nil cellValue means skip
}

// mapCell maps a CellData to a cellValue based on its type.
// Returns nil if the cell should be skipped (blank types).
func mapCell(cd structure.CellData) *cellValue {
	switch cd.GetType() {
	case "*record.Number", "*record.Rk":
		return &cellValue{IsNumeric: true, NumVal: cd.GetFloat64()}
	case "*record.LabelSSt", "*record.LabelBIFF8", "*record.LabelBIFF5", "*record.BoolErr":
		return &cellValue{IsNumeric: false, StrVal: cd.GetString()}
	case "*record.Blank", "*record.FakeBlank":
		return nil
	default:
		// Unknown types: try string representation
		s := cd.GetString()
		if s != "" {
			return &cellValue{IsNumeric: false, StrVal: s}
		}
		return nil
	}
}

// mapWorkbook extracts sheet data from a parsed Workbook.
func mapWorkbook(wb *xls.Workbook) []sheetData {
	sheets := wb.GetSheets()
	result := make([]sheetData, len(sheets))
	for i, s := range sheets {
		result[i].Name = s.GetName()
		rows := s.GetRows()
		result[i].Rows = make([][]cellValue, len(rows))
		for ri, row := range rows {
			cols := row.GetCols()
			mapped := make([]cellValue, len(cols))
			for ci, cd := range cols {
				cv := mapCell(cd)
				if cv != nil {
					mapped[ci] = *cv
				}
				// zero-value cellValue (IsNumeric=false, StrVal="") means skip
			}
			result[i].Rows[ri] = mapped
		}
	}
	return result
}

// ConvertReader reads XLS data from reader, converts it to XLSX, and writes to writer.
func ConvertReader(reader io.ReadSeeker, writer io.Writer) error {
	wb, err := xls.OpenReader(reader)
	if err != nil {
		return fmt.Errorf("xlsconv: failed to parse input: %w", err)
	}

	sheets := mapWorkbook(&wb)

	if err := writeXlsx(writer, sheets); err != nil {
		return fmt.Errorf("xlsconv: failed to write xlsx: %w", err)
	}
	return nil
}

// ConvertFile converts an XLS file at inputPath to an XLSX file at outputPath.
func ConvertFile(inputPath string, outputPath string) error {
	inFile, err := os.Open(inputPath)
	if err != nil {
		return fmt.Errorf("xlsconv: failed to open input file: %w", err)
	}
	defer inFile.Close()

	outFile, err := os.Create(outputPath)
	if err != nil {
		return fmt.Errorf("xlsconv: failed to create output file: %w", err)
	}
	defer outFile.Close()

	return ConvertReader(inFile, outFile)
}

// colName converts a 0-based column index to an Excel column letter (A, B, ..., Z, AA, ...).
func colName(col int) string {
	name := ""
	for {
		name = string(rune('A'+col%26)) + name
		col = col/26 - 1
		if col < 0 {
			break
		}
	}
	return name
}

// cellRef returns an Excel cell reference like "A1", "B2" for 0-based row and col.
func cellRef(row, col int) string {
	return colName(col) + strconv.Itoa(row+1)
}

// writeXlsx generates a minimal valid XLSX (Office Open XML) zip archive.
func writeXlsx(w io.Writer, sheets []sheetData) error {
	zw := zip.NewWriter(w)
	defer zw.Close()

	// Collect all shared strings
	var sharedStrings []string
	ssIndex := map[string]int{}

	for _, sheet := range sheets {
		for _, row := range sheet.Rows {
			for _, cell := range row {
				if !cell.IsNumeric && cell.StrVal != "" {
					if _, ok := ssIndex[cell.StrVal]; !ok {
						ssIndex[cell.StrVal] = len(sharedStrings)
						sharedStrings = append(sharedStrings, cell.StrVal)
					}
				}
			}
		}
	}

	// Write sheet XML files
	for i, sheet := range sheets {
		fw, err := zw.Create(fmt.Sprintf("xl/worksheets/sheet%d.xml", i+1))
		if err != nil {
			return err
		}
		if err := writeSheetXML(fw, sheet, ssIndex); err != nil {
			return err
		}
	}

	// If no sheets, create one empty sheet
	if len(sheets) == 0 {
		fw, err := zw.Create("xl/worksheets/sheet1.xml")
		if err != nil {
			return err
		}
		if err := writeSheetXML(fw, sheetData{Name: "Sheet1"}, ssIndex); err != nil {
			return err
		}
	}

	// Write shared strings
	fw, err := zw.Create("xl/sharedStrings.xml")
	if err != nil {
		return err
	}
	if err := writeSharedStringsXML(fw, sharedStrings); err != nil {
		return err
	}

	// Write styles
	fw, err = zw.Create("xl/styles.xml")
	if err != nil {
		return err
	}
	if err := writeStylesXML(fw); err != nil {
		return err
	}

	// Write workbook.xml
	fw, err = zw.Create("xl/workbook.xml")
	if err != nil {
		return err
	}
	if err := writeWorkbookXML(fw, sheets); err != nil {
		return err
	}

	// Write xl/_rels/workbook.xml.rels
	fw, err = zw.Create("xl/_rels/workbook.xml.rels")
	if err != nil {
		return err
	}
	if err := writeWorkbookRels(fw, sheets); err != nil {
		return err
	}

	// Write [Content_Types].xml
	fw, err = zw.Create("[Content_Types].xml")
	if err != nil {
		return err
	}
	if err := writeContentTypes(fw, sheets); err != nil {
		return err
	}

	// Write _rels/.rels
	fw, err = zw.Create("_rels/.rels")
	if err != nil {
		return err
	}
	if _, err := io.WriteString(fw, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`+
		`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>`+
		`</Relationships>`); err != nil {
		return err
	}

	return nil
}

func writeSheetXML(w io.Writer, sheet sheetData, ssIndex map[string]int) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`)
	buf.WriteString(`<sheetData>`)

	for ri, row := range sheet.Rows {
		buf.WriteString(fmt.Sprintf(`<row r="%d">`, ri+1))
		for ci, cell := range row {
			ref := cellRef(ri, ci)
			if cell.IsNumeric {
				buf.WriteString(fmt.Sprintf(`<c r="%s"><v>%s</v></c>`, ref, strconv.FormatFloat(cell.NumVal, 'f', -1, 64)))
			} else if cell.StrVal != "" {
				idx := ssIndex[cell.StrVal]
				buf.WriteString(fmt.Sprintf(`<c r="%s" t="s"><v>%d</v></c>`, ref, idx))
			}
			// skip empty cells (blank/fakeblank)
		}
		buf.WriteString(`</row>`)
	}

	buf.WriteString(`</sheetData>`)
	buf.WriteString(`</worksheet>`)
	_, err := w.Write(buf.Bytes())
	return err
}

func writeSharedStringsXML(w io.Writer, strings []string) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(fmt.Sprintf(`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="%d" uniqueCount="%d">`, len(strings), len(strings)))
	for _, s := range strings {
		buf.WriteString(`<si><t>`)
		xml.Escape(&buf, []byte(s))
		buf.WriteString(`</t></si>`)
	}
	buf.WriteString(`</sst>`)
	_, err := w.Write(buf.Bytes())
	return err
}

func writeStylesXML(w io.Writer) error {
	_, err := io.WriteString(w, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`+
		`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`+
		`<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>`+
		`<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>`+
		`<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>`+
		`<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>`+
		`<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>`+
		`</styleSheet>`)
	return err
}

func writeWorkbookXML(w io.Writer, sheets []sheetData) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`)
	buf.WriteString(`<sheets>`)

	sheetCount := len(sheets)
	if sheetCount == 0 {
		sheetCount = 1
	}
	for i := 0; i < sheetCount; i++ {
		name := "Sheet1"
		if i < len(sheets) {
			name = sheets[i].Name
		}
		var escaped bytes.Buffer
		xml.Escape(&escaped, []byte(name))
		buf.WriteString(fmt.Sprintf(`<sheet name="%s" sheetId="%d" r:id="rIdSheet%d"/>`, escaped.String(), i+1, i+1))
	}

	buf.WriteString(`</sheets>`)
	buf.WriteString(`</workbook>`)
	_, err := w.Write(buf.Bytes())
	return err
}

func writeWorkbookRels(w io.Writer, sheets []sheetData) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)

	sheetCount := len(sheets)
	if sheetCount == 0 {
		sheetCount = 1
	}
	for i := 0; i < sheetCount; i++ {
		buf.WriteString(fmt.Sprintf(`<Relationship Id="rIdSheet%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>`, i+1, i+1))
	}

	buf.WriteString(`<Relationship Id="rIdStyles1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`)
	buf.WriteString(`<Relationship Id="rIdSS1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`)
	buf.WriteString(`</Relationships>`)
	_, err := w.Write(buf.Bytes())
	return err
}

func writeContentTypes(w io.Writer, sheets []sheetData) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`)
	buf.WriteString(`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`)
	buf.WriteString(`<Default Extension="xml" ContentType="application/xml"/>`)
	buf.WriteString(`<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`)

	sheetCount := len(sheets)
	if sheetCount == 0 {
		sheetCount = 1
	}
	for i := 0; i < sheetCount; i++ {
		buf.WriteString(fmt.Sprintf(`<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`, i+1))
	}

	buf.WriteString(`<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`)
	buf.WriteString(`<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`)
	buf.WriteString(`</Types>`)
	_, err := w.Write(buf.Bytes())
	return err
}
