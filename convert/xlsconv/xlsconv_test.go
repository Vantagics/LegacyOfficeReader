package xlsconv

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"math/rand"
	"strings"
	"testing"
	"testing/quick"
)

func TestConvertReader_InvalidInput(t *testing.T) {
	// Provide invalid (non-XLS) data
	input := bytes.NewReader([]byte("this is not a valid XLS file"))
	var output bytes.Buffer

	err := ConvertReader(input, &output)
	if err == nil {
		t.Fatal("expected error for invalid input, got nil")
	}
	if !strings.Contains(err.Error(), "xlsconv") {
		t.Errorf("error message should contain 'xlsconv' prefix, got: %s", err.Error())
	}
}

func TestConvertFile_NonexistentInput(t *testing.T) {
	err := ConvertFile("/nonexistent/path/to/file.xls", "/tmp/output.xlsx")
	if err == nil {
		t.Fatal("expected error for nonexistent input file, got nil")
	}
	if !strings.Contains(err.Error(), "xlsconv") {
		t.Errorf("error message should contain 'xlsconv' prefix, got: %s", err.Error())
	}
}

func TestConvertFile_InvalidOutputPath(t *testing.T) {
	err := ConvertFile("/nonexistent/input.xls", "/nonexistent/dir/output.xlsx")
	if err == nil {
		t.Fatal("expected error, got nil")
	}
	if !strings.Contains(err.Error(), "xlsconv") {
		t.Errorf("error message should contain 'xlsconv' prefix, got: %s", err.Error())
	}
}

// Feature: legacy-to-ooxml-conversion, Property 3: XLS 工作表名称保留
// **Validates: Requirements 2.3**
//
// For any set of worksheet names, after mapping through writeXlsx,
// the output XLSX workbook.xml should contain all sheet names in order
// and the count should match.
func TestProperty_SheetNamePreservation(t *testing.T) {
	config := &quick.Config{MaxCount: 100}

	prop := func(sheetCount uint8, seed int64) bool {
		// Limit sheet count to 1-5
		numSheets := int(sheetCount)%5 + 1
		rng := rand.New(rand.NewSource(seed))

		// Generate random sheet names (1-20 alphanumeric chars each)
		sheets := make([]sheetData, numSheets)
		expectedNames := make([]string, numSheets)
		for i := 0; i < numSheets; i++ {
			nameLen := 1 + rng.Intn(20)
			name := xlsRandomString(rng, nameLen)
			sheets[i] = sheetData{Name: name, Rows: nil}
			expectedNames[i] = name
		}

		// Write XLSX to buffer
		var buf bytes.Buffer
		if err := writeXlsx(&buf, sheets); err != nil {
			t.Logf("writeXlsx failed: %v", err)
			return false
		}

		// Open the zip and read xl/workbook.xml
		reader := bytes.NewReader(buf.Bytes())
		zr, err := zip.NewReader(reader, int64(buf.Len()))
		if err != nil {
			t.Logf("failed to open zip: %v", err)
			return false
		}

		wbXML, err := xlsReadZipFile(zr, "xl/workbook.xml")
		if err != nil {
			t.Logf("failed to read workbook.xml: %v", err)
			return false
		}

		// Parse sheet names from workbook.xml
		parsedNames, err := parseSheetNames(wbXML)
		if err != nil {
			t.Logf("failed to parse sheet names: %v", err)
			return false
		}

		// Verify count matches
		if len(parsedNames) != numSheets {
			t.Logf("expected %d sheets, got %d", numSheets, len(parsedNames))
			return false
		}

		// Verify names are in order
		for i, expected := range expectedNames {
			if parsedNames[i] != expected {
				t.Logf("sheet %d: expected name %q, got %q", i, expected, parsedNames[i])
				return false
			}
		}

		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property failed: XLS sheet name preservation: %v", err)
	}
}

// xlsRandomString generates a random alphanumeric string of the given length.
func xlsRandomString(rng *rand.Rand, length int) string {
	const charset = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
	b := make([]byte, length)
	for i := range b {
		b[i] = charset[rng.Intn(len(charset))]
	}
	return string(b)
}

// xlsReadZipFile reads the content of a file inside a zip archive.
func xlsReadZipFile(zr *zip.Reader, name string) (string, error) {
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return "", err
			}
			defer rc.Close()
			var buf bytes.Buffer
			if _, err := io.Copy(&buf, rc); err != nil {
				return "", err
			}
			return buf.String(), nil
		}
	}
	return "", fmt.Errorf("file %s not found in zip", name)
}

// parseSheetNames extracts sheet names from workbook.xml content in order.
func parseSheetNames(xmlContent string) ([]string, error) {
	type Sheet struct {
		Name string `xml:"name,attr"`
	}
	type Sheets struct {
		Sheet []Sheet `xml:"sheets>sheet"`
	}
	var wb Sheets
	if err := xml.Unmarshal([]byte(xmlContent), &wb); err != nil {
		return nil, err
	}
	names := make([]string, len(wb.Sheet))
	for i, s := range wb.Sheet {
		names[i] = s.Name
	}
	return names, nil
}

// mockCellData implements structure.CellData for property testing.
type mockCellData struct {
	typ      string
	strVal   string
	floatVal float64
}

func (m mockCellData) GetString() string   { return m.strVal }
func (m mockCellData) GetFloat64() float64 { return m.floatVal }
func (m mockCellData) GetInt64() int64     { return int64(m.floatVal) }
func (m mockCellData) GetXFIndex() int     { return 0 }
func (m mockCellData) GetType() string     { return m.typ }

// Feature: legacy-to-ooxml-conversion, Property 4: XLS 单元格数据类型正确映射
// **Validates: Requirements 2.4, 2.5, 5.1, 5.2, 5.3**
//
// For any CellData with a known type, mapCell should correctly map:
// - Numeric types → cellValue with IsNumeric=true and correct NumVal
// - String types → cellValue with IsNumeric=false and correct StrVal
// - Blank types → nil
func TestProperty_CellDataTypeMapping(t *testing.T) {
	numericTypes := []string{"*record.Number", "*record.Rk"}
	stringTypes := []string{"*record.LabelSSt", "*record.LabelBIFF8", "*record.LabelBIFF5", "*record.BoolErr"}
	blankTypes := []string{"*record.Blank", "*record.FakeBlank"}
	allTypes := make([]string, 0, len(numericTypes)+len(stringTypes)+len(blankTypes))
	allTypes = append(allTypes, numericTypes...)
	allTypes = append(allTypes, stringTypes...)
	allTypes = append(allTypes, blankTypes...)

	config := &quick.Config{MaxCount: 100}

	prop := func(typeIdx uint8, seed int64) bool {
		rng := rand.New(rand.NewSource(seed))

		// Pick a type from the known types
		idx := int(typeIdx) % len(allTypes)
		typ := allTypes[idx]

		// Generate random values
		strVal := xlsRandomString(rng, 1+rng.Intn(50))
		floatVal := rng.Float64()*2000 - 1000 // range [-1000, 1000)

		cell := mockCellData{typ: typ, strVal: strVal, floatVal: floatVal}
		result := mapCell(cell)

		// Check based on type category
		isNumeric := typ == "*record.Number" || typ == "*record.Rk"
		isString := typ == "*record.LabelSSt" || typ == "*record.LabelBIFF8" ||
			typ == "*record.LabelBIFF5" || typ == "*record.BoolErr"
		isBlank := typ == "*record.Blank" || typ == "*record.FakeBlank"

		if isBlank {
			if result != nil {
				t.Logf("type %q: expected nil for blank, got %+v", typ, result)
				return false
			}
			return true
		}

		if result == nil {
			t.Logf("type %q: expected non-nil result, got nil", typ)
			return false
		}

		if isNumeric {
			if !result.IsNumeric {
				t.Logf("type %q: expected IsNumeric=true, got false", typ)
				return false
			}
			if result.NumVal != floatVal {
				t.Logf("type %q: expected NumVal=%f, got %f", typ, floatVal, result.NumVal)
				return false
			}
		}

		if isString {
			if result.IsNumeric {
				t.Logf("type %q: expected IsNumeric=false, got true", typ)
				return false
			}
			if result.StrVal != strVal {
				t.Logf("type %q: expected StrVal=%q, got %q", typ, strVal, result.StrVal)
				return false
			}
		}

		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property failed: XLS cell data type mapping: %v", err)
	}
}
