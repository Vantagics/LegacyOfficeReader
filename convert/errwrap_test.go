package convert_test

import (
	"bytes"
	"errors"
	"math/rand"
	"strings"
	"testing"
	"testing/quick"

	"github.com/shakinm/xlsReader/convert/docconv"
	"github.com/shakinm/xlsReader/convert/pptconv"
	"github.com/shakinm/xlsReader/convert/xlsconv"
)

// converterEntry pairs a converter's ConvertReader function with its expected prefix.
type converterEntry struct {
	name    string
	prefix  string
	convert func(r *bytes.Reader, w *bytes.Buffer) error
}

var converters = []converterEntry{
	{
		name:   "pptconv",
		prefix: "pptconv",
		convert: func(r *bytes.Reader, w *bytes.Buffer) error {
			return pptconv.ConvertReader(r, w)
		},
	},
	{
		name:   "xlsconv",
		prefix: "xlsconv",
		convert: func(r *bytes.Reader, w *bytes.Buffer) error {
			return xlsconv.ConvertReader(r, w)
		},
	},
	{
		name:   "docconv",
		prefix: "docconv",
		convert: func(r *bytes.Reader, w *bytes.Buffer) error {
			return docconv.ConvertReader(r, w)
		},
	},
}

// Feature: legacy-to-ooxml-conversion, Property 7: 错误包装含正确前缀
// **Validates: Requirements 6.1, 6.2, 6.3, 6.4, 6.5, 6.6**
//
// For any converter (pptconv, xlsconv, docconv) and any random invalid input,
// the returned error message should contain the corresponding package prefix,
// and the original error should be recoverable via errors.Unwrap.
func TestProperty_ErrorWrapPrefix(t *testing.T) {
	config := &quick.Config{MaxCount: 100}

	for _, conv := range converters {
		conv := conv // capture range variable
		t.Run(conv.name+"_ConvertReader", func(t *testing.T) {
			prop := func(seed int64) bool {
				rng := rand.New(rand.NewSource(seed))

				// Generate random bytes as invalid input (1-256 bytes)
				dataLen := 1 + rng.Intn(256)
				data := make([]byte, dataLen)
				for i := range data {
					data[i] = byte(rng.Intn(256))
				}

				reader := bytes.NewReader(data)
				var output bytes.Buffer

				err := conv.convert(reader, &output)
				if err == nil {
					// Random bytes happened to parse successfully — skip this case
					return true
				}

				// Verify error message contains the correct prefix
				if !strings.Contains(err.Error(), conv.prefix) {
					t.Logf("%s: error %q does not contain prefix %q", conv.name, err.Error(), conv.prefix)
					return false
				}

				// Verify the original error can be unwrapped
				unwrapped := errors.Unwrap(err)
				if unwrapped == nil {
					t.Logf("%s: errors.Unwrap returned nil for error %q", conv.name, err.Error())
					return false
				}

				return true
			}

			if err := quick.Check(prop, config); err != nil {
				t.Errorf("Property failed: %s error wrap prefix: %v", conv.name, err)
			}
		})

		t.Run(conv.name+"_ConvertFile", func(t *testing.T) {
			prop := func(seed int64) bool {
				rng := rand.New(rand.NewSource(seed))

				// Generate a random nonexistent path
				pathLen := 5 + rng.Intn(20)
				path := "/nonexistent/" + errwrapRandomString(rng, pathLen) + ".tmp"

				var err error
				switch conv.name {
				case "pptconv":
					err = pptconv.ConvertFile(path, "/tmp/out.pptx")
				case "xlsconv":
					err = xlsconv.ConvertFile(path, "/tmp/out.xlsx")
				case "docconv":
					err = docconv.ConvertFile(path, "/tmp/out.docx")
				}

				if err == nil {
					t.Logf("%s: expected error for nonexistent path %q, got nil", conv.name, path)
					return false
				}

				// Verify error message contains the correct prefix
				if !strings.Contains(err.Error(), conv.prefix) {
					t.Logf("%s: error %q does not contain prefix %q", conv.name, err.Error(), conv.prefix)
					return false
				}

				// Verify the original error can be unwrapped
				unwrapped := errors.Unwrap(err)
				if unwrapped == nil {
					t.Logf("%s: errors.Unwrap returned nil for error %q", conv.name, err.Error())
					return false
				}

				return true
			}

			if err := quick.Check(prop, config); err != nil {
				t.Errorf("Property failed: %s ConvertFile error wrap prefix: %v", conv.name, err)
			}
		})
	}
}

// errwrapRandomString generates a random alphanumeric string of the given length.
func errwrapRandomString(rng *rand.Rand, length int) string {
	const charset = "abcdefghijklmnopqrstuvwxyz0123456789"
	b := make([]byte, length)
	for i := range b {
		b[i] = charset[rng.Intn(len(charset))]
	}
	return string(b)
}
