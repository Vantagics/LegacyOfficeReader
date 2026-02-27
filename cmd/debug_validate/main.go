package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Println("Error opening:", err)
		return
	}
	defer r.Close()

	errors := 0
	for _, f := range r.File {
		if !strings.HasSuffix(f.Name, ".xml") && !strings.HasSuffix(f.Name, ".rels") {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			fmt.Printf("ERROR opening %s: %v\n", f.Name, err)
			errors++
			continue
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			fmt.Printf("ERROR reading %s: %v\n", f.Name, err)
			errors++
			continue
		}
		
		// Validate XML
		decoder := xml.NewDecoder(strings.NewReader(string(data)))
		for {
			_, err := decoder.Token()
			if err != nil {
				if err.Error() == "EOF" {
					break
				}
				fmt.Printf("XML ERROR in %s: %v\n", f.Name, err)
				errors++
				break
			}
		}
	}
	
	if errors == 0 {
		fmt.Println("All XML files are valid!")
	} else {
		fmt.Printf("%d errors found\n", errors)
	}
}
