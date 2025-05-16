package main

import (
	"archive/zip"
	"encoding/base64"
	"fmt"
	"io"
	"strings"

	"github.com/beevik/etree"
)

func main() {
	// Open ZIP file
	r, err := zip.OpenReader("Krech_Motak.docx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer r.Close()

	// Iterate over all files in ZIP archive
	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			fmt.Printf("Found: %s\n", f.Name)
			file, err := f.Open()
			if err != nil {
				fmt.Println(err)
				return
			}
			ParseXML(file)
			defer file.Close()
		}
	}
}

func ParseXML(f io.Reader) {
	doc := etree.NewDocument()
	citations := 0
	if _, err := doc.ReadFrom(f); err != nil {
		panic(err)
	}
	root := doc.Root()
	fmt.Printf("Root Element: %s\n", root.Tag)
	// Testing: Iterating over all paragraphs
	for _, p := range doc.FindElements("//sdt") {
		//if i == 4 {
		//	break
		//}
		for _, x := range p.FindElements(".//instrText") {
			encodedString := x.Text()
			if !strings.HasPrefix(encodedString, "ADDIN CitaviPlaceholder") {
				continue
			}
			// we first need to remove the parts like ADDIN CITAVI etc.
			// as well as the trailing }
			cleanEncodedString := strings.ReplaceAll(encodedString, "ADDIN CitaviPlaceholder{", "")
			cleanEncodedString = strings.TrimSuffix(cleanEncodedString, "}")
			//	fmt.Printf("%s", cleanEncodedString)

			_, err := base64.StdEncoding.DecodeString(cleanEncodedString)
			if err != nil {
				citations++
				fmt.Printf("%s\n\n", cleanEncodedString)
				fmt.Printf("%s\n\n", encodedString)
				fmt.Printf("%v\n", x)
				fmt.Println("Fehler beim Dekodieren:", err)
			}

			//decodedString := string(decodedBytes)
			//	fmt.Println("Dekodierter String:", decodedString)
		}
	}
	fmt.Printf("We found %d citations.", citations)
}
