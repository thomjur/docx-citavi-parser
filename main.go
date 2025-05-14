package main

import (
	"archive/zip"
	"fmt"
	"github.com/beevik/etree"
	"io"
)

func main() {
	// Open ZIP file
	r, err := zip.OpenReader("document.docx")
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
	if _, err := doc.ReadFrom(f); err != nil {
		panic(err)
	}
	root := doc.Root()
	fmt.Printf("Root Element: %s\n", root.Tag)
	// Testing: Iterating over all paragraphs
	for _, p := range doc.FindElements("//p") {
		for _, c := range p.ChildElements() {
			for _, t := range c.FindElements("//t") {
				fmt.Printf("%s\n", t.Text())
			}
		}
	}
}
