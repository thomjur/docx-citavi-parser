package main

import (
	"archive/zip"
	"encoding/base64"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"log"
	"os"
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
	for i, p := range doc.FindElements("//sdt") {
		if i == 8 {
			break
		}
		// TODO: there might references where Base64 part is split up into multiple
		// instrText elements
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

			decodedBytes, err := base64.StdEncoding.DecodeString(cleanEncodedString)
			if err != nil {
				fmt.Println("Fehler beim Dekodieren:", err)
				continue
			}

			//decodedString := string(decodedBytes)

			titleStr, err := parseTitleFromEntry(decodedBytes)
			if err != nil {
				log.Fatalf("%e", err)
			}
			fmt.Println(titleStr)
			// fmt.Println("Dekodierter String:", decodedString)
			// writeToFile([]byte(decodedString), i)
		}
	}
	fmt.Printf("We found %d citations.", citations)
}

func writeToFileJSON(t []byte, i int) {
	err := os.WriteFile(fmt.Sprintf("json%d.json", i), t, 0644)
	if err != nil {
		log.Fatal(err)
	}
}

func parseTitleFromEntry(e []byte) (string, error) {
	var jsonData map[string]interface{}
	var sb strings.Builder
	err := json.Unmarshal(e, &jsonData)
	if err != nil {
		return "", errors.New("error when unmarshalling JSON")
	}
	// next is Entries list
	var entriesData []map[string]interface{}
	if entries, ok := jsonData["Entries"].([]interface{}); ok { // Korrektur: Assertion auf []interface{}
		for _, entry := range entries {
			if entryMap, ok := entry.(map[string]interface{}); ok {
				entriesData = append(entriesData, entryMap)
			} else {
				return "", errors.New("error: Entries  Map[string]interface{}")
			}
		}
	} else {
		return "", errors.New("error: jsonData['Entries'] ist kein []interface{}")
	}
	// next we need to get Reference part of entry
	var referencesData []map[string]interface{}
	for _, entry := range entriesData {
		if referencesMap, ok := entry["Reference"].(map[string]interface{}); ok {
			referencesData = append(referencesData, referencesMap)
		} else {
			return "", errors.New("error: Could not find references")
		}
	}
	// try parse titles
	for _, ref := range referencesData {
		if title, ok := ref["Title"].(string); ok {
			sb.WriteString(fmt.Sprintf(" %s ||", title))
		}
	}

	return sb.String(), nil
}
