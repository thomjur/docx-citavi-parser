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
	"regexp"
	"strings"

	"github.com/beevik/etree"
	"github.com/thomjur/verifybibtex/parser"
)

type BibEntry struct {
	Title       string
	Year        string
	CitationKey string // This key is taken from the BibTeX file
	Pages       string
}

func (e *BibEntry) prettyPrint() string {
	return fmt.Sprintf("Title: %s, Year: %s, CitationKey: %s", e.Title, e.Year, e.CitationKey)
}

func (e *BibEntry) toCiteproc() string {
	return fmt.Sprintf("[@%s]", e.CitationKey)
}

func main() {
	// Open ZIP file
	r, err := zip.OpenReader("Krech_Motak.docx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer r.Close()

	// Parsing and loading BibTeX
	bibliography, err := parseBibTeXFile("bib.bib")
	if err != nil {
		log.Fatalln("could not parse bibtex file")
	}
	bibliography.Info()

	// Iterate over all files in ZIP archive
	for _, f := range r.File {
		if f.Name == "word/document.xml" || f.Name == "word/footnotes.xml" {
			fmt.Printf("Found: %s\n", f.Name)
			file, err := f.Open()
			if err != nil {
				fmt.Println(err)
				return
			}
			ParseXML(file, bibliography, f.Name)
			defer file.Close()
		}
	}
}

func ParseXML(f io.Reader, bibliography *parser.BibTeXFile, fileName string) {
	doc := etree.NewDocument()
	citations := 0
	if _, err := doc.ReadFrom(f); err != nil {
		panic(err)
	}
	root := doc.Root()
	fmt.Printf("Root Element: %s\n", root.Tag)
	// Iterating over all sdt elements which are also used by Citavi
	for _, p := range doc.FindElements("//sdt") {
		// Finding instrText elements where bibliographical data is Base64 encoded
		// Some Base64 encodings are split up into several instrText elements
		var sb strings.Builder
		for i, x := range p.FindElements(".//instrText") {
			encodedString := x.Text()
			// If first element does not start with ADDIN Citavi... we can skip
			if !strings.HasPrefix(encodedString, "ADDIN CitaviPlaceholder") && i == 0 {
				break
			}
			// we first need to remove the parts like ADDIN CITAVI etc.
			// as well as the trailing } (if present)
			if strings.HasPrefix(encodedString, "ADDIN CitaviPlaceholder") {
				encodedString = strings.ReplaceAll(encodedString, "ADDIN CitaviPlaceholder{", "")
			}
			// Check if this is final part of string
			if strings.HasSuffix(encodedString, "}") {
				encodedString = strings.TrimSuffix(encodedString, "}")
				sb.WriteString(encodedString)
				break
			}
			sb.WriteString(encodedString)
		}

		// Decoding Base64 encoding
		decodedBytes, err := base64.StdEncoding.DecodeString(sb.String())
		if err != nil {
			fmt.Println("Error when decoding Base64:", err)
			continue
		}

		bibEntryList, c, err := parseEntry(decodedBytes, bibliography)
		if err != nil {
			//fmt.Printf("%e", err)
			continue
		}
		citations += c
		// Creating string that should be added
		contentString := createCitationString(bibEntryList)
		// Create new element with this string in XML tree
		textContainer := p.FindElement("./sdtContent//r/t")
		if textContainer != nil {
			textContainer.SetText(fmt.Sprintf("%s %s", textContainer.Text(), contentString))
		}

	}
	fmt.Printf("We found %d citations.\n", citations)
	// Writing new XML
	if fileName == "word/document.xml" {
		doc.WriteToFile("NEWDOC.xml")
	} else {
		doc.WriteToFile("NEWFN.xml")
	}
}

func createCitationString(l []BibEntry) string {
	var sb strings.Builder
	for _, e := range l {
		sb.WriteString(fmt.Sprintf(" [@%s", e.CitationKey))
		if e.Pages != "" {
			sb.WriteString(fmt.Sprintf(", %s]", e.Pages))
		} else {
			sb.WriteString("]")
		}
	}
	return sb.String()

}

func writeToFileJSON(t []byte, i int) {
	err := os.WriteFile(fmt.Sprintf("json%d.json", i), t, 0644)
	if err != nil {
		log.Fatal(err)
	}
}

func parseEntry(e []byte, bibliography *parser.BibTeXFile) ([]BibEntry, int, error) {
	var jsonData map[string]interface{}
	var bibEntryList []BibEntry

	// Unmarshalling first layer of Citavi JSON
	err := json.Unmarshal(e, &jsonData)
	if err != nil {
		return []BibEntry{}, 0, errors.New("error shalling JSON")
	}

	// Next is Entries list
	var entriesData []map[string]interface{}
	if entries, ok := jsonData["Entries"].([]interface{}); ok {
		for _, entry := range entries {
			if entryMap, ok := entry.(map[string]interface{}); ok {
				entriesData = append(entriesData, entryMap)
			} else {
				return []BibEntry{}, 0, errors.New("error: Entries Map[string]interface{} could not be parsed!")
			}
		}
	} else {
		return []BibEntry{}, 0, errors.New("error: jsonData['Entries'] could not be parsed as []interface{}!")
	}

	// Next we need to get Reference part of each entry
	// we also get PageRange data
	var referencesData []map[string]interface{}
	var pageRangeData []map[string]interface{}
	for _, entry := range entriesData {
		if referencesMap, ok := entry["Reference"].(map[string]interface{}); ok {
			referencesData = append(referencesData, referencesMap)
		} else {
			return []BibEntry{}, 0, errors.New("error: Could not find references")
		}
		if pageRange, ok := entry["PageRange"].(map[string]interface{}); ok {
			pageRangeData = append(pageRangeData, pageRange)
		}
	}

	// Try parsing titles, years, and citation key from BibTeX file
	citations := 0
	for _, ref := range referencesData {
		bibEntry := BibEntry{}
		if title, ok := ref["Title"].(string); ok {
			bibEntry.Title = title
		}
		// Check if subtitle is present and add to main title in this case"
		if subtitle, ok := ref["Subtitle"].(string); ok {
			bibEntry.Title = fmt.Sprintf("%s %s", bibEntry.Title, subtitle)
		}
		if year, ok := ref["Year"].(string); ok {
			bibEntry.Year = year
		}
		err := findCitationKey(&bibEntry, bibliography)
		if err != nil {
			//log.Println(err)
		} else {
			citations++
		}
		bibEntryList = append(bibEntryList, bibEntry)
	}

	// Finally, we add page numbers
	for i, pages := range pageRangeData {
		if startPage, ok := pages["StartPage"].(map[string]interface{}); ok {
			if pageRange, ok := startPage["OriginalString"].(string); ok {
				bibEntryList[i].Pages = pageRange
			}
		}
		if endPage, ok := pages["EndPage"].(map[string]interface{}); ok {
			if pageRange, ok := endPage["OriginalString"].(string); ok {
				if pageRange != bibEntryList[i].Pages {
					bibEntryList[i].Pages = fmt.Sprintf("%s-%s", bibEntryList[i].Pages, pageRange)
				}
			}
		}
	}

	return bibEntryList, citations, nil
}

// Helper function to clean literature titles
func cleanTitle(t string) string {
	t = strings.ToLower(t)
	// First, we replace all ä ö ü ß
	t = strings.ReplaceAll(t, "ä", "a")
	t = strings.ReplaceAll(t, "ö", "o")
	t = strings.ReplaceAll(t, "ü", "u")
	t = strings.ReplaceAll(t, "ß", "ss")

	// Removing all non alphanumerical signs
	re := regexp.MustCompile(`\W`)
	t = re.ReplaceAllString(t, "")
	return t
}

func findCitationKey(bibEntry *BibEntry, bibliography *parser.BibTeXFile) error {
	// Iterating over all entries in BibTeX file in hope to find a matching entry
	// TODO: Currently only comparing complete titles to find matching key
	found := false
	cleanTitleBib := cleanTitle(bibEntry.Title)
	for _, e := range bibliography.Entries {
		if title, exists := e.Fields["title"]; exists {
			cleanTitleEntry := cleanTitle(title)
			if cleanTitleBib == cleanTitleEntry {
				if e.Key != "" {
					bibEntry.CitationKey = e.Key
					found = true
					break
				}
			}
		}
	}
	if !found {
		return errors.New(fmt.Sprintf("could not find any matching entries for %s in BibTeX file which is clean: %s", bibEntry.Title, cleanTitleBib))
	} else {
		return nil
	}
}

func parseBibTeXFile(path string) (*parser.BibTeXFile, error) {
	BibTeXFilePath := "bib.bib"
	// Trying to open the file
	file, err := os.Open(path)
	if err != nil {
		fmt.Println("Something went terribly wrong :-(")
	}
	defer file.Close()

	bibtexData, err2 := parser.ParseNewBibTeXFile(file)
	if err2 != nil {
		fmt.Println("Something went terribly wrong :-(")
	}
	// Don't forget to add filename afterwards
	bibtexData.FilePath = BibTeXFilePath
	return bibtexData, nil
}
