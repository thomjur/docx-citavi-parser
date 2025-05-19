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
	// Parsing cli args
	args := os.Args[1:]

	if len(args) != 2 {
		log.Fatalln("The DOCX Citavi Parser expects two arguments: (1) Path to DOCX file (2) Path to BibTeX file.")
	}
	// Open ZIP/DOCX file
	r, err := zip.OpenReader(args[0])
	if err != nil {
		log.Fatalln(err)
	}
	defer r.Close()

	// Parsing and loading BibTeX
	bibliography, err := parseBibTeXFile("bib.bib")
	if err != nil {
		log.Fatalln("Could not parse BibTeX file.")
	}

	// Iterating over all files in ZIP archive
	// To find document.xml and footnotes.xml to check for refs
	for _, f := range r.File {
		if f.Name == "word/document.xml" || f.Name == "word/footnotes.xml" {
			fmt.Printf("Found: %s\n", f.Name)
			file, err := f.Open()
			if err != nil {
				fmt.Println(err)
				return
			}
			parseXML(file, bibliography, f.Name)
			defer file.Close()
		}
	}
}

// Helper function to clean literature titles
// to make them comparable between Citavi entries and BibTeX file
// Issue is that BibTeX often includes transcriptions such as \"o for ö etc.}
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

// Helper function to create Citeproc citation like strings
// e.g., [@<AUTHOR>, <PAGES>]
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

// Function to find citation key in a BibTeX file based on the title of the
// entry in Citavi embedded citation in DOCX file
// TODO: Currently only matching complete titles (incl. subtitles) to find matching key
func findCitationKey(bibEntry *BibEntry, bibliography *parser.BibTeXFile) error {
	// Iterating over all entries in BibTeX file in hope to find a matching entry
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

// Function to parse BibTeX file using thomjur/verifybibtex/parser library
func parseBibTeXFile(path string) (*parser.BibTeXFile, error) {
	// Trying to open the file
	file, err := os.Open(path)
	if err != nil {
		fmt.Println("Something went terribly wrong :-(")
	}
	defer file.Close()

	bibtexData, err := parser.ParseNewBibTeXFile(file)
	if err != nil {
		fmt.Println("Something went terribly wrong :-(")
	}
	// Don't forget to add filename afterwards
	bibtexData.FilePath = path
	return bibtexData, nil
}

// Function to parse a Citavi Citation entry in DOCX
// Returns a list of BibEntry objects that are cited in this parseEntry plus the
// number of citations found
func parseEntry(e []byte, bibliography *parser.BibTeXFile) ([]BibEntry, int, error) {
	var jsonData map[string]interface{}
	var bibEntryList []BibEntry

	// Since the Citavi JSON is too complex and I don't have a corresponding
	// struct, I just parse certain paths in the JSON objects

	// Unmarshalling first layer of Citavi JSON
	err := json.Unmarshal(e, &jsonData)
	if err != nil {
		return []BibEntry{}, 0, errors.New("Error unmarshalling JSON")
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

// Main function to parse the document.xml or footnotes.xml
// We are using the etree library here
func parseXML(f io.Reader, bibliography *parser.BibTeXFile, fileName string) {
	doc := etree.NewDocument()
	citations := 0
	if _, err := doc.ReadFrom(f); err != nil {
		panic(err)
	}

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

		// If no valid Base64 encoding in Citavi Entry was found, we can stop here
		if sb.String() == "" {
			continue
		}

		// Decoding Base64 encoding
		decodedBytes, err := base64.StdEncoding.DecodeString(sb.String())
		if err != nil {
			fmt.Println("Error when decoding Base64:", err)
			continue
		}

		// Parsing citations
		bibEntryList, c, err := parseEntry(decodedBytes, bibliography)
		if err != nil {
			//fmt.Printf("%e", err)
			continue
		}
		citations += c
		// Creating string that should be added to the word file
		contentString := createCitationString(bibEntryList)
		// TODO: I don't know enough about Word XML schema, so we are adding the
		// citeproc citations like [@AUTHOR, pages] behind the original in-text
		// citation within the same w:t element
		textContainer := p.FindElement("./sdtContent//r/t")
		if textContainer != nil {
			textContainer.SetText(fmt.Sprintf("%s %s", textContainer.Text(), contentString))
		}
	}
	fmt.Printf("Found %d citations.\n", citations)
	// Writing new XML
	if fileName == "word/document.xml" {
		doc.WriteToFile("NEWDOCUMENT.xml")
	} else {
		doc.WriteToFile("NEWFOOTNOTES.xml")
	}
}
