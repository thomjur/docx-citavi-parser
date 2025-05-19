# DOCX-CITAVI-PARSER

This program parses a DOCX file that includes embedded citations using CITAVI's Word plugin. It compares the embedded citations with entries in a BibTeX file to find the matching citation keys/IDs. It then adds the keys in the form of a Citeproc citation used by Pandoc.

For example:

> ... (Krech, Elwert 2024, 2-6) ...

=>

> ... (Krech, Elwert 2024, 2-6) [@krech.2024, 2-6] ...

The program can be used in typesetting pipelines and should facilitate creating a Markdown file with Pandoc, including a first version of CiteProc citations, since Pandoc can, as of today, only handle embedded citations created using the Zotero plugin for Word.

## Usage
You can download the Windows, Mac, or Linux executable or build your own version using `go build` in the main directory of this repository.

The CLI program expects a DOCX file with CITAVI citations and a corresponding BibTeX file.

Simply execute in the command line:

`docx-citavi-parser <PATH-TO-WORD-FILE> <PATH-TO-BIBTEX-FILE>`

The results are a `NEWDOCUMENT.xml` and `NEWFOOTNOTES.xml`, which include the Citeproc citations and should replace the former `document.xml` and `footnotes.xml` in the DOCX archive (in the subfolder word).

Please note that finding the keys in BibTeX currently only works via title comparisons, which misses quite a few entries, particularly those with a lot of non-Latin characters or accentuations.

## Versions

### 0.0.1 (May 10, 2025)

- initial version
