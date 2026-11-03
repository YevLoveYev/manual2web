# manual2web

Turns a manual (PDF or Word .docx) into **one self-contained HTML file** with a clickable
table of contents, full-text search, links between sections, and copyable deep links.
The output needs no server, no internet and no installs: put it on a network drive and
double-click it. Access control is whatever the folder's permissions are.

```
manual2web/
  pdf2web.py            PDF  -> HTML   (needs PyMuPDF)
  docx2web.py           Word -> HTML   (Python standard library only)
  webviewer.py          shared code (standard library only)
  viewer_template.html  the page design: sidebar, search, styles
  configs/
    lom_2021.json       settings that reproduce the Laboratory Operations Manual (PDF)
    lom_docx_example.json  the same idea for a Word version of the manual
  make_web_manual.bat   Windows: drag a .pdf or .docx onto it
```

## Requirements

| Input | Needs |
|---|---|
| **.docx** | Python 3.6 or newer. Nothing to install. |
| **.pdf**  | Python 3.8 or newer **and** the PyMuPDF package (one wheel file, no other dependencies). |

**Word is the better source** if you have it: headings, list numbering and tables come straight
from the document's structure instead of being reconstructed from page layout.

### Installing PyMuPDF on a machine without internet

On any computer with internet, download the wheel for the target machine (example: Windows, 64-bit, Python 3.11):

```
pip download pymupdf --only-binary=:all: --platform win_amd64 --python-version 3.11 -d wheels
```

Copy the `wheels` folder over, then on the target machine:

```
pip install --no-index --find-links wheels pymupdf
```

If `pip` isn't allowed, ask IT to install PyMuPDF. Note that PyMuPDF is released under the
AGPL (or a commercial license from Artifex); internal use is normally fine, but your IT or legal
office may want to know.

## Usage

```
python docx2web.py  "Lab Manual.docx"  -o "Lab Manual.html"
python docx2web.py  "Chapter files"    -o "Admin Manual.html" --title "Administrative Manual"
python pdf2web.py   "Lab Manual.pdf"   -o "Lab Manual.html"
python pdf2web.py   "Lab Manual.pdf"   -o "Lab Manual.html" -c configs/lom_2021.json
```

Options for both scripts:

| Option | Meaning |
|---|---|
| `-o FILE` | output file (default: same name as input, `.html`) |
| `-c FILE` | JSON config (see below) |
| `--title`, `--subtitle` | text shown in the header and on the cover |
| `--no-embed` | PDF only: don't embed the original PDF (smaller file, but no "Original" page button on headings). Word files are never embedded. |

On Windows you can also drag a `.pdf`, a `.docx`, or a folder of chapter `.docx` files onto `make_web_manual.bat`.

The script prints what it found, e.g.

```
structure from toc_links: 10 chapters, 300 sections/exhibits, 70 procedures
204 cross-reference links, 0 glued words fixed
Wrote Lab Manual.html (5.9 MB)
```

To publish a new edition, convert again and replace the file on the drive.

## A folder of chapter files -> one page

If each chapter is its own Word file, give `docx2web.py` the folder (or a list of files):

```
python docx2web.py "C:\Manuals\AM chapters" -o "Admin Manual.html" --title "Administrative Manual"
```

- Every `.docx` in the folder becomes one chapter. The chapter number is the first number in the
  file name (`AM6111.docx` -> Chapter 6111) and chapters are put in number order (612 before 6111).
  Change the rule with `"chapter_from_filename"` in a config if your names differ.
- The chapter title comes from the file's Title-style paragraph, or its single top heading, or the
  document's Title property (File > Info), or failing that the file name. "Chapter 6111" / "AM6111"
  prefixes are removed so the title isn't doubled.
- **Header block.** If each file starts with a block of lines ending in the chapter number, e.g.

  ```
  Drug Enforcement Administration      <- organization (shown on the cover)
  Agents Manual                        <- manual name (page title)
  Reporting System                     <- group     (sidebar heading)
  General Policies                     <- sub-group (sidebar heading)
  Reporting Requirements               <- chapter title
  6211                                 <- chapter number
  ```

  those lines are used as shown and not repeated in the text. The block can be plain lines, a table,
  or Word's page header; the number line can be `6211`, `AM 6211`, `Chapter 6211` or decorated like `**6126**`. If a file has
  fewer lines (no sub-group, say), the line right above the number is still the title. Files without
  such a block are listed in the script's output and fall back to the rules below.
  Turn this off with `"header_block": false`.
- **Appendix files.** If the block has an `Appendix D` line instead of a number, the title is taken from the
  same line (`Appendix D – Abbreviations`) or the next line with text. Appendices are placed after all
  chapters, in letter order, under an **Appendices** heading, and "Appendix D" in the text becomes a link.
- All headings inside a file become that chapter's sections. Word cross-references, bookmarks and
  footnotes stay inside their own chapter, even if two files use the same internal names.
- Word's temporary lock files (`~$...docx`) are ignored. Old `.doc` files must be saved as `.docx` first.
- The script prints the file -> chapter mapping, so you can check the order and titles at a glance.

To link chapters to each other, add an autolink rule for the way your manual cites them, e.g.
`{ "regex": "\\bAM\\s+(\\d{3,4}(?:\\.\\d+)?)", "kind": "section" }` turns "AM 6112.1" into a link.

## Hosting: network drive vs SharePoint

The output is one self-contained file, so it works anywhere it is **opened from**: a network drive,
a OneDrive/SharePoint folder synced to the computer, or a local copy.

**SharePoint Online does not display .html files in the browser. It downloads them** (renaming to
`.aspx` doesn't help). Options that work:

1. **Network drive.** Double-click opens it in the browser. Simplest.
2. **SharePoint as storage + sync.** Put the file in a document library; users click **Sync** (or
   *Add shortcut to My files*) once, then open it from File Explorer like a network-drive file. It
   updates automatically when you replace the file.
3. **Download and open.** Clicking the file in SharePoint downloads it; opening the download works,
   but people end up with old copies after updates.
4. **An internal web server** (IIS or similar) run by IT. This gives a real web address and works
   with any browser.

On-premises SharePoint Server can be set by IT to allow HTML to display ("permissive browser file
handling"), but that is an administrator decision with security trade-offs.

## How the structure is found

**Word (`docx2web.py`)**
- Headings come from Word's *Heading 1-9* styles (or any style with an outline level).
- The top heading level becomes the chapters; the next three levels become sections in the sidebar.
  If there is only one Heading 1 (a document title), Heading 2 becomes the chapter level.
- Word's numbering is computed exactly (A., 1., a., i., bullets, "7302.12"-style heading numbers).
- Word's own Table of Contents, headers, footers and deleted tracked changes are left out.
- **No Heading styles?** Bold lines that look like headings (`CHAPTER 70 ...`, `7001 ...`,
  `7001.1 ...`, `Exhibit ...`) are promoted automatically. Adjust with `heading_patterns`.
- Pictures in EMF/WMF format (old clip-art, some Visio pastes) can't be shown by browsers; they're
  replaced by a note. Re-insert them as PNG in Word if they matter.

**PDF (`pdf2web.py`)**, in order of preference:
1. Links on the PDF's own "Table of Contents" pages (PDFs exported from Word have these),
2. the PDF's bookmarks,
3. bold numbered headings in the text.

Page headers and footers that repeat on many pages are removed automatically. Text is checked for
words that the PDF glued together and these are repaired. A PDF is a page layout, not a document
structure, so a few complex layouts (two-column "Action By / Action" procedure pages, for example)
come out simplified. The **Original** button on every heading shows the exact PDF page.

## Config file

All keys are optional. The PDF and Word scripts share most of them.

```json
{
  "title": "Laboratory Operations Manual",
  "subtitle": "Office of Forensic Sciences · June 2021",
  "org": "Drug Enforcement Administration · Office of Forensic Sciences",
  "meta": "June 2021 edition",

  "chapter_regex": "^CHAPTER\\s+(\\d+)\\s+(.*)$",
  "section_regex": "^(\\d{4}(?:\\.\\d+)?)\\s*(.*)$",
  "exhibit_regex": "^(?:Exhibit\\s*)?(\\d+/\\d+)\\s*(.*)$",

  "autolink": [
    { "regex": "\\b((?:PRO|TSK)-\\d{4}(?:\\.\\d+)?[A-Z]?\\d?)", "kind": "code" },
    { "regex": "\\bLOM\\s+(\\d{4}(?:\\.\\d+)?[A-Z]?)",          "kind": "section" },
    { "regex": "\\bExhibit\\s+(\\d+/\\d+)",                      "kind": "exhibit" }
  ]
}
```

- `toc_case`: capitalization of sidebar entries and chapter/section headings: `"title"` (default,
  Title Case; acronyms and codes like LIMS, SFL2, AM6311 are kept), `"upper"` (ALL CAPS) or `"asis"`.
- `keep_upper`: extra words that must stay in capitals, e.g. `["NARCAN", "OCDETF"]` (2-4 letter acronyms
  like PPE are kept automatically). `keep_case`: words with capitals inside them that are correct,
  e.g. `["SharePoint", "eGuardian"]`. Anything else with odd casing (naLOXone, privilEges, a shouted
  ASSISTANCE) is normalized in titles.
- `autolink` turns references in the text into links. The first `( )` group is the key looked up:
  `code` = a procedure code, `section` = a section number (closest parent section if the exact one
  doesn't exist, so `LOM 7101.2A` goes to 7101.2), `exhibit` = an exhibit number.
- PDF-only keys: `skip_pages`, `skip_toc_pages`, `page_label_offset`, `chapter_bookmark_level`,
  `sections_from` (`toc_links` / `bookmarks` / `headings`), `level_rule` (`dots` or `lom`),
  `section_titles` (fix a title the PDF splits badly), `procedures` (procedure documents whose
  titles are only in the contents pages), `drop_line_regex`, `footer_auto`.
- Word-only keys: `chapter_level`, `procedure_regex`, `fallback_headings`, `heading_patterns`.

See `configs/lom_2021.json` for a complete working example.

## Using the page

- `/` or `Ctrl+K` search; arrow keys and Enter pick a result; `F3` / `Shift+F3` step through every match.
- Quotes force an exact phrase: `"chain of custody"`. Section numbers and codes work too: `7302.12`, `PRO-7305.22G`.
- Hover a heading: **Link** copies a deep link to that section (`...Manual.html#s7302-12`);
  **Original** (PDF only) opens that page of the original PDF.
- `Ctrl+P` prints a clean copy without the sidebar. A dark theme is available (moon button).

Network-drive tip: share the full UNC path (`\\server\share\Manuals\Lab Manual.html`) rather than a
mapped drive letter; letters differ between computers.
