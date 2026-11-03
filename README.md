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
python pdf2web.py   "Lab Manual.pdf"   -o "Lab Manual.html"
python pdf2web.py   "Lab Manual.pdf"   -o "Lab Manual.html" -c configs/lom_2021.json
```

Options for both scripts:

| Option | Meaning |
|---|---|
| `-o FILE` | output file (default: same name as input, `.html`) |
| `-c FILE` | JSON config (see below) |
| `--title`, `--subtitle` | text shown in the header and on the cover |
| `--no-embed` | don't embed the original. For PDFs, embedding powers the "Original" page viewer; for Word it powers "Download original". Leaving it out makes the HTML smaller. |

On Windows you can also drag a `.pdf` or `.docx` onto `make_web_manual.bat` (it arrives as `make_web_manual.bat.txt` so email allows it: rename it to remove `.txt` first).

The script prints what it found, e.g.

```
structure from toc_links: 10 chapters, 300 sections/exhibits, 70 procedures
204 cross-reference links, 0 glued words fixed
Wrote Lab Manual.html (5.9 MB)
```

To publish a new edition, convert again and replace the file on the drive.

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
- Print gives a clean copy without the sidebar. A dark theme is available.

Network-drive tip: share the full UNC path (`\\server\share\Manuals\Lab Manual.html`) rather than a
mapped drive letter; letters differ between computers.
