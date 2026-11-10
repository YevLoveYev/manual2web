# manual2web

Turns a manual (PDF or Word .docx) into **one self-contained HTML file** with a clickable
table of contents, full-text search, links between sections, and copyable deep links.
The output needs no server, no internet and no installs: put it on a network drive and
double-click it. Access control is whatever the folder's permissions are.

```
manual2web/
  pdf2web.py            PDF  -> HTML   (needs PyMuPDF)
  docx2web.py           Word -> HTML   (Python standard library only)
  links2excel.py        Word -> Excel list of every hyperlink (standard library only)
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

  those lines are used as shown and not repeated in the text. The lines are recognized by what they are,
  not only by position: the first line is the organization; any line that is the manual's name ("Agents
  Manual", "Agent Manual", "Agent's Manual", or whatever `title` says in the config) is the overall title
  wherever it appears; blank lines are skipped; the remaining lines are chapter, sub-chapter and title.
  Other spellings of the manual's name can be listed in `"manual_names": [...]`. Lines naming the
  organization (Drug Enforcement Administration/Agency, or `org` from the config) and banner lines such as
  "DEA SENSITIVE", "This manual is the property of ..." or "U.S. Department of Justice" are never used as a
  chapter or title; add more patterns with `"header_ignore": [...]`. "CHAPTER 61 GENERAL" and
  "Subchapter 611 Authority and Conduct" are shown as "General" and "Authority and Conduct", and a number
  with its title on one line ("6111 AUTHORITY") works too. A title that wraps onto a second line (or a second
  paragraph) is joined back together; text inside a logo or text box in the header is ignored.
  `"header_levels": 2` = how many header lines are chapter levels in the sidebar (chapter, sub-chapter). The block can be plain lines, a table,
  or Word's page header; the number line can be `6211`, `AM 6211`, `Chapter 6211` or decorated like `**6126**`. If a file has
  fewer lines (no sub-group, say), the line right above the number is still the title. Files without
  such a block are listed in the script's output and fall back to the rules below.
  Turn this off with `"header_block": false`.
- **Appendix files.** If the block has an `Appendix D` line instead of a number, the title is taken from the
  same line (`Appendix D – Abbreviations`) or the next line with text. Appendices are placed after all
  chapters, in letter order, under an **Appendices** heading, and "Appendix D" in the text becomes a link.
- **Sections without Word heading styles.** In appendix files (and any file without Heading styles) an
  ALL-CAPS line on its own becomes a section, e.g. `AMPHETAMINES:` or `OTHER MEANS OF IDENTIFICATION`
  (trailing colon removed). Labels such as `TEST MATERIALS:`, `TEST PROCEDURE:`, `COLOR REACTIONS:`,
  `CAUTION:` and `NOTE:` stay as text, as do ALL-CAPS table cells and list items. Change the label list with
  `"caps_heading_ignore"`, or switch this off with `"caps_headings": false` (`true` = in every file).
  In appendix files those labels stay text even when Word styles them as headings (`"appendix_labels"`),
  and all sections are listed at one level, with no nesting (`"appendix_flat": true`).
- All headings inside a file become that chapter's sections. Word cross-references, bookmarks and
  footnotes stay inside their own chapter, even if two files use the same internal names.
- Word's temporary lock files (`~$...docx`) are ignored. Old `.doc` files must be saved as `.docx` first.
- The script prints the file -> chapter mapping, so you can check the order and titles at a glance.

To link chapters to each other, add an autolink rule for the way your manual cites them, e.g.
`{ "regex": "\\bAM\\s+(\\d{3,4}(?:\\.\\d+)?)", "kind": "section" }` turns "AM 6112.1" into a link.

## Hyperlink report (links2excel.py)

Lists every hyperlink in a set of Word files in an Excel workbook, to see which are good, broken or risky.
Standard library only, like docx2web.py.

```
python links2excel.py "AM chapters"                      -> "AM chapters - links.xlsx"
python links2excel.py "AM chapters" -o links.xlsx -c configs/agents_manual.json
python links2excel.py "AM chapters" --check              also try to open web links / network paths
```

It reads body text, tables, text boxes, page headers/footers, footnotes and endnotes, and finds normal
hyperlinks, links to places inside the document, HYPERLINK fields, Word cross-references and links on
pictures. Word's own table of contents is skipped unless you add `--include-toc`.

Sheets:
- **Links**: one row per link, with file, chapter number and title (from the header block), the section
  heading it sits under, where it is (body / page header / footnote), link text, type, status, target
  (clickable), SharePoint site, how many times that target is used, and notes. Filter and sort as needed.
- **Unique targets**: each address once, with how often and in which files it's used.
- **Summary**: totals by status, type, SharePoint site, and the files with the most broken links.

Status: **OK** (target found), **BROKEN** (bookmark missing, Word's "Error! Reference source not found",
file missing, or a 404), **CHECK MANUALLY** (needs your login or the share isn't reachable),
**NOT CHECKED** (web link; use `--check`), **EMAIL**.
Notes flag SharePoint sharing links (can expire), personal OneDrive links, library views instead of
documents, links to files that are themselves part of the set (candidates for internal links), http
instead of https, drive letters, and link text that shows a different address than the real target.

`--check` can't sign in to SharePoint or other sites that need your Windows login, so those come back as
CHECK MANUALLY; public web links and network paths are tested properly.

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

- `footer`: text for a thin bar fixed to the bottom of the page, e.g.
  `"footer": "Official Agents Manual - <your wording>"`. Empty or missing = no bar. Can also be given as
  `--footer "..."` on the command line. When printing, the text appears once at the end.
- `link_base`: the address where the page is published, e.g.
  `"https://agency.sharepoint.us/sites/X/SitePages/Manual.html"`. The **Link** next to each heading then
  points to that address + the section (`...Manual.html#6111-22`), so copied links work for everyone.
  Clicking it still just jumps to the section. Without `link_base` the link is the section only (`#6111-22`).
- `section_id_prefix` / `chapter_id_prefix`: how section and chapter addresses start. Default `"s"` / `"ch"`
  (`#s6111-22`, `#ch6111`); set both to `""` for plain numbers (`#6111-22`, `#6111`).
- `footer_size`: footer text size in pixels (default 12.5), e.g. `"footer_size": 14`. Long footer text wraps
  onto more lines on narrow windows instead of being cut off.
- `popup`: a notice that pops up when the page opens, with an OK button. `true` = shows the `footer` text,
  or give its own text: `"popup": "..."`. `popup_frequency`: `"always"` (every time the page opens, default)
  or `"session"` (once until the browser is closed). `popup_title` / `popup_button` change "Notice" / "OK".
- `cover`: `false` removes the title block at the top of the page, so it starts with the first chapter.
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
- Hover a heading: **Copy link** copies the address of that section (with `link_base`, the full SharePoint
  address). If the browser or SharePoint blocks copying, a box shows the address selected; press Ctrl+C.
  **Original** (PDF only) opens that page of the original PDF.
- `Ctrl+P` prints a clean copy without the sidebar. A dark theme is available (moon button).

Network-drive tip: share the full UNC path (`\\server\share\Manuals\Lab Manual.html`) rather than a
mapped drive letter; letters differ between computers.
