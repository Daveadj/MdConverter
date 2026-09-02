# Markdown Converter

Converts documents to Markdown and Markdown to DOCX, PDF and HTML. It runs
entirely in the browser — there is no server, no install, and no upload. Files
are read and written locally and never leave the machine.

## Running it

Double-click `start.bat`. It serves this folder on <http://localhost:8770/> and
opens a browser.

Opening `index.html` directly from disk does *not* work: pdf.js loads its worker
from a CDN, and browsers block that over `file://`. Any static server will do —
`npx serve -l 8770` is an alternative if Python is not installed.

An internet connection is needed on first load, to fetch the libraries listed
below. Everything after that is local.

## What it converts

**To Markdown** — PDF, DOCX, XLSX/XLS, PPTX, CSV/TSV, HTML, EPUB, JSON, XML,
TXT.

**From Markdown** — DOCX, PDF (via the browser's print engine), standalone HTML.

Drop files on the left, pick one in the queue to preview it, then save
individually or download the batch as a zip.

## How well it works

Round-tripping Markdown through DOCX and back preserves headings, tables, bold
and italic, links, ordered and unordered lists including nesting, blockquotes,
code blocks, and exact numbers. The one thing it does not preserve is the
language tag on a fenced code block, because Word has nowhere to store it.

Two conversions get real help from heuristics, and both can be switched off in
**Options** if they misjudge a particular file:

- **Inferred headings.** Plenty of documents — including CCF's own payment
  framework — never apply heading styles and simply bold their section titles.
  For DOCX, a short, fully-bold, non-sentence paragraph becomes a heading. For
  PDF, where font *size* is often identical throughout, the signal is that
  heading text is drawn with a different embedded font resource than the body.
  Depth comes from numbering: `1.` is an `##`, `2.1` is a `###`.
- **PDF tables.** A run of lines that split into the same number of columns
  becomes a pipe table. A header row that spans its columns too widely to split
  on its own is re-cut against the column positions the data rows establish, so
  it is not lost and does not cost you the first data row.

## Known limits

These are properties of the formats, not bugs waiting to be fixed:

- **Conversion is lossy in one direction by design.** Markdown has no concept of
  page geometry, fonts, or positioning. `PDF → MD → PDF` returns the *content*,
  not the *document*. This is a content tool, not a format-preserving one.
- **Scanned PDFs produce nothing.** There is no OCR. The file will convert with
  a warning saying so.
- **Lists in a PDF need a bullet character in the text layer.** Word-exported
  PDFs have one and convert correctly. PDFs printed from a browser draw list
  markers outside the text layer, so their lists come back as prose.
- **Spreadsheets laid out as formatted reports convert poorly.** A pipe table
  can only express a grid of values; merged cells, spacer columns and
  hand-placed subtotals survive as data but not as a readable table. Sheets wide
  enough for this to bite are flagged with a warning.
- **Legacy `.doc`** is not supported. Re-save as `.docx`.

## Layout

| File | Contains |
| --- | --- |
| `index.html` | Markup and the CDN script tags |
| `assets/app.css` | All styling, light and dark |
| `src/to-md.js` | Every *into* Markdown converter, behind one `convert()` |
| `src/from-md.js` | DOCX, PDF and HTML generation |
| `src/app.js` | UI: the queue, previews and downloads |

`src/app.js` knows nothing about file formats and the two converter modules know
nothing about the DOM, so either half can be changed without disturbing the
other.

## Swapping in a server-side engine

The weakest conversion here is PDF → Markdown, because reconstructing structure
from positioned text runs is guesswork that Microsoft's **MarkItDown** and
**Pandoc** have years of accumulated heuristics for. If that quality matters
more than having no install, `ToMd.convert()` in `src/to-md.js` is a single
dispatch point — replacing the `case 'pdf'` branch with a `fetch()` to a local
MarkItDown service is a one-line change, and nothing else, the UI included,
needs to know.

## Libraries

Loaded from CDN, all permissively licensed: marked, Turndown (+ GFM plugin),
mammoth.js, SheetJS, pdf.js, JSZip, DOMPurify, docx.js.

If a CDN fails, the app says which library is missing rather than silently
dropping the feature that depended on it.
