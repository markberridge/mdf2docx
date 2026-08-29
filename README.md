# mdf2docx

Converts MDF module development form PDFs into editable Word documents that
keep the original layout: the pale blue label and header cells, the bordered
boxes, column proportions, bold labels and heading hierarchy.

It reads each PDF's own geometry (cell borders, fill colours, font sizes)
rather than matching against a fixed list of field names, so it adapts to
whatever a given module form contains: any number of learning outcomes,
assessment components, availability rows or requisites.

## Install

Python 3.9 or newer.

```
pip install -r requirements.txt
```

That is `pdfplumber` and `python-docx`. Nothing else, no Word or LibreOffice
needed.

## Use

One file:

```
python3 mdf2docx.py SPR112_Comparative_Politics.pdf
```

writes `SPR112_Comparative_Politics.docx` next to the PDF.

A whole folder:

```
python3 mdf2docx.py --batch pdfs/ --out docx/
```

Every PDF in `pdfs/` is converted into `docx/`, four at a time. Roughly two
seconds per document, so a hundred forms take about a minute. Add
`--recursive` to include subfolders; the folder structure is mirrored in the
output.

## Options

| Option | Effect |
|---|---|
| `--body-size PT` | body text size, default 10. Headings scale with it, keeping their relative proportions |
| `--font NAME` | output font, default Arial |
| `--page-breaks` | keep the PDF's original page breaks. Off by default, so sections flow and every heading is spaced identically |
| `--jobs N` | parallel workers in batch mode, default 4 |
| `--recursive` | search subfolders |
| `--no-check` | skip the completeness check |

## What it does to the layout

- **A4 landscape**, same margins as the source.
- **Text size.** The source PDFs are shrink-to-fit prints, so their body text is
  about 6.6pt. Output is 10pt by default, with sub-headings, section headings
  and the title scaled by the same factor to preserve the original proportions.
- **Column widths** start from the proportions measured in the PDF. Because the
  output text is larger, some columns no longer fit; those are widened just
  enough that no heading or value breaks mid-word, and the space is taken from
  columns that have room to spare. Everything else is left as it was.
- **Shared grids.** All the stacked label/value boxes use one set of widths, so
  the labels line up down the page even though each box is its own table.
- **Page breaks** are dropped by default and headings are set to keep with the
  content below them, so no heading is stranded at the foot of a page. Rows
  will not split across pages, and a long table repeats its header row.
- **Split tables.** A data table the PDF broke across pages is rejoined into
  one table. Any repeated header row is discarded. The stacked label boxes are
  deliberately not merged with each other.
- **Font.** The source uses DejaVu Sans, which most Windows and Mac machines do
  not have, so output uses Arial to stop Word substituting something
  unpredictable. Use `--font` to change it.

## Checking the output

After writing each document the converter compares every word in the source
PDF against the words in the Word file, and reports anything that did not carry
over. On success you get:

```
All 738 words from the PDF are present.
```

In batch mode any file that falls short is listed again at the end under
"Check these", and a file that fails outright is reported without stopping the
run. The exit code is 0 if everything converted, 2 if anything failed.

This matters on a large batch: a form with an unusual structure would otherwise
produce a quietly incomplete document that looks fine at a glance.

## If something looks wrong

- **A file fails to convert.** The error is printed next to the filename. A
  "No /Root object" message means the file is not a readable PDF.
- **Coverage warning.** Some text did not reach the output, usually because a
  table in that form is drawn differently. Send me that PDF and I will adjust
  the parser.
- **A table is too cramped.** Try a smaller `--body-size`. At 10pt the widest
  table in the standard form (module availability, 13 columns) fits, but a form
  with longer values in those columns may need 9.
- **Text is the wrong font in Word.** Pass `--font` with a font installed on the
  machines that will open the documents.

## Files

- `mdf2docx.py` - the converter
- `requirements.txt` - dependencies
- `make_fixture.py` - builds a synthetic form with a page-spanning table, a
  merged full-width row and stacked boxes, for testing changes to the parser
