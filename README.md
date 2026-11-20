# Synthetic reproducer: exact (negative) top and bottom page margins

Everything here is invented and was made by `make-negative-margins.ps1` with Word 16 (COM) on Windows.

`negative-margins.docx` / `.doc`: A4, top and bottom margins **-60 pt** (`w:pgMar w:top="-1200" w:bottom="-1200"`; in the `.doc`, negative `sprmSDyaTop` / `sprmSDyaBottom`), header and footer 36 pt from the edge with six 14 pt lines each, so both bands are taller than the margins. The body is 100 lines of 11 pt text.

`positive-margins.docx` / `.doc`: the same with +60 pt margins (the control).

ECMA-376 Part 1 §17.6.11 (`pgMar`, `top` and `bottom` are `ST_SignedTwipsMeasure`): a negative value measures the text from the page edge whatever the header (footer), which then overlaps it.

| file | Word 16 | WordCraft `main` | with the fix |
|---|---|---|---|
| negative-margins.docx | 2 pages, body from 60 pt to 60 pt above the bottom, over the header and footer | 3 pages, body below the header and above the footer | 2 pages, as Word |
| negative-margins.doc | 2 pages (reopened in Word) | 3 pages | 2 pages |
| positive-margins.docx / .doc | 3 pages | 3 pages | 3 pages |

- `*-word.pdf`: Word's export.
- `negative-margins-{docx,doc}-main.pdf`: `wordcraft-cli convert` on `main`.
- `negative-margins-{docx,doc}-fix.pdf`: with the fix (craft-libs readers keep the sign, wordcraft layout honours it).
- `side-by-side-page1.png`: page 1, Word / main / fix.

Also measured in Word on `negative-margins.docx` (Page Setup through COM): setting the left margin to 80 keeps top and bottom at -60; turning the page to landscape gives left and right 60 and 60 (sizes, no longer exact).
