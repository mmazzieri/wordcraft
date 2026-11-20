# Synthetic reproducer: turned text in a table cell wraps within the row in Word

Everything here is invented and was made by `make-turned-cell.ps1` with Word 16 (COM) on Windows.

Each `turned-<rule>-<n>.docx` has a one-row table with two cells. The first cell is 40 pt wide and holds "TURNED LABEL OF SEVERAL INVENTED WORDS" at 9 pt, turned upward (`w:textDirection w:val="btLr"`) and vertically centred. The second cell is 300 pt wide and holds n lines of text. The row height is auto (`auto`) or at least 78 pt (`atleast`).

Table height, from the PDFs (pt):

| file | Word 16 | WordCraft `main` |
|---|---|---|
| turned-auto-3 | 69 | 225 |
| turned-auto-6 | 137 | 225 |
| turned-auto-10 | 227 | 225 |
| turned-atleast-3 | 80 | 225 |
| turned-atleast-6 | 137 | 225 |
| turned-atleast-10 | 227 | 225 |

- **Word:** the row takes the height the other cell (or the minimum) gives it. The turned text wraps into lines as long as that height. What does not fit is cut off: in `turned-auto-3` only "TURNED / LABEL OF / SEVERAL" shows.
- **WordCraft:** the turned text keeps the length it needs unwrapped, and the row grows to it (`crates/layout/src/table.rs`, the `turn.is_turned()` branch).

`*-word.pdf` is Word's export, `*-main.pdf` is `wordcraft-cli convert` on `main` f8ec3cc0, and `side-auto-3.png` puts the two side by side.

**Correction (20/11):** turning text through Word's COM (`Range.Orientation`) also sets `w:cantSplit` on the row and, on a row with no height, an at-least height of 1134 twips (56.7 pt). So the `turned-auto-*` rows are really at least 56.7 pt, and all six rows are `cantSplit`. Word won't clear `cantSplit` on a row with turned text (setting `AllowBreakAcrossPages` back to true is ignored).

## More cases (`more/`, `make-turned-more.ps1`)

Same cells and the same 9 pt label, longer: "TURNED LABEL OF SEVERAL INVENTED WORDS THAT IS LONGER THAN ONE ROW". `rules.py` prints the y of every horizontal rule at least 100 pt long, per page (Ghostscript at 144 dpi): those are the table's borders. Heights are in pt. `base` is `wordcraft-cli` built from `main` 2922fe56. `#5634` is that PR's head, a224da94.

| file | what | Word 16 | base | #5634 |
|---|---|---|---|---|
| `merge-1` | 3 auto rows, 1 line each; column 1 merged over the 3 rows, turned | rows 23.0 / 23.0 / 22.4 | 22.6 / 22.0 / **329.0** | 22.6 / 22.0 / 22.4 |
| `merge-3` | the same, 3 lines each | rows 68.0 / 67.4 / 68.0 | 65.6 / 65.4 / **242.6** | 65.6 / 65.4 / 66.0 |
| `split-atleast-28` | 28 filler paragraphs, then a turned row at least 120 pt (2 lines beside it); about 68 pt left on page 1 | moved whole to page 2, 120.2 | moved, **374.0** | moved, 121.0 |
| `split-atleast-28-nocantsplit` | the same, `cantSplit` removed from the XML | moved whole to page 2, 120.2 | moved, **374.0** | moved, 121.0 |
| `split-atleast-28-upright` | control: the label upright, row splittable | moved whole to page 2, 162.8 | | |
| `split-auto` | a turned row (`cantSplit`, as Word saves it) beside 70 lines: taller than a page | starts on page 2, breaks over pages 2–4; 4 pages | **3 pages: the row is cut at the foot of page 2, lines 37–70 are lost** | the same as base |
| `split-auto-upright-cantsplit` | control: the same with the label upright | 4 pages, as above | 3 pages, cut, as above | |
| `split-auto-nocantsplit` | `split-auto` with `cantSplit` removed from the XML | starts on page 1 and breaks over pages 1–3, like an upright row; 3 pages | **starts on page 2; 6 pages** | the same as base |
| `split-auto-upright` | control: upright, splittable | 3 pages | 3 pages | 3 pages |

- **Vertical merge:** the turned cell doesn't make the merged rows taller in Word. Its text wraps at the merged height (68 pt in `merge-1`, where words break inside, "SEVERA / L"), and the rest is cut off.
- **At-least row at the foot of a page:** Word moves the row whole to the next page when its least height doesn't fit, with or without `cantSplit` and with or without turned text. The turned row stays at its least height.
- **Row taller than a page:** with `cantSplit`, the row is cut. This happens with upright text as well, so it is #4736, not turned text. But Word always saves turned rows as `cantSplit`, so any turned row taller than a page hits it.
- **Turned row without `cantSplit`, taller than a page:** WordCraft moves it to the next page and gives 6 pages, where Word gives 3. Upright text splits correctly. Word itself never writes such a row, so this only comes from other producers or hand edits.

The `*-nocantsplit` and `*-upright-cantsplit` files were edited by hand (only `<w:cantSplit/>` removed or added) and exported by Word. The others come from the script.
