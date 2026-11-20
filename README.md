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

Not measured: a row whose cells are all turned, a vertically merged turned cell, and a row split across pages.
