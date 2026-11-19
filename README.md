# Text box with Wrap text in shape off holding a table

Synthetic files, all text and pictures invented.

- `make-table-box.ps1` makes `word-saved.docx` with Word for Windows (Microsoft 365) via COM: a floating text box, left 57.65 pt from the page, holding a 3-column table (fixed `tcW` 3417/2786/3365 twips, green row border, three inline pictures). Wrap text in shape off, Resize shape to fit text on.
- `make-variants.ps1` edits the XML by hand: the table's `tblW` becomes `0 auto` and `tblLayout fixed` is dropped (a table whose AutoFit was never turned off), and the box's saved width (`wp:extent`, `a:ext`, VML `width`) becomes a stale 19.85 pt, as a file saved by another Word can carry. Variants: `wrap` none/square, `a:spAutoFit`/`a:noAutofit`.
- `*.word.pdf`: exported by Word from each variant. `none-autofit.wordcraft-0.18.1.pdf`: WordCraft 0.18.1.
