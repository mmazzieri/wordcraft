# Synthetic reproducer: a page's last line holding only a picture, at 1.15 spacing

Everything here is invented and was made by `make-picture-threshold.ps1` with Word 16 (COM) on Windows.

Each `pic-m<mode>-h<H>.docx` is A4 (841.9 pt), top margin 70.85 pt, bottom 56.7 pt (a 714.35 pt body), Normal style Calibri 11 pt at 1.15 lines (`w:line="276" w:lineRule="auto"`), no space before or after. It holds one line of text, then a paragraph holding only an inline picture 200 pt wide and H pt tall. `m11` / `m15` is the compatibility mode.

Word keeps the picture on page 1 while the picture itself fits above the bottom margin: the 1.15 spacing's leading and the font's descent below it may run into the margin. WordCraft also asks for the font's descent below the picture, which a line of only pictures does not have below them.

| file | Word 16 | WordCraft `main` | with the fix |
|---|---|---|---|
| pic-m11-h696.0 | 1 | 1 | 1 |
| pic-m11-h697.0 | 1 | 1 | 1 |
| pic-m11-h698.0 | 1 | **2** | 1 |
| pic-m11-h698.5 | 1 | **2** | 1 |
| pic-m11-h699.0 | 2 | 2 | 2 |
| pic-m15-h696.0 | 1 | 1 | 1 |
| pic-m15-h697.0 | 1 | **2** | 1 |
| pic-m15-h698.0 | 1 | **2** | 1 |
| pic-m15-h698.5 | 2 | 2 | 2 |
| pic-m15-h699.0 | 2 | 2 | 2 |

In 0.1 pt steps, Word keeps pictures up to 698.6 pt (mode 11) and 698.2 pt (mode 15). WordCraft `main` keeps them up to 697.1 pt and 696.7 pt. In `pic-m11-h698.0-word.pdf` the picture runs from 86.2 pt to 784.4 pt, and the body ends at 785.2 pt.

`pic-m11-h698.0-{word,main,fix}.pdf`: Word's export, `wordcraft-cli convert` on `main`, and with the fix. `pic-m15-h698.0-word.pdf`: Word's export in mode 15.
