# Rule anchored in the next paragraph's space before

Reproducer for a layout issue: a paragraph after a Top and Bottom object that lies in its space
before (or reaches into its first line).

- `make-rule-in-space-before.py` writes the three `.docx` files from its own inputs, without any
  word processor (`python make-rule-in-space-before.py .`). Text is invented; the files carry no
  document properties.
- `*-main.pdf`: WordCraft `main` (09faec42), `wordcraft-cli convert`.

Each document: Times New Roman 10 pt, single spacing, A4 with 1" margins. Paragraph 2 holds a
144 × 0.5 pt rectangle anchored to it (`wp:positionV relativeFrom="paragraph"`, `wrapTopAndBottom`,
no distance from text).

| file | rule | paragraph 3 |
|---|---|---|
| `rule-in-space-before.docx` | 18 pt below paragraph 2's top: inside paragraph 3's space before | 12 pt before |
| `rule-over-first-line.docx` | 26 pt below paragraph 2's top: over paragraph 3's first line | 12 pt before |
| `rule-in-space-after.docx` | 18 pt below paragraph 2's top: inside paragraph 2's 12 pt space after | no space before |
