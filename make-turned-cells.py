"""Writes the turned-cell reproducers of storytold/wordcraft#7111 (and the extra cases discussed on
storytold/wordcraft#5634), from our own inputs and without any word processor. Text is invented.

    python make-turned-cells.py [out-dir]

All files: A4, 1" margins, Calibri 11 pt, single spacing, no space after; a bordered table whose
first column is 40 pt wide and holds a label in 9 pt turned upward (w:textDirection btLr, centred
vertically), and whose second column is 300 pt wide.

- turned-<rule>-<n>.docx: one row; the second cell holds n lines (3, 6, 10); the row height is
  auto, or at least 78 pt (rule "atleast"). The label is "TURNED LABEL OF SEVERAL INVENTED WORDS".
- more/merge-<n>.docx: three rows, the first column merged over the three (vMerge) with a longer
  label; the second cell of each row holds n lines (1, 3).
- more/split-auto*.docx: one auto-height row whose second cell holds 70 lines, so the row is taller
  than a page. split-auto: turned, w:cantSplit; -nocantsplit: turned, no cantSplit; -upright: the
  label not turned, no cantSplit; -upright-cantsplit: not turned, w:cantSplit.
- more/split-atleast*.docx: 47 filler paragraphs, then a row at least 120 pt high with two lines in
  the second cell, so the page ends inside the row's least height. split-atleast: turned,
  w:cantSplit; -nocantsplit: turned, no cantSplit; -upright: not turned, no cantSplit.
"""
import sys
import zipfile
from pathlib import Path

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

CONTENT_TYPES = HEAD + """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

RELS = HEAD + f"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = HEAD + f"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="{REL}/styles" Target="styles.xml"/>
</Relationships>"""

STYLES = HEAD + f"""<w:styles {W}><w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"""

LABEL = "TURNED LABEL OF SEVERAL INVENTED WORDS"
LONG_LABEL = "TURNED LABEL OF SEVERAL INVENTED WORDS THAT IS LONGER THAN ONE ROW"
BORDERS = ('<w:tblBorders>' + "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                                      for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
           + '</w:tblBorders>')


def para(text, half_points=None):
    rpr = f'<w:rPr><w:sz w:val="{half_points}"/></w:rPr>' if half_points else ''
    return f'<w:p><w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def label_cell(text, turned=True, vmerge=None):
    merge = {None: '', 'restart': '<w:vMerge w:val="restart"/>', 'continue': '<w:vMerge/>'}[vmerge]
    direction = '<w:textDirection w:val="btLr"/>' if turned else ''
    content = para(text, 18) if vmerge != 'continue' else '<w:p/>'
    return (f'<w:tc><w:tcPr><w:tcW w:w="800" w:type="dxa"/>{merge}{direction}<w:vAlign w:val="center"/></w:tcPr>'
            f'{content}</w:tc>')


def text_cell(lines):
    return ('<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>'
            + "".join(para(line) for line in lines) + '</w:tc>')


def row(cells, height=None, cant_split=False):
    props = ''
    if cant_split:
        props += '<w:cantSplit/>'
    if height:
        props += f'<w:trHeight w:val="{height}" w:hRule="atLeast"/>'
    trpr = f'<w:trPr>{props}</w:trPr>' if props else ''
    return f'<w:tr>{trpr}{cells}</w:tr>'


def table(rows):
    return (f'<w:tbl><w:tblPr><w:tblW w:w="6800" w:type="dxa"/>{BORDERS}<w:tblLayout w:type="fixed"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="800"/><w:gridCol w:w="6000"/></w:tblGrid>'
            f'{rows}</w:tbl>')


def document(body):
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    return HEAD + f'<w:document {W}><w:body>{body}{sect}</w:body></w:document>'


def write(path, body):
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", document(body))


def lines(n, prefix="Line"):
    return [f"{prefix} {i} of the other cell" for i in range(1, n + 1)]


if __name__ == "__main__":
    out = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    out.mkdir(parents=True, exist_ok=True)
    for rule, height in (("auto", None), ("atleast", 1560)):
        for n in (3, 6, 10):
            body = para("Before the table.") + table(row(label_cell(LABEL) + text_cell(lines(n)), height)) + para("After the table.")
            write(out / f"turned-{rule}-{n}.docx", body)
    for n in (1, 3):
        rows = "".join(row(label_cell(LONG_LABEL, vmerge="restart" if r == 1 else "continue")
                           + text_cell([f"Row {r} line {i}" for i in range(1, n + 1)])) for r in (1, 2, 3))
        write(out / "more" / f"merge-{n}.docx", para("Before the table.") + table(rows) + para("After the table."))
    for suffix, turned, cant in (("", True, True), ("-nocantsplit", True, False),
                                 ("-upright", False, False), ("-upright-cantsplit", False, True)):
        body = para("Before the table.") + table(row(label_cell(LONG_LABEL, turned) + text_cell(lines(70)), cant_split=cant)) + para("After the table.")
        write(out / "more" / f"split-auto{suffix}.docx", body)
    fillers = "".join(para(f"Filler paragraph {i} before the table.") for i in range(1, 48))
    for suffix, turned, cant in (("", True, True), ("-nocantsplit", True, False), ("-upright", False, False)):
        body = fillers + table(row(label_cell(LONG_LABEL, turned) + text_cell(lines(2)), 2400, cant)) + para("After the table.")
        write(out / "more" / f"split-atleast{suffix}.docx", body)
