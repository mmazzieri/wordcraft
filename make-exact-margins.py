"""Writes the reproducer of storytold/wordcraft#7107 and storytold/craft-libs#2516, from our own
inputs and without any word processor: a header and a footer taller than the top and bottom
margins, with those margins negative ("exact", ECMA-376 Part 1 17.6.11) in one document and
positive in the other. Text is invented.

    python make-exact-margins.py [out-dir]

Both files: A4, left and right 72 pt, header and footer 36 pt from the edge, six-line header and
footer in 14 pt, then 100 body lines in 11 pt, single spacing, no space after.
- negative-margins.docx: w:top and w:bottom -1200 (exact 60 pt);
- positive-margins.docx: w:top and w:bottom 1200 (at least 60 pt), the control.

There is no .doc here: writing a .doc needs a compound-file writer. The .doc side of the reader
change is covered by the test in storytold/craft-libs#2518, which builds the sprms from the spec.
"""
import sys
import zipfile
from pathlib import Path

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
WML = "application/vnd.openxmlformats-officedocument.wordprocessingml"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

CONTENT_TYPES = HEAD + f"""<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="{WML}.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="{WML}.styles+xml"/>
<Override PartName="/word/header1.xml" ContentType="{WML}.header+xml"/>
<Override PartName="/word/footer1.xml" ContentType="{WML}.footer+xml"/>
</Types>"""

RELS = HEAD + f"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = HEAD + f"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="{REL}/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="{REL}/header" Target="header1.xml"/>
<Relationship Id="rId3" Type="{REL}/footer" Target="footer1.xml"/>
</Relationships>"""

STYLES = HEAD + f"""<w:styles {W}><w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"""


def para(text, half_points=None):
    rpr = f'<w:rPr><w:sz w:val="{half_points}"/></w:rPr>' if half_points else ''
    return f'<w:p><w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def band(tag, label):
    lines = "".join(para(f"{label} LINE {i}", 28) for i in range(1, 7))
    return HEAD + f'<w:{tag} {W} {R}>{lines}</w:{tag}>'


def document(top_bottom):
    body = "".join(para(f"Body line {i} of the invented text.") for i in range(1, 101))
    sect = ('<w:sectPr><w:headerReference w:type="default" r:id="rId2"/><w:footerReference w:type="default" r:id="rId3"/>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            f'<w:pgMar w:top="{top_bottom}" w:right="1440" w:bottom="{top_bottom}" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
            '</w:sectPr>')
    return HEAD + f'<w:document {W} {R}><w:body>{body}{sect}</w:body></w:document>'


def write(path, top_bottom):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/header1.xml", band("hdr", "HEADER"))
        z.writestr("word/footer1.xml", band("ftr", "FOOTER"))
        z.writestr("word/document.xml", document(top_bottom))


if __name__ == "__main__":
    out = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    out.mkdir(parents=True, exist_ok=True)
    write(out / "negative-margins.docx", -1200)
    write(out / "positive-margins.docx", 1200)
