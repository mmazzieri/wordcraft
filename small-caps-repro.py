"""Small Caps spaces and line height: the reproducer of storytold/wordcraft#7062 and #7063, written
from our own inputs and without any word processor or template. Text is invented.

    python small-caps-repro.py [out-dir]

small-caps-repro.docx: Letter, 1" margins, Times New Roman Bold 10 pt, multiple line spacing 1.15.
- "Small Caps:", then one long all-lowercase paragraph in Small Caps;
- "plain (for comparison):", then the same paragraph without Small Caps;
- 40 spaces between two letters, then 40 no-break spaces, in Small Caps and then plain.
"""
import sys
import zipfile
from pathlib import Path

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

CONTENT_TYPES = HEAD + """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

RELS = HEAD + """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = HEAD + """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

STYLES = HEAD + f"""<w:styles {W}><w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"""

TEXT = ("the reviewer should complete this panel only when the issuing office has said in the "
        "notice or in the papers quoted there that the reviewer may limit itself to this panel "
        "without completing any other panel of this part of the form")


def run(text, small):
    rpr = '<w:b/>' + ('<w:smallCaps/>' if small else '')
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'


def para(text, small, after=0):
    return f'<w:p><w:pPr><w:spacing w:after="{after}"/></w:pPr>{run(text, small)}</w:p>'


def document():
    body = ""
    for label, small in (("Small Caps", True), ("plain (for comparison)", False)):
        body += para(label + ":", False)
        body += para(TEXT, small, after=240)
    for small in (True, False):
        body += para("a" + " " * 40 + "a", small)
        body += para("a" + " " * 40 + "a", small)
    sect = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    return HEAD + f'<w:document {W}><w:body>{body}{sect}</w:body></w:document>'


if __name__ == "__main__":
    out = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    out.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out / "small-caps-repro.docx", "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", document())
