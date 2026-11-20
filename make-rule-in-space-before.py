"""Writes small DOCX files, from our own inputs and without any word processor: a rule anchored to a
paragraph with Top and Bottom wrapping, placed Y pt below that paragraph's top, then a paragraph
with space before. Text is invented.

    python make-rule-in-space-before.py [out-dir]

Files (Times New Roman 10 pt, single spacing, A4, 1" margins; the rule is 144 x 0.5 pt):
- rule-in-space-before.docx: Y = 18, the rule lies inside the next paragraph's 12 pt space before;
- rule-over-first-line.docx: Y = 26, the rule reaches into the next paragraph's first line;
- rule-in-space-after.docx: Y = 18, the paragraph holding the rule has 12 pt space after and the
  next paragraph none: the rule lies in the space after.
"""
import sys
import zipfile
from pathlib import Path

EMU = 12700  # per point

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


def para(text, before=0, after=0, inner=""):
    return (f'<w:p><w:pPr><w:spacing w:before="{before}" w:after="{after}"/></w:pPr>'
            f'{inner}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>')


def rule(y_pt):
    cx, cy = 144 * EMU, int(0.5 * EMU)
    return (
        '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" '
        'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="page"><wp:posOffset>{72 * EMU}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{int(y_pt * EMU)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:wrapTopAndBottom/><wp:docPr id="1" name="Rule"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr/><wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
        '<a:ln><a:noFill/></a:ln></wps:spPr><wps:bodyPr/></wps:wsp>'
        '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
    )


def document(y_pt, before, after):
    body = (
        para("First paragraph of invented text.")
        + para("The rule is anchored to this paragraph.", after=after, inner=rule(y_pt))
        + para("Second paragraph, after the rule.", before=before)
        + para("Third paragraph of invented text.")
    )
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document {NS}><w:body>{body}{sect}</w:body></w:document>'


def write(path, y_pt, before, after):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", document(y_pt, before, after))


if __name__ == "__main__":
    out = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    out.mkdir(parents=True, exist_ok=True)
    # Twips: 240 = 12 pt.
    write(out / "rule-in-space-before.docx", 18, before=240, after=0)
    write(out / "rule-over-first-line.docx", 26, before=240, after=0)
    write(out / "rule-in-space-after.docx", 18, before=0, after=240)
