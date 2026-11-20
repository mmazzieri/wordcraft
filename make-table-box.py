"""Writes the reproducer for the case raised on storytold/wordcraft#2496 (a text box with Wrap text in
shape off that holds a table), from our own inputs and without any word processor. Text and
pictures are invented; the three pictures (a circle, a square, a triangle) are drawn here.

    python make-table-box.py [out-dir]

Every file: A4, one paragraph of body text, and a floating text box anchored to it, 57.65 pt from
the page's left edge and 500 pt from its top, no line or fill, insets 5.4 pt left and right. The box
holds a 3 x 3 table with fixed cell widths (tcW 3417, 2786 and 3365 twips) but an automatic
preferred width (tblW 0 auto) and no fixed layout, as a table whose AutoFit was never turned off; a
green border under row 1 and over row 2; "TRAINING BODY" in the middle of row 1; one picture per
cell in row 3.

- none-autofit.docx: wrap="none", a:spAutoFit, saved box width a stale 19.85 pt;
- none-fixed.docx: wrap="none", a:noAutofit, 19.85 pt;
- square-autofit.docx: wrap="square", a:spAutoFit, 19.85 pt;
- wide-none-autofit.docx: wrap="none", a:spAutoFit, saved at the table's width (490 pt).
"""
import struct
import sys
import zipfile
import zlib
from pathlib import Path

HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      f'xmlns:r="{REL}" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"')
EMU = 12700

CONTENT_TYPES = HEAD + """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

RELS = HEAD + f"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = HEAD + f"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="{REL}/styles" Target="styles.xml"/>
<Relationship Id="rId10" Type="{REL}/image" Target="media/pic0.png"/>
<Relationship Id="rId11" Type="{REL}/image" Target="media/pic1.png"/>
<Relationship Id="rId12" Type="{REL}/image" Target="media/pic2.png"/>
</Relationships>"""

STYLES = HEAD + """<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"""


def png(kind, size=60):
    colour = [(30, 60, 160), (110, 60, 110), (190, 60, 60)][kind]
    rows = bytearray()
    for y in range(size):
        rows.append(0)
        for x in range(size):
            if kind == 0:
                inside = (x - 30) ** 2 + (y - 30) ** 2 <= 25 ** 2
            elif kind == 1:
                inside = 8 <= x < 52 and 8 <= y < 52
            else:
                inside = 4 <= y <= 55 and abs(x - 30) <= (y - 4) * 26 / 51
            rows += bytes(colour if inside else (255, 255, 255))

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(bytes(rows))) + chunk(b"IEND", b""))


def inline_picture(rid, n, size_pt=45):
    e = size_pt * EMU
    return ('<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{e}" cy="{e}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="{10 + n}" name="Picture {n}"/><wp:cNvGraphicFramePr/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{10 + n}" name="pic{n}.png"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{e}" cy="{e}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')


GREEN = 'w:val="single" w:sz="12" w:space="0" w:color="008000"'
WIDTHS = (3417, 2786, 3365)


def cell(width, content, border=""):
    borders = f'<w:tcBorders>{border}</w:tcBorders>' if border else ''
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{borders}</w:tcPr>{content}</w:tc>'


def table():
    centre = '<w:pPr><w:jc w:val="center"/></w:pPr>'
    row1 = "".join(cell(w, f'<w:p>{centre}<w:r><w:t>TRAINING BODY</w:t></w:r></w:p>' if i == 1 else '<w:p/>',
                        f'<w:bottom {GREEN}/>') for i, w in enumerate(WIDTHS))
    row2 = "".join(cell(w, '<w:p/>', f'<w:top {GREEN}/>') for w in WIDTHS)
    row3 = "".join(cell(w, f'<w:p>{centre}{inline_picture(f"rId{10 + i}", i)}</w:p>') for i, w in enumerate(WIDTHS))
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in WIDTHS)
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
            + "".join(f'<w:{s} w:val="nil"/>' for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
            + f'</w:tblBorders></w:tblPr><w:tblGrid>{grid}</w:tblGrid>'
            f'<w:tr>{row1}</w:tr><w:tr>{row2}</w:tr><w:tr>{row3}</w:tr></w:tbl><w:p/>')


def text_box(width_pt, wrap, fit):
    cx, cy = int(width_pt * EMU), 80 * EMU
    return ('<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
            'relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="page"><wp:posOffset>{int(57.65 * EMU)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{500 * EMU}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            '<wp:docPr id="1" name="Text box"/><wp:cNvGraphicFramePr/>'
            '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
            '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>'
            f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>'
            f'<wps:txbx><w:txbxContent>{table()}</w:txbxContent></wps:txbx>'
            f'<wps:bodyPr wrap="{wrap}" lIns="{int(5.4 * EMU)}" tIns="0" rIns="{int(5.4 * EMU)}" bIns="0">{fit}</wps:bodyPr>'
            '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')


def document(width_pt, wrap, fit):
    body = (f'<w:p>{text_box(width_pt, wrap, fit)}<w:r><w:t xml:space="preserve">Certificate of attendance '
            '(invented text). The box below holds a three-column table: two green rules and three pictures.</w:t></w:r></w:p>')
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1417" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    return HEAD + f'<w:document {NS}><w:body>{body}{sect}</w:body></w:document>'


def write(path, width_pt, wrap, fit, pictures):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES)
        for i, data in enumerate(pictures):
            z.writestr(f"word/media/pic{i}.png", data)
        z.writestr("word/document.xml", document(width_pt, wrap, fit))


if __name__ == "__main__":
    out = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    out.mkdir(parents=True, exist_ok=True)
    pictures = [png(k) for k in range(3)]
    write(out / "none-autofit.docx", 19.85, "none", "<a:spAutoFit/>", pictures)
    write(out / "none-fixed.docx", 19.85, "none", "<a:noAutofit/>", pictures)
    write(out / "square-autofit.docx", 19.85, "square", "<a:spAutoFit/>", pictures)
    write(out / "wide-none-autofit.docx", 490, "none", "<a:spAutoFit/>", pictures)
