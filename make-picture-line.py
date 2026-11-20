"""Writes the reproducer of storytold/wordcraft#7109, from our own inputs and without any word
processor: one line of text, then a paragraph holding only an inline picture of height H, at
multiple line spacing 1.15. The picture is a plain grey PNG that this script draws. Text is invented.

    python make-picture-line.py [out-dir]

Every file: A4 (595.3 x 841.9 pt), top 70.85 pt, bottom, left and right 56.7 pt, so the body is
714.35 pt tall; Normal is Calibri 11 pt, no space before or after, line="276" lineRule="auto".
The picture is 200 pt wide. One file per H (696, 697, 698, 698.5, 699 pt) and compatibility mode
(11, 15): pic-m<mode>-h<H>.docx.
"""
import struct
import sys
import zipfile
import zlib
from pathlib import Path

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
NS = (W + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
      ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
      ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
      ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 12700  # per point

CONTENT_TYPES = HEAD + """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

RELS = HEAD + f"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = HEAD + f"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="{REL}/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="{REL}/settings" Target="settings.xml"/>
<Relationship Id="rId3" Type="{REL}/image" Target="media/grey.png"/>
</Relationships>"""

STYLES = HEAD + f"""<w:styles {W}><w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"""


def settings(mode):
    return HEAD + (f'<w:settings {W}><w:compat><w:compatSetting w:name="compatibilityMode" '
                   f'w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>')


def grey_png(size=200, level=0xB0):
    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
    rows = b"".join(b"\x00" + bytes([level]) * size for _ in range(size))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 0, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b""))


def picture(height_pt, width_pt=200):
    cx, cy = int(width_pt * EMU), int(round(height_pt * EMU))
    return (
        '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:docPr id="1" name="Grey picture"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="grey.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
        '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
    )


def document(height_pt):
    body = ('<w:p><w:r><w:t>A line of invented text.</w:t></w:r></w:p>'
            f'<w:p>{picture(height_pt)}</w:p>')
    sect = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1417" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')
    return HEAD + f'<w:document {NS}><w:body>{body}{sect}</w:body></w:document>'


def write(path, height_pt, mode, png):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/settings.xml", settings(mode))
        z.writestr("word/media/grey.png", png)
        z.writestr("word/document.xml", document(height_pt))


if __name__ == "__main__":
    out = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    out.mkdir(parents=True, exist_ok=True)
    png = grey_png()
    for mode in (11, 15):
        for h in (696, 697, 698, 698.5, 699):
            write(out / f"pic-m{mode}-h{h:.1f}.docx", h, mode, png)
