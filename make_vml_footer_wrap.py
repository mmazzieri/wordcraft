"""Synthetic reproducer: a footer whose only paragraph holds a VML picture placed absolutely
(`position:absolute`, centred, `margin-top:0`, no `mso-position-vertical-relative`) with
`<w10:wrap type="square"/>`, in a document that declares no `w:compatibilityMode` (Word opens
it in compatibility mode). Written byte by byte as WordprocessingML (ECMA-376) and VML
(ECMA-376 Part 4) with the Python standard library only: no word processor, no python-docx.
All text is invented; the picture is a flat grey PNG generated here.

usage: python make_vml_footer_wrap.py OUT_DIR
"""
import struct
import sys
import zlib
import zipfile
from pathlib import Path

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = (f'xmlns:w="{W}" xmlns:r="{R}" xmlns:v="urn:schemas-microsoft-com:vml" '
      'xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word"')
HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


def png(w: int, h: int, grey: int) -> bytes:
    raw = b''.join(b'\x00' + bytes([grey]) * w for _ in range(h))
    def chunk(t: bytes, d: bytes) -> bytes:
        return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 0, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b''))


def vml_picture(style: str, wrap: str | None) -> str:
    w10 = f'<w10:wrap type="{wrap}"/>' if wrap else ''
    return (f'<w:r><w:pict><v:shape id="Pic1" type="#_x0000_t75" style="{style}" stroked="f">'
            f'<v:imagedata r:id="rIdImg" o:title=""/>{w10}</v:shape></w:pict></w:r>')


def build(path: Path, *, wrap: str | None, footer_style: str, body: str, compat: str,
          header: str = '', ftr_ppr: str = '') -> None:
    sect = ('<w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/>'
            '<w:footerReference w:type="default" r:id="rIdFtr"/>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1417" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/>'
            '</w:sectPr>')
    document = f'{HEAD}<w:document {NS}><w:body>{body}{sect}</w:body></w:document>'
    hdr = f'{HEAD}<w:hdr {NS}>{header or "<w:p/>"}</w:hdr>'
    picture = vml_picture(footer_style, wrap) if footer_style else ''
    text = '' if picture else '<w:r><w:t>Invented footer line</w:t></w:r>'
    ftr = f'{HEAD}<w:ftr {NS}><w:p>{ftr_ppr}{picture}{text}</w:p></w:ftr>'
    settings = f'{HEAD}<w:settings {NS}>{compat}</w:settings>'
    styles = (f'{HEAD}<w:styles {NS}><w:docDefaults><w:rPrDefault><w:rPr>'
              '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault/></w:docDefaults>'
              '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>')
    ct = (f'{HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Default Extension="png" ContentType="image/png"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
          '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
          '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
          '</Types>')
    rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    root_rels = (f'{HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 f'<Relationship Id="rId1" Type="{rel}/officeDocument" Target="word/document.xml"/></Relationships>')
    doc_rels = (f'{HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rIdSty" Type="{rel}/styles" Target="styles.xml"/>'
                f'<Relationship Id="rIdSet" Type="{rel}/settings" Target="settings.xml"/>'
                f'<Relationship Id="rIdHdr" Type="{rel}/header" Target="header1.xml"/>'
                f'<Relationship Id="rIdFtr" Type="{rel}/footer" Target="footer1.xml"/></Relationships>')
    img_rels = (f'{HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rIdImg" Type="{rel}/image" Target="media/band.png"/></Relationships>')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in [('[Content_Types].xml', ct), ('_rels/.rels', root_rels),
                           ('word/document.xml', document), ('word/styles.xml', styles),
                           ('word/settings.xml', settings), ('word/header1.xml', hdr),
                           ('word/footer1.xml', ftr), ('word/_rels/document.xml.rels', doc_rels),
                           ('word/_rels/header1.xml.rels', img_rels), ('word/_rels/footer1.xml.rels', img_rels)]:
            z.writestr(name, data)
        z.writestr('word/media/band.png', png(600, 50, 0x99))


# Legacy layout options a converter writes into `w:compat`, with no `compatibilityMode`: Word
# then lays the document out in its Word 2007 compatibility mode.
LEGACY_COMPAT = ('<w:compat><w:useWord2002TableStyleRules/><w:growAutofit/><w:useNormalStyleForList/>'
                 '<w:doNotUseIndentAsNumberingTabStop/><w:allowSpaceOfSameStyleInTable/>'
                 '<w:doNotAutofitConstrainedTables/><w:autofitToFirstFixedWidthCell/>'
                 '<w:splitPgBreakAndParaMark/><w:doNotVertAlignCellWithSp/>'
                 '<w:doNotBreakConstrainedForcedTable/><w:doNotVertAlignInTxbx/></w:compat>')


def compat(mode: int | None) -> str:
    if mode is None:
        return LEGACY_COMPAT
    return (f'<w:compat><w:compatSetting w:name="compatibilityMode" '
            f'w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat>')


def para(text: str, *, line: int = 276) -> str:
    t = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ''
    return (f'<w:p><w:pPr><w:spacing w:after="0" w:line="{line}" w:lineRule="auto"/></w:pPr>{t}</w:p>')


SQUARE = 'position:absolute;width:480pt;height:40pt;margin-top:0;margin-left:0;mso-position-horizontal:center;z-index:1'


def cases():
    lines = lambda n, start=0: ''.join(para(f'Body line {i + 1:03d} of the invented text.') for i in range(start, start + n))
    empties = lambda n: ''.join(para('') for _ in range(n))
    # A square-wrapped VML picture in the footer (or the same picture unwrapped), with and
    # without a stated compatibility mode. Page 1 ends with a run of empty paragraphs.
    for wrap in ('square', None):
        for mode in (None, 15):
            name = f'footer-{wrap or "nowrap"}-{"legacy" if mode is None else f"m{mode}"}'
            yield name, dict(wrap=wrap, footer_style=SQUARE, compat=compat(mode),
                             body=lines(36) + empties(8) + lines(30, 36))


if __name__ == '__main__':
    out = Path(sys.argv[1] if len(sys.argv) > 1 else '.')
    out.mkdir(parents=True, exist_ok=True)
    for name, kw in cases():
        build(out / f'{name}.docx', **kw)
        print(name)
