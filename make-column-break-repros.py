"""Writes the column-break reproducers of storytold/wordcraft#7053, #7054 and #7057, from our own
inputs and without any word processor. Text is invented.

    python make-column-break-repros.py [out-dir]

Files (Arial 8 pt, A4, 1" top and bottom, 2 cm sides):
- column-break-repro.docx (compatibility mode 15): four times, a one-column heading, then two
  columns with three lines on the left and a paragraph in 18 pt with 24 pt space before that opens
  with a column break, then a continuous break; then 48 one-column filler lines.
- column-break-line-mode14-repro.docx (compatibility mode 14): the same with no space before,
  then 57 filler lines.
- column-break-list-number-repro.docx (compatibility mode 15): a numbered list of four items in two
  columns, the fourth opening with a column break, then a continuous break.
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
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

RELS = HEAD + """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = HEAD + """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""

STYLES = HEAD + f"""<w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="16"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"""

NUMBERING = HEAD + f"""<w:numbering {W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"""


def settings(mode):
    return HEAD + (f'<w:settings {W}><w:compat><w:compatSetting w:name="compatibilityMode" '
                   f'w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>')


def sect(columns):
    cols = f'<w:cols w:num="{columns}" w:space="720"/>' if columns > 1 else '<w:cols w:space="720"/>'
    return ('<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1134" w:bottom="1440" w:left="1134" w:header="0" w:footer="0" w:gutter="0"/>'
            f'{cols}</w:sectPr>')


def para(text, ppr=""):
    return f'<w:p><w:pPr>{ppr}</w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def section_end(columns):
    return f'<w:p><w:pPr>{sect(columns)}</w:pPr></w:p>'


def break_para(text, before_twips, size_half_points=36, ppr_extra=""):
    rpr = f'<w:rPr><w:sz w:val="{size_half_points}"/></w:rPr>' if size_half_points else ''
    return (f'<w:p><w:pPr>{ppr_extra}<w:spacing w:before="{before_twips}"/>{rpr}</w:pPr>'
            f'<w:r><w:br w:type="column"/></w:r><w:r>{rpr}<w:t>{text}</w:t></w:r></w:p>')


def document(body, columns_last=1):
    return HEAD + f'<w:document {W}><w:body>{body}{sect(columns_last)}</w:body></w:document>'


def repeated(before_twips, fillers):
    body = ""
    for n in range(1, 5):
        body += para(f"Section {n}: three short lines in the left column, then a paragraph that opens with a column break.")
        body += section_end(1)
        body += "".join(para(f"Left column, line {i}.") for i in range(1, 4))
        body += break_para("Right column", before_twips)
        body += section_end(2)
    body += "".join(para(f"Filler line {i} in one column, below the last two-column section.") for i in range(1, fillers + 1))
    body += para("The last line of the document.")
    return document(body)


def list_number():
    num = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
    body = para("Top paragraph in one column.") + section_end(1)
    body += "".join(para(f"Left column item {i}.", num) for i in range(3))
    body += break_para("Right column item.", 0, size_half_points=0, ppr_extra=num)
    body += section_end(2)
    body += para("After the columns, one column again.")
    return document(body)


def write(path, document_xml, mode):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/numbering.xml", NUMBERING)
        z.writestr("word/settings.xml", settings(mode))
        z.writestr("word/document.xml", document_xml)


if __name__ == "__main__":
    out = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    out.mkdir(parents=True, exist_ok=True)
    write(out / "column-break-repro.docx", repeated(480, 48), 15)
    write(out / "column-break-line-mode14-repro.docx", repeated(0, 57), 14)
    write(out / "column-break-list-number-repro.docx", list_number(), 15)
