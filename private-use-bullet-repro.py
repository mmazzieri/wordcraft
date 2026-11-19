"""Synthetic reproducer: bullet levels whose private-use code the level's font lacks.

Writes a minimal DOCX with invented text. Usage: python genlist.py out.docx
"""
import sys
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# (font, level text, label)
CASES = [
    ("Symbol", "", "A: U+F0B7 in Symbol (Word's usual bullet)"),
    ("Symbol", "", "B: U+F080 in Symbol (Symbol has no glyph at 0x80)"),
    ("Wingdings", "", "C: U+F080 in Wingdings"),
    ("Times New Roman", "", "D: U+F0A7 in Times New Roman"),
    ("Not Installed Sans", "", "E: U+F081 in a font that is not installed"),
]


def level(font, text):
    return (
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
        f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}" w:hint="default"/></w:rPr></w:lvl>'
    )


def main(out):
    abstracts = "".join(
        f'<w:abstractNum w:abstractNumId="{i}"><w:multiLevelType w:val="hybridMultilevel"/>{level(f, t)}</w:abstractNum>'
        for i, (f, t, _) in enumerate(CASES)
    )
    nums = "".join(f'<w:num w:numId="{i + 1}"><w:abstractNumId w:val="{i}"/></w:num>' for i in range(len(CASES)))
    numbering = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {W}>{abstracts}{nums}</w:numbering>'
    body = ""
    for i, (_, _, label) in enumerate(CASES):
        for item in ("first item", "second item"):
            body += (
                f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{i + 1}"/></w:numPr></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>'
                f'<w:t xml:space="preserve">{label}: {item}</w:t></w:r></w:p>'
            )
    doc = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}'
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1417" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/>'
        "</w:sectPr></w:body></w:document>"
    )
    settings = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
        '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
        "</w:settings>"
    )
    ct = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
        "</Types>"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        "</Relationships>"
    )
    drels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{R}/settings" Target="settings.xml"/>'
        f'<Relationship Id="rId2" Type="{R}/numbering" Target="numbering.xml"/>'
        "</Relationships>"
    )
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/settings.xml", settings)
        z.writestr("word/numbering.xml", numbering)


if __name__ == "__main__":
    main(sys.argv[1])
