"""Synthetic reproducer: private-use symbol characters (U+F0xx) in runs.

Writes a minimal DOCX with invented text. Usage: python gen.py out.docx
"""
import sys
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

CASES = [
    ("Times New Roman", "", "A: U+F0A7 in Times New Roman"),
    ("Symbol", "", "B: U+F0A7 in Symbol"),
    ("Wingdings", "", "C: U+F0A7 in Wingdings"),
    ("Times New Roman", "", "D: U+F0B7 in Times New Roman"),
    ("Symbol", "", "E: U+F0B7 in Symbol"),
    ("Times New Roman", "", "F: U+F0FC in Times New Roman"),
    ("Wingdings", "", "G: U+F0FC in Wingdings"),
    ("Times New Roman", "", "H: U+F06F in Times New Roman"),
    ("Arial", "", "K: U+F0A7 in Arial"),
    ("Calibri", "", "L: U+F0A7 in Calibri"),
    (None, "", "M: U+F0A7 with no rFonts (theme font)"),
    ("Not Installed Sans", "", "N: U+F0A7 in a font that is not installed"),
    ("Segoe MDL2 Assets", "", "O: U+E700 in Segoe MDL2 Assets"),
    ("Times New Roman", "", "P: U+E700 in Times New Roman"),
]


def run(font, text):
    fonts = f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>' if font else ""
    return (
        f'<w:r><w:rPr>{fonts}'
        f'<w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'
    )


def para(font, ch, label):
    return (
        '<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>'
        + run(font, ch)
        + '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        '<w:sz w:val="24"/></w:rPr><w:tab/></w:r>'
        + run("Times New Roman", label)
        + "</w:p>"
    )


def sym_para(font, code, label):
    return (
        '<w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr>'
        f'<w:sym w:font="{font}" w:char="{code}"/></w:r>'
        '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:tab/></w:r>'
        + run("Times New Roman", label)
        + "</w:p>"
    )


def main(out):
    body = "".join(para(*c) for c in CASES)
    body += sym_para("Wingdings", "F0A7", "I: w:sym Wingdings F0A7")
    body += sym_para("Symbol", "F0B7", "J: w:sym Symbol F0B7")
    doc = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
        + body
        + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
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
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
        "</Relationships>"
    )
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc)
        z.writestr("word/_rels/document.xml.rels", drels)
        z.writestr("word/settings.xml", settings)


if __name__ == "__main__":
    main(sys.argv[1])
