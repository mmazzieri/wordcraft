"""Synthetic repro: one font that isn't installed per paragraph, no w:altName, a different font
table family class each (swiss, modern, auto, decorative, script, not listed at all).

Invented text only. Usage: missing-font-families-repro.py <out.docx>
"""
import sys, zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
</Types>'''
RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
DOCRELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
</Relationships>'''

# (font name, font table family class, or None = no font table entry)
CASES = [('Missing Face Swiss', 'swiss'), ('Missing Face Modern', 'modern'), ('Missing Face Auto', 'auto'),
         ('Missing Face Decorative', 'decorative'), ('Missing Face Script', 'script'), ('Missing Face Unlisted', None)]

def run(font, text, b=False, i=False):
    props = f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>' + ('<w:b/>' if b else '') + ('<w:i/>' if i else '') + '<w:sz w:val="28"/>'
    return f'<w:r><w:rPr>{props}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'

body = ''.join(f'<w:p>{run(f, "Sphinx of black quartz ")}{run(f, "bold ", b=True)}{run(f, "italic", i=True)}</w:p>' for f, _ in CASES)
doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}<w:sectPr/></w:body></w:document>'
ft = ''.join(f'<w:font w:name="{f}"><w:charset w:val="00"/><w:family w:val="{fam}"/><w:pitch w:val="variable"/></w:font>' for f, fam in CASES if fam)
ft = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts {W}>{ft}</w:fonts>'
with zipfile.ZipFile(sys.argv[1], 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', CT); z.writestr('_rels/.rels', RELS)
    z.writestr('word/_rels/document.xml.rels', DOCRELS); z.writestr('word/document.xml', doc); z.writestr('word/fontTable.xml', ft)
