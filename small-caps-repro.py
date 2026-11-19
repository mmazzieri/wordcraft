"""Small Caps spaces: a synthetic reproducer (invented text). Times New Roman Bold 10 pt."""
import docx
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement

TEXT = ("the reviewer should complete this panel only when the issuing office has said in the "
        "notice or in the papers quoted there that the reviewer may limit itself to this panel "
        "without completing any other panel of this part of the form")

def run(p, text, small=True):
    r = p.add_run(text)
    r.font.name = 'Times New Roman'; r.font.size = Pt(10); r.bold = True
    if small:
        r._r.get_or_add_rPr().append(OxmlElement('w:smallCaps'))
    return r

d = docx.Document()
s = d.sections[0]
s.page_width, s.page_height = Inches(8.5), Inches(11)
s.left_margin = s.right_margin = Inches(1)
for label, small in (("Small Caps", True), ("plain (for comparison)", False)):
    h = d.add_paragraph(); run(h, label + ":", small=False)
    p = d.add_paragraph(); p.paragraph_format.space_after = Pt(12)
    run(p, TEXT, small)
# 40 spaces between two letters: Word measures the gap at 80% in Small Caps
for small in (True, False):
    p = d.add_paragraph(); p.paragraph_format.space_after = Pt(0)
    run(p, "a" + " " * 40 + "a", small)
    p = d.add_paragraph(); p.paragraph_format.space_after = Pt(0)
    run(p, "a" + "\u00a0" * 40 + "a", small)
d.save('small-caps-repro.docx')
