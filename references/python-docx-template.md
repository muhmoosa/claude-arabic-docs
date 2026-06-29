# python-docx template for RTL Arabic documents

Most people generating Arabic `.docx` files from Python reach for **python-docx**.
Its high-level API is lossy for RTL: there is no `START` alignment in
`WD_ALIGN_PARAGRAPH` (so the natural call emits the buggy `<w:jc w:val="right"/>`
— see SKILL.md Rule 5), there is no API for `<w:bidi/>` at all, and `font.rtl`
is unreliable across versions. The practical conclusion: **treat the OOXML XML
layer as your real API.**

The helper layer below does exactly that. It encodes SKILL.md rules 1–8 *by
construction* — you write natural mixed Arabic/Latin prose and the tokenizer
splits it into correctly-flagged runs for you. Always finish with a
`harden_rtl.py` pass to backfill the document-level and section-level structure
that even this layer doesn't set (themeFontLang, docDefaults language, section
bidi).

## The helper module

```python
"""arabic-rtl-docs — python-docx helper layer.

Encodes SKILL.md rules 1-9.1 so the caller writes natural prose and gets
correct OOXML. After saving, ALWAYS run:  python harden_rtl.py out.docx

Key behaviors:
  * split_bidi follows Rule 7.5 — brackets, parens, and quotes stay in the RTL
    run; only strongly-LTR tokens (Latin letters/digits and URL/IBAN/range
    punctuation, including an interior en-dash) become LTR runs. So
    '(10 – 19)' produces one LTR run for '10 – 19' wrapped in RTL '(' / ')'
    that Word mirrors correctly.
  * add_run sets font pairing per Rule 9.1 — w:ascii/w:hAnsi = Times New Roman,
    w:cs = Traditional Arabic. Override per-call with latin_face / cs_face.

Usage:
    from docx import Document
    doc = Document()
    add_heading(doc, "تقرير حالة التطبيق", level=1)
    add_para(doc, "تطبيق Claude Desktop يعتمد على Bun JIT.")   # mixed Ar+Latin auto-split
    add_para(doc, "النطاق المعتمد (10 – 19) أو (20 فأكثر).")  # brackets stay RTL, range stays LTR
    add_code_block(doc, r"C:\\Users\\me\\claude.exe --debug")
    t = doc.add_table(rows=2, cols=2)
    make_table_rtl(t)
    fill_cell(t.rows[0].cells[0], "البند")
    doc.save("out.docx")
    # then: python harden_rtl.py out.docx --validate
"""

import re

from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# --- Eastern Arabic-Indic digits for Arabic prose (Rule 8) -------------------

_AR_DIGITS = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")


def ar_num(n):
    """123 -> '١٢٣'. Use for clause numbers, dates, currency in Arabic prose."""
    return str(n).translate(_AR_DIGITS)


# --- direct OOXML helpers (python-docx's high-level API is too lossy here) ---

def _set_jc(pPr, val):
    for jc in pPr.findall(qn("w:jc")):
        pPr.remove(jc)
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), val)
    pPr.append(jc)


def _set_bidi(pPr, on=True):
    existing = pPr.find(qn("w:bidi"))
    if on and existing is None:
        pPr.insert(0, OxmlElement("w:bidi"))   # bidi sorts early in the pPr schema
    elif not on and existing is not None:
        pPr.remove(existing)


def para_rtl(p):
    """RTL paragraph with logical start alignment (Rule 5: start, not right)."""
    pPr = p._p.get_or_add_pPr()
    _set_bidi(pPr, on=True)
    _set_jc(pPr, "start")


def para_center(p, rtl=True):
    pPr = p._p.get_or_add_pPr()
    _set_bidi(pPr, on=rtl)
    _set_jc(pPr, "center")


def para_ltr_left(p):
    """Pure LTR, left aligned — for code blocks, URLs, IBANs (Rule 7)."""
    pPr = p._p.get_or_add_pPr()
    _set_bidi(pPr, on=False)
    _set_jc(pPr, "left")


# --- bidi tokenizer: solves Rule 7.5 (bracket trap) by construction ---------

# A "strongly-LTR segment" per SKILL.md Rule 7.5: starts on a letter or digit,
# extends through Latin letters/digits and a small set of safe inline tokens
# (URL/IBAN/path/range characters, including an en-dash), and treats an interior
# space as part of the segment ONLY when it's followed by another LTR-safe char.
# Brackets, parentheses, quotes, currency symbols, and all other neutrals are
# deliberately NOT in this set — they belong in the RTL run so Word's bidi
# engine mirrors them.
_LTR_SEG_RE = re.compile(
    r"[A-Za-z0-9](?:[A-Za-z0-9@._\-/:+%,–]|[ ](?=[A-Za-z0-9@._\-/:+%,–]))*"
)


def split_bidi(text):
    """Split text into [(segment, is_rtl)] tuples per SKILL.md Rule 7.5.

    LTR matches come from _LTR_SEG_RE — they start on a letter/digit and keep
    going through Latin letters/digits, URL/IBAN punctuation, and an interior
    en-dash. Everything outside those matches (brackets, quotes, currency
    symbols, Arabic letters/punctuation) becomes RTL so Word's bidi engine
    mirrors brackets correctly.

    Examples:
        'تطبيق Claude يعتمد'   -> [('تطبيق ', True), ('Claude', False), (' يعتمد', True)]
        '(10 – 19)'             -> [('(', True), ('10 – 19', False), (')', True)]
        '(20 فأكثر)'             -> [('(', True), ('20', False), (' فأكثر)', True)]
        'user@example.com'      -> [('user@example.com', False)]
    """
    if not text:
        return []
    out, last = [], 0
    for m in _LTR_SEG_RE.finditer(text):
        if m.start() > last:
            out.append((text[last:m.start()], True))   # RTL gap (brackets, Arabic, neutrals)
        out.append((m.group(0), False))                 # strongly-LTR segment
        last = m.end()
    if last < len(text):
        out.append((text[last:], True))
    return out


# --- run builder with proper RTL + lang + font pairing per Rules 3, 9.1 -----

# Font pairing (Rule 9.1): Latin face on w:ascii/w:hAnsi, Arabic face on w:cs.
# Traditional Arabic looks much better than Arial for Arabic body text, and
# Times New Roman pairs cleanly with it for Latin runs. Override per-call with
# latin_face / cs_face if you need a different stack.
_DEFAULT_LATIN = "Times New Roman"
_DEFAULT_CS = "Traditional Arabic"
_CODE_FACE = "Consolas"


def _set_run_fonts(rPr, latin_face, cs_face):
    """Write <w:rFonts w:ascii w:hAnsi w:cs/> directly. python-docx's
    r.font.name only sets the Latin slots; Rule 9.1 needs w:cs too."""
    for old in rPr.findall(qn("w:rFonts")):
        rPr.remove(old)
    rf = OxmlElement("w:rFonts")
    rf.set(qn("w:ascii"), latin_face)
    rf.set(qn("w:hAnsi"), latin_face)
    rf.set(qn("w:cs"), cs_face)
    rPr.append(rf)


def add_run(p, text, *, rtl, size=11, bold=False, color=None, code=False,
            latin_face=None, cs_face=None):
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.bold = bold
    if color:
        r.font.color.rgb = color

    rPr = r._r.get_or_add_rPr()

    # Font pairing (Rule 9.1) — set Latin and complex-script faces independently.
    if code:
        _set_run_fonts(rPr, _CODE_FACE, _CODE_FACE)
    else:
        _set_run_fonts(rPr, latin_face or _DEFAULT_LATIN, cs_face or _DEFAULT_CS)

    # RTL flag on the run (Rule 3)
    if rtl and rPr.find(qn("w:rtl")) is None:
        rPr.append(OxmlElement("w:rtl"))

    # Language tags (Rule 3) — primary + complex-script
    for old in rPr.findall(qn("w:lang")):
        rPr.remove(old)
    lang = OxmlElement("w:lang")
    lang.set(qn("w:val"), "ar-SA" if rtl else "en-US")
    lang.set(qn("w:bidi"), "ar-SA")
    rPr.append(lang)

    # Inline code shading
    if code:
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "F2F2F2")
        rPr.append(shd)
    return r


def _emit_runs(p, text, **run_kw):
    """Tokenize text and emit one correctly-flagged run per script segment."""
    for seg, is_rtl in split_bidi(text):
        add_run(p, seg, rtl=is_rtl, **run_kw)


# --- the API the caller actually uses ---------------------------------------

def add_para(doc, text, *, size=11, bold=False, color=None,
             center=False, space_after=4):
    """Add an RTL paragraph. Mixed Ar+Latin text is auto-split into proper runs."""
    p = doc.add_paragraph()
    para_center(p, rtl=True) if center else para_rtl(p)
    p.paragraph_format.space_after = Pt(space_after)
    _emit_runs(p, text, size=size, bold=bold, color=color)
    return p


def add_heading(doc, text, *, level=1, size=None, space_before=12, space_after=8):
    """RTL heading. Avoids Word's built-in heading styles (which may not carry
    bidi unless you also patch the style — see Rule 6); writes a bold RTL para."""
    p = doc.add_paragraph()
    para_rtl(p)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    _emit_runs(p, text, size=size or (16 - 2 * level), bold=True)
    return p


def add_code_block(doc, code):
    """Pure LTR code block — gray background, monospace, left aligned."""
    p = doc.add_paragraph()
    para_ltr_left(p)
    p.paragraph_format.space_after = Pt(8)
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "F4F4F4")
    pPr.append(shd)
    for i, line in enumerate(code.split("\n")):
        if i > 0:
            p.add_run().add_break()
        add_run(p, line, rtl=False, size=9.5, code=True)
    return p


def make_table_rtl(table):
    """Inject <w:bidiVisual/> so columns flip in MS Word (Rule 4): the first
    cell in your code renders on the RIGHT."""
    tbl = table._element
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    if tblPr.find(qn("w:bidiVisual")) is None:
        tblPr.append(OxmlElement("w:bidiVisual"))


def fill_cell(cell, text, *, bold=False, size=11, color=None):
    """Write an RTL paragraph into a table cell, reusing its empty first
    paragraph (so you don't get a stray blank line above the text)."""
    p = cell.paragraphs[0]
    para_rtl(p)
    _emit_runs(p, text, size=size, bold=bold, color=color)
    return p
```

## Minimal worked example

```python
from docx import Document
from docx.shared import RGBColor

doc = Document()

# Title (centered, bold)
add_para(doc, "تقرير حالة تطبيق Claude Desktop", size=18, bold=True, center=True)

# Meta table — label | value, with columns flipped RTL
t = doc.add_table(rows=3, cols=2)
t.style = "Table Grid"
make_table_rtl(t)
rows = [
    ("التاريخ",   "١٩/٠٥/٢٠٢٦م"),
    ("الإصدار",   "v4.7.1"),
    ("المُعِدّ",   "فريق الهندسة"),
]
for (label, value), row in zip(rows, t.rows):
    fill_cell(row.cells[0], label, bold=True)
    fill_cell(row.cells[1], value)

# Body paragraph with mixed Arabic + Latin — auto-split into correct runs
add_para(doc, "يعتمد تطبيق Claude Desktop على محرك Bun JIT لتشغيل الإضافات.")

# A clause with Arabic-Indic digits
add_para(doc, f"راجع البند رقم ({ar_num(4)}) من وثيقة المتطلبات.")

# Rule 7.5 demo — brackets stay in the RTL run; the range stays as ONE LTR run
add_para(doc, "النطاق المعتمد (10 – 19) أو (20 فأكثر).")

# Pure-LTR code block
add_code_block(doc, r"C:\Users\me\AppData\claude\claude.exe --verbose")

# Reference with an LTR URL kept in its own run automatically
add_para(doc, "المصدر: https://docs.claude.com/skills للمزيد من التفاصيل.")

doc.save("out.docx")
```

## Always finish with the hardening pass

The helper layer sets paragraph-, run-, and table-level RTL correctly, but it
does **not** write the document-level identity that MS Word checks
(`themeFontLang` in `settings.xml`, the complex-script language in
`docDefaults`, section `<w:bidi/>`). Run the script to backfill those, then
validate:

```bash
python harden_rtl.py out.docx            # backfill structural RTL (themeFontLang, docDefaults, sectPr)
python harden_rtl.py out.docx --validate # confirm no content-level rule slipped through
```

A clean `--validate` run (0 errors) plus a known-good layout means the document
will render correctly in Microsoft Word, not just in a LibreOffice PDF preview.
