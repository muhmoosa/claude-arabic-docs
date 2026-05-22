#!/usr/bin/env python3
"""
harden_rtl.py — post-process a .docx so it renders correctly RTL in MS Word.

Most docx generators (docx-js, python-docx, openpyxl-via-converter) emit
paragraph-level RTL flags but skip section-level <w:bidi/>, table-level
<w:bidiVisual/>, per-run <w:rtl/>, AND the document-default complex-script
language declaration. The resulting file looks correct in LibreOffice but
Word renders headings, footers, table direction, and auto-numbering
left-to-right. This script patches the file:

  1. Adds <w:lang ... w:bidi="ar-SA"/> to docDefaults/rPrDefault/rPr in
     styles.xml. This is the SINGLE most important fix — without a
     complex-script language declaration at the document-default level,
     Word does NOT apply Arabic layout rules even when paragraphs are
     marked bidi. The default locale can be overridden with --locale.
  2. Adds <w:bidi/> to every <w:sectPr> that lacks it.
  3. Adds <w:bidiVisual/> to every <w:tblPr> that lacks it.
  4. Adds <w:bidi/> to every <w:pPr> whose paragraph contains RTL text
     but lacks the flag (skips paragraphs explicitly marked left-aligned
     with no RTL runs — i.e. pure LTR content).
  5. Adds <w:rtl/> to every <w:rPr> whose run text contains RTL chars
     but lacks the flag.
  6. Also applies the same patches to header*.xml and footer*.xml.

The script is idempotent — running it twice is a no-op on the second pass.

It also offers a separate --validate mode that catches *content-level* RTL
mistakes the structural hardening can't auto-fix (see SKILL.md rules 5, 7, 8):
a paragraph that is bidi but uses jc="right" instead of "start", a single run
that mixes Arabic and Latin letters, Latin digits inside an RTL run, and Latin
punctuation trailing an RTL-only run. These pass silently in a LibreOffice PDF
preview but still render wrong in MS Word. --validate reports them and exits 1
on errors; the opt-in --fix-jc rewrites jc="right" -> "start" where it is
unambiguously safe to do so.

Usage:
    python harden_rtl.py document.docx                       # in place (structural harden)
    python harden_rtl.py in.docx -o out.docx                 # to new file
    python harden_rtl.py doc.docx --report                   # show what changed
    python harden_rtl.py doc.docx --locale he-IL             # Hebrew instead of Arabic
    python harden_rtl.py doc.docx --validate                 # report content issues, exit 1 on errors
    python harden_rtl.py doc.docx --validate --fix-jc        # also rewrite right -> start where safe

Dependencies: standard library only.
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Unicode ranges that count as "RTL script". We use a broad set so this also
# helps Hebrew, Syriac, Thaana, NKo, Mandaic, etc.
_RTL_RANGES = [
    (0x0590, 0x05FF),  # Hebrew
    (0x0600, 0x06FF),  # Arabic
    (0x0700, 0x074F),  # Syriac
    (0x0750, 0x077F),  # Arabic Supplement
    (0x0780, 0x07BF),  # Thaana
    (0x07C0, 0x07FF),  # NKo
    (0x0800, 0x083F),  # Samaritan
    (0x0840, 0x085F),  # Mandaic
    (0x08A0, 0x08FF),  # Arabic Extended-A
    (0xFB1D, 0xFDFF),  # Hebrew/Arabic Presentation Forms-A
    (0xFE70, 0xFEFF),  # Arabic Presentation Forms-B
]


def _has_rtl(text: str) -> bool:
    for ch in text:
        cp = ord(ch)
        for lo, hi in _RTL_RANGES:
            if lo <= cp <= hi:
                return True
    return False


@dataclass
class Report:
    settings_themefontlang_added: int = 0
    doc_default_lang_added: int = 0
    doc_default_ppr_added: int = 0
    normal_style_bidi_added: int = 0
    sect_bidi_added: int = 0
    tbl_bidivisual_added: int = 0
    p_bidi_added: int = 0
    r_rtl_added: int = 0
    files_changed: list[str] = field(default_factory=list)

    def total(self) -> int:
        return (
            self.settings_themefontlang_added
            + self.doc_default_lang_added
            + self.doc_default_ppr_added
            + self.normal_style_bidi_added
            + self.sect_bidi_added
            + self.tbl_bidivisual_added
            + self.p_bidi_added
            + self.r_rtl_added
        )

    def summary(self) -> str:
        lines = [
            f"settings.xml themeFontLang bidi:  {self.settings_themefontlang_added}",
            f"docDefaults complex-script lang:  {self.doc_default_lang_added}",
            f"docDefaults pPrDefault bidi/jc:   {self.doc_default_ppr_added}",
            f"Normal style bidi/jc added:       {self.normal_style_bidi_added}",
            f"Section <w:bidi/> added:          {self.sect_bidi_added}",
            f"Table <w:bidiVisual/> added:      {self.tbl_bidivisual_added}",
            f"Paragraph <w:bidi/> added:        {self.p_bidi_added}",
            f"Run <w:rtl/> added:               {self.r_rtl_added}",
            f"Files touched:                    {len(self.files_changed)}",
        ]
        if self.files_changed:
            lines.append("  " + "\n  ".join(self.files_changed))
        return "\n".join(lines)


# ---------- XML patching ----------

# Order inside <w:pPr> matters per the OOXML schema. We insert <w:bidi/> as
# late as possible while still preserving the canonical order. The simplest
# correct rule is: put it right before <w:rPr> if present, otherwise at the
# end of <w:pPr>. Word is tolerant beyond that.

_W_RE_PPR_OPEN = re.compile(r"<w:pPr\b[^/>]*>")
_W_RE_PPR_EMPTY = re.compile(r"<w:pPr\b[^>]*?/>")
_W_RE_RPR_OPEN = re.compile(r"<w:rPr\b[^/>]*>")
_W_RE_RPR_EMPTY = re.compile(r"<w:rPr\b[^>]*?/>")

_TAG_NAME = r"[A-Za-z][A-Za-z0-9._:-]*"


def _add_child_to_block(
    xml: str,
    open_re: re.Pattern,
    empty_re: re.Pattern,
    parent_close: str,
    child: str,
    *,
    before: str | None = None,
) -> tuple[str, bool]:
    """Insert `child` inside the first matching parent block in `xml`.

    Returns (new_xml, changed). If `before` is given (e.g. "<w:rPr"), the
    child is inserted before that tag if present, otherwise right before
    the parent's close tag.
    """
    m = empty_re.search(xml)
    if m:
        # <w:pPr/> → expand to <w:pPr><child/></w:pPr>
        full = m.group(0)
        open_tag = full[:-2] + ">"  # strip "/>" -> ">"
        replacement = f"{open_tag}{child}{parent_close}"
        return xml[: m.start()] + replacement + xml[m.end():], True

    m = open_re.search(xml)
    if not m:
        return xml, False

    # Find the matching close tag for this parent
    close_idx = xml.find(parent_close, m.end())
    if close_idx == -1:
        return xml, False

    insertion_point = close_idx
    if before:
        # Insert before the first occurrence of `before` between m.end() and close_idx
        sub = xml[m.end():close_idx]
        rel = sub.find(before)
        if rel != -1:
            insertion_point = m.end() + rel

    return xml[:insertion_point] + child + xml[insertion_point:], True


def _insert_at_first_marker(
    parent_xml: str,
    open_re: re.Pattern,
    empty_re: re.Pattern,
    close_tag: str,
    child: str,
    *,
    markers: list[str],
) -> tuple[str, bool]:
    """Insert child before the first of `markers` found in the parent block.

    If none of the markers are present, insert just before the close tag.
    This respects OOXML schema ordering, which Word's validator enforces.
    """
    for marker in markers:
        if marker in parent_xml:
            return _add_child_to_block(
                parent_xml, open_re, empty_re, close_tag, child, before=marker,
            )
    return _add_child_to_block(parent_xml, open_re, empty_re, close_tag, child)


# Schema-order markers (children that must come AFTER our injected tag).
_PPR_AFTER_BIDI = ["<w:adjustRightInd", "<w:snapToGrid", "<w:spacing", "<w:ind",
                   "<w:contextualSpacing", "<w:mirrorIndents", "<w:suppressOverlap",
                   "<w:jc", "<w:textDirection", "<w:textAlignment", "<w:outlineLvl",
                   "<w:rPr"]
_RPR_AFTER_RTL = ["<w:cs", "<w:em", "<w:lang", "<w:eastAsianLayout", "<w:specVanish",
                  "<w:oMath", "<w:rPrChange"]
_TBLPR_AFTER_BIDIVISUAL = ["<w:tblStyleRowBandSize", "<w:tblStyleColBandSize",
                           "<w:tblW", "<w:jc", "<w:tblCellSpacing", "<w:tblInd",
                           "<w:tblBorders", "<w:tblShd", "<w:tblLayout",
                           "<w:tblCellMar", "<w:tblLook", "<w:tblCaption",
                           "<w:tblDescription", "<w:tblPrChange"]


def _ensure_pPr_has_bidi(p_xml: str) -> tuple[str, bool]:
    """Ensure the <w:pPr> inside this <w:p>...</w:p> has <w:bidi/>."""
    if "<w:bidi/>" in p_xml or "<w:bidi " in p_xml:
        return p_xml, False

    if "<w:pPr" not in p_xml:
        m = re.search(r"<w:p\b[^>]*>", p_xml)
        if not m:
            return p_xml, False
        insert = "<w:pPr><w:bidi/></w:pPr>"
        return p_xml[: m.end()] + insert + p_xml[m.end():], True

    return _insert_at_first_marker(
        p_xml, _W_RE_PPR_OPEN, _W_RE_PPR_EMPTY, "</w:pPr>", "<w:bidi/>",
        markers=_PPR_AFTER_BIDI,
    )


def _ensure_rPr_has_rtl(r_xml: str) -> tuple[str, bool]:
    """Ensure the <w:rPr> inside this <w:r>...</w:r> has <w:rtl/>."""
    if "<w:rtl/>" in r_xml or "<w:rtl " in r_xml:
        return r_xml, False

    if "<w:rPr" not in r_xml:
        m = re.search(r"<w:r\b[^>]*>", r_xml)
        if not m:
            return r_xml, False
        insert = "<w:rPr><w:rtl/></w:rPr>"
        return r_xml[: m.end()] + insert + r_xml[m.end():], True

    return _insert_at_first_marker(
        r_xml, _W_RE_RPR_OPEN, _W_RE_RPR_EMPTY, "</w:rPr>", "<w:rtl/>",
        markers=_RPR_AFTER_RTL,
    )


def _ensure_sectPr_has_bidi(sect_xml: str) -> tuple[str, bool]:
    """Inject <w:bidi/> inside a <w:sectPr>...</w:sectPr> block.

    OOXML schema places <w:bidi> after <w:textDirection> and before
    <w:rtlGutter>/<w:docGrid>/<w:printerSettings>/<w:sectPrChange>.
    Appending at the end of <w:sectPr> violates the schema and trips
    Word's validator.
    """
    if "<w:bidi/>" in sect_xml or "<w:bidi " in sect_xml:
        return sect_xml, False

    open_re = re.compile(r"<w:sectPr\b[^/>]*>")
    empty_re = re.compile(r"<w:sectPr\b[^>]*?/>")
    # Insert before the first of these later-in-schema tags if present.
    for marker in ("<w:rtlGutter", "<w:docGrid", "<w:printerSettings", "<w:sectPrChange"):
        if marker in sect_xml:
            return _add_child_to_block(
                sect_xml,
                open_re,
                empty_re,
                "</w:sectPr>",
                "<w:bidi/>",
                before=marker,
            )
    return _add_child_to_block(
        sect_xml,
        open_re,
        empty_re,
        "</w:sectPr>",
        "<w:bidi/>",
    )


def _ensure_tblPr_has_bidiVisual(tbl_xml: str) -> tuple[str, bool]:
    """Inject <w:bidiVisual/> inside the <w:tblPr> of a table block."""
    if "<w:bidiVisual/>" in tbl_xml or "<w:bidiVisual " in tbl_xml:
        return tbl_xml, False

    # Find <w:tblPr>...</w:tblPr> or <w:tblPr/>
    if "<w:tblPr" not in tbl_xml:
        # Insert a minimal <w:tblPr><w:bidiVisual/></w:tblPr> right after <w:tbl ...>
        m = re.search(r"<w:tbl\b[^>]*>", tbl_xml)
        if not m:
            return tbl_xml, False
        insert = "<w:tblPr><w:bidiVisual/></w:tblPr>"
        return tbl_xml[: m.end()] + insert + tbl_xml[m.end():], True

    open_re = re.compile(r"<w:tblPr\b[^/>]*>")
    empty_re = re.compile(r"<w:tblPr\b[^>]*?/>")
    return _insert_at_first_marker(
        tbl_xml, open_re, empty_re, "</w:tblPr>", "<w:bidiVisual/>",
        markers=_TBLPR_AFTER_BIDIVISUAL,
    )


# ---------- Block iteration ----------

def _iter_blocks(xml: str, open_tag: str, close_tag: str):
    """Yield (start, end, block_text) tuples for each <open_tag ...>...</close_tag>.

    Handles nesting correctly via simple counter.
    """
    i = 0
    open_re = re.compile(rf"<{open_tag}\b[^/>]*>")
    empty_re = re.compile(rf"<{open_tag}\b[^>]*?/>")
    close_str = f"</{close_tag}>"

    while i < len(xml):
        m_open = open_re.search(xml, i)
        m_empty = empty_re.search(xml, i)

        # Pick whichever comes first
        candidates = []
        if m_open:
            candidates.append((m_open.start(), "open", m_open))
        if m_empty:
            candidates.append((m_empty.start(), "empty", m_empty))
        if not candidates:
            return
        candidates.sort()
        pos, kind, m = candidates[0]

        if kind == "empty":
            yield m.start(), m.end(), m.group(0)
            i = m.end()
            continue

        # Find matching close, accounting for nested same-tag opens
        depth = 1
        scan = m.end()
        while depth > 0:
            n_open = open_re.search(xml, scan)
            n_close = xml.find(close_str, scan)
            if n_close == -1:
                return  # malformed; bail
            if n_open and n_open.start() < n_close:
                depth += 1
                scan = n_open.end()
            else:
                depth -= 1
                scan = n_close + len(close_str)
        end = scan
        yield m.start(), end, xml[m.start():end]
        i = end


def _replace_blocks(xml: str, open_tag: str, close_tag: str, transform) -> tuple[str, int]:
    """Apply `transform(block_xml) -> (new_block_xml, changed_bool)` to each block.

    Walks right-to-left so indexes stay valid as we splice.
    """
    blocks = list(_iter_blocks(xml, open_tag, close_tag))
    changes = 0
    for start, end, block in reversed(blocks):
        new_block, changed = transform(block)
        if changed:
            xml = xml[:start] + new_block + xml[end:]
            changes += 1
    return xml, changes


# ---------- Run-level helpers ----------

# Extract concatenated text of a run by grabbing <w:t>...</w:t> contents.
_W_T_RE = re.compile(r"<w:t\b[^>]*>(.*?)</w:t>", re.DOTALL)


def _run_text(r_xml: str) -> str:
    return "".join(m.group(1) for m in _W_T_RE.finditer(r_xml))


# ---------- Per-file driver ----------

def _harden_xml(xml: str, report: Report) -> str:
    # 1. Sections — inject <w:bidi/> into every <w:sectPr>
    def _xform_sect(sect_xml: str):
        return _ensure_sectPr_has_bidi(sect_xml)

    xml, n = _replace_blocks(xml, "w:sectPr", "w:sectPr", _xform_sect)
    report.sect_bidi_added += n

    # 2. Tables — inject <w:bidiVisual/> into every <w:tbl>'s <w:tblPr>
    def _xform_tbl(tbl_xml: str):
        return _ensure_tblPr_has_bidiVisual(tbl_xml)

    xml, n = _replace_blocks(xml, "w:tbl", "w:tbl", _xform_tbl)
    report.tbl_bidivisual_added += n

    # 3. Paragraphs — inject <w:bidi/> into <w:pPr> when the paragraph has RTL
    #    text but lacks the flag. Skip purely-LTR paragraphs (no RTL anywhere
    #    AND explicitly left-aligned) — those are e.g. English contact lines.
    def _xform_p(p_xml: str):
        runs_text = "".join(_W_T_RE.findall(p_xml))
        if not runs_text:
            # Empty paragraphs (spacers, signature lines with only dots) — leave alone.
            return p_xml, False
        if not _has_rtl(runs_text):
            # Pure LTR content — leave alone.
            return p_xml, False
        return _ensure_pPr_has_bidi(p_xml)

    xml, n = _replace_blocks(xml, "w:p", "w:p", _xform_p)
    report.p_bidi_added += n

    # 4. Runs — inject <w:rtl/> into <w:rPr> when the run text has RTL chars
    def _xform_r(r_xml: str):
        text = _run_text(r_xml)
        if not text or not _has_rtl(text):
            return r_xml, False
        return _ensure_rPr_has_rtl(r_xml)

    xml, n = _replace_blocks(xml, "w:r", "w:r", _xform_r)
    report.r_rtl_added += n

    return xml


def _is_target_file(name: str) -> bool:
    base = os.path.basename(name)
    if base == "document.xml":
        return True
    if base.startswith("header") and base.endswith(".xml"):
        return True
    if base.startswith("footer") and base.endswith(".xml"):
        return True
    return False


# Map of locales accepted by Word's w:lang w:bidi attribute. Add more as needed.
_BIDI_LOCALES = {
    "ar-SA": ("ar-SA", "Arabic (Saudi Arabia)"),
    "ar-EG": ("ar-EG", "Arabic (Egypt)"),
    "ar-AE": ("ar-AE", "Arabic (UAE)"),
    "ar-MA": ("ar-MA", "Arabic (Morocco)"),
    "he-IL": ("he-IL", "Hebrew (Israel)"),
    "fa-IR": ("fa-IR", "Persian (Iran)"),
    "ur-PK": ("ur-PK", "Urdu (Pakistan)"),
}


def _harden_styles_lang(xml: str, locale: str, report: Report) -> str:
    """Inject <w:lang ... w:bidi="<locale>"/> into docDefaults/rPrDefault/rPr.

    This is the single most important fix for getting Word to treat the
    document as RTL. Without a docDefaults complex-script language, Word
    ignores paragraph-level bidi flags and renders the doc LTR.
    """
    # If an existing w:lang with w:bidi="<some RTL locale>" is present, leave it.
    existing = re.search(r'<w:lang\b[^/]*w:bidi="([a-zA-Z-]+)"[^>]*?/>', xml)
    if existing:
        return xml

    desired_lang = f'<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="{locale}"/>'

    # Path A: docDefaults/rPrDefault/rPr exists and is non-empty — append <w:lang> to it.
    m_full = re.search(
        r'(<w:docDefaults\b[^>]*>\s*<w:rPrDefault\b[^>]*>\s*<w:rPr\b[^>]*>)'
        r'(.*?)'
        r'(</w:rPr>\s*</w:rPrDefault>)',
        xml, re.DOTALL,
    )
    if m_full:
        # If there's an existing <w:lang/> without a bidi attribute, replace it.
        inner = m_full.group(2)
        m_lang = re.search(r'<w:lang\b[^>]*?/>', inner)
        if m_lang:
            old_lang = m_lang.group(0)
            # Merge: pull val and eastAsia if present, add bidi=
            val = re.search(r'w:val="([^"]*)"', old_lang)
            ea = re.search(r'w:eastAsia="([^"]*)"', old_lang)
            new_lang = '<w:lang'
            new_lang += f' w:val="{val.group(1) if val else "en-US"}"'
            new_lang += f' w:eastAsia="{ea.group(1) if ea else "en-US"}"'
            new_lang += f' w:bidi="{locale}"/>'
            new_inner = inner[:m_lang.start()] + new_lang + inner[m_lang.end():]
        else:
            new_inner = inner + desired_lang
        new_block = m_full.group(1) + new_inner + m_full.group(3)
        xml = xml[:m_full.start()] + new_block + xml[m_full.end():]
        report.doc_default_lang_added += 1
        return xml

    # Path B: docDefaults/rPrDefault exists but rPr is self-closing or missing.
    m_self = re.search(
        r'(<w:docDefaults\b[^>]*>\s*<w:rPrDefault\b[^>]*>)\s*<w:rPr\s*/>\s*(</w:rPrDefault>)',
        xml, re.DOTALL,
    )
    if m_self:
        new_block = m_self.group(1) + f'<w:rPr>{desired_lang}</w:rPr>' + m_self.group(2)
        xml = xml[:m_self.start()] + new_block + xml[m_self.end():]
        report.doc_default_lang_added += 1
        return xml

    m_no_rpr = re.search(
        r'(<w:docDefaults\b[^>]*>\s*<w:rPrDefault\b[^>]*>)\s*(</w:rPrDefault>)',
        xml, re.DOTALL,
    )
    if m_no_rpr:
        new_block = m_no_rpr.group(1) + f'<w:rPr>{desired_lang}</w:rPr>' + m_no_rpr.group(2)
        xml = xml[:m_no_rpr.start()] + new_block + xml[m_no_rpr.end():]
        report.doc_default_lang_added += 1
        return xml

    # Path C: docDefaults exists but no rPrDefault — add one.
    m_no_rprdef = re.search(r'(<w:docDefaults\b[^>]*>)', xml)
    if m_no_rprdef:
        insertion = f'<w:rPrDefault><w:rPr>{desired_lang}</w:rPr></w:rPrDefault>'
        xml = xml[:m_no_rprdef.end()] + insertion + xml[m_no_rprdef.end():]
        report.doc_default_lang_added += 1
        return xml

    # Path D: no docDefaults at all — add one near the start of <w:styles>.
    m_styles = re.search(r'<w:styles\b[^>]*>', xml)
    if m_styles:
        insertion = (
            f'<w:docDefaults><w:rPrDefault><w:rPr>{desired_lang}</w:rPr></w:rPrDefault>'
            f'<w:pPrDefault/></w:docDefaults>'
        )
        xml = xml[:m_styles.end()] + insertion + xml[m_styles.end():]
        report.doc_default_lang_added += 1

    return xml


def _harden_styles_ppr_default(xml: str, report: Report) -> str:
    """Ensure docDefaults/pPrDefault contains <w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>.

    Without this, Word's default paragraph style ("Normal") has no reading
    direction set. Word then renders headings and any paragraph that lacks
    an EXPLICIT bidi flag using the doc-default LTR layout — even when the
    section has <w:bidi/>. The symptom is: tables flow RTL correctly but
    headings and body paragraphs render left-aligned in MS Word.

    LibreOffice doesn't show this bug, so it slips past PDF preview checks.
    """
    desired_inner = '<w:bidi/><w:jc w:val="right"/>'

    # Case A: <w:pPrDefault/> is self-closing — expand it.
    m_self = re.search(r'<w:pPrDefault\s*/>', xml)
    if m_self:
        replacement = f'<w:pPrDefault><w:pPr>{desired_inner}</w:pPr></w:pPrDefault>'
        xml = xml[:m_self.start()] + replacement + xml[m_self.end():]
        report.doc_default_ppr_added += 1
        return xml

    # Case B: <w:pPrDefault> exists but its <w:pPr> is missing/empty.
    m_no_ppr = re.search(
        r'(<w:pPrDefault\b[^>]*>)\s*(</w:pPrDefault>)', xml, re.DOTALL,
    )
    if m_no_ppr:
        replacement = m_no_ppr.group(1) + f'<w:pPr>{desired_inner}</w:pPr>' + m_no_ppr.group(2)
        xml = xml[:m_no_ppr.start()] + replacement + xml[m_no_ppr.end():]
        report.doc_default_ppr_added += 1
        return xml

    m_self_ppr = re.search(
        r'(<w:pPrDefault\b[^>]*>)\s*<w:pPr\s*/>\s*(</w:pPrDefault>)', xml, re.DOTALL,
    )
    if m_self_ppr:
        replacement = m_self_ppr.group(1) + f'<w:pPr>{desired_inner}</w:pPr>' + m_self_ppr.group(2)
        xml = xml[:m_self_ppr.start()] + replacement + xml[m_self_ppr.end():]
        report.doc_default_ppr_added += 1
        return xml

    # Case C: <w:pPrDefault><w:pPr>…existing children…</w:pPr></w:pPrDefault>
    m_full = re.search(
        r'(<w:pPrDefault\b[^>]*>\s*<w:pPr\b[^>]*>)(.*?)(</w:pPr>\s*</w:pPrDefault>)',
        xml, re.DOTALL,
    )
    if m_full:
        inner = m_full.group(2)
        has_bidi = '<w:bidi/>' in inner or '<w:bidi ' in inner
        has_jc = '<w:jc ' in inner or '<w:jc/>' in inner
        if has_bidi and has_jc:
            return xml  # already complete

        add = ''
        if not has_bidi:
            add += '<w:bidi/>'
        if not has_jc:
            add += '<w:jc w:val="right"/>'
        # Prepend our additions so they appear in schema-correct early positions.
        new_inner = add + inner
        new_block = m_full.group(1) + new_inner + m_full.group(3)
        xml = xml[:m_full.start()] + new_block + xml[m_full.end():]
        report.doc_default_ppr_added += 1
        return xml

    # Case D: no <w:pPrDefault> at all — add one inside <w:docDefaults>.
    m_dd = re.search(r'(<w:docDefaults\b[^>]*>)', xml)
    if m_dd:
        # Find closing tag
        close = xml.find('</w:docDefaults>', m_dd.end())
        if close != -1:
            insertion = f'<w:pPrDefault><w:pPr>{desired_inner}</w:pPr></w:pPrDefault>'
            xml = xml[:close] + insertion + xml[close:]
            report.doc_default_ppr_added += 1
            return xml

    return xml


def _harden_settings_themefontlang(xml: str, locale: str, report: Report) -> str:
    """Ensure settings.xml has <w:themeFontLang w:bidi="<locale>"/>.

    This is THE setting MS Word actually checks when deciding whether to engage
    its RTL rendering pipeline. Without it, even a fully-bidi'd document (every
    paragraph, run, table, and section RTL-flagged, docDefaults lang set) still
    renders left-to-right in Word.

    The other settings tell Word the document *contains* RTL text;
    themeFontLang tells Word the document *is* an RTL document and to engage
    the matching layout engine. The distinction matters: tables get RTL via
    bidiVisual unconditionally, but headings and body paragraphs only get
    proper RTL alignment when the theme font language confirms the doc-wide
    direction.

    The schema order inside <w:settings> places <w:themeFontLang> early —
    after <w:zoom>, <w:displayBackgroundShape>, etc. — but Word tolerates
    looser placement.
    """
    existing = re.search(r'<w:themeFontLang\b[^/]*w:bidi="([a-zA-Z-]+)"[^/]*/>', xml)
    if existing:
        return xml

    desired = f'<w:themeFontLang w:val="en-US" w:eastAsia="en-US" w:bidi="{locale}"/>'

    # Path A: themeFontLang exists but without w:bidi — augment it.
    m_existing = re.search(r'<w:themeFontLang\b([^/]*)/>', xml)
    if m_existing:
        attrs = m_existing.group(1)
        val = re.search(r'w:val="([^"]*)"', attrs)
        ea = re.search(r'w:eastAsia="([^"]*)"', attrs)
        new_tag = '<w:themeFontLang'
        new_tag += f' w:val="{val.group(1) if val else "en-US"}"'
        new_tag += f' w:eastAsia="{ea.group(1) if ea else "en-US"}"'
        new_tag += f' w:bidi="{locale}"/>'
        xml = xml[:m_existing.start()] + new_tag + xml[m_existing.end():]
        report.settings_themefontlang_added += 1
        return xml

    # Path B: no themeFontLang yet — insert it in the SCHEMA-CORRECT position.
    # Per OOXML, <w:themeFontLang> sits late in <w:settings>: after <w:mathPr>,
    # <w:rsids>, <w:attachedSchema>, and before <w:clrSchemeMapping>. The safest
    # approach is to insert immediately BEFORE the first of these later elements,
    # falling back to just before </w:settings>.
    m_settings = re.search(r'<w:settings\b[^>]*>', xml)
    if m_settings:
        # Markers that, per schema, come AFTER <w:themeFontLang>. Insert before
        # the first one we find.
        later_markers = [
            '<w:clrSchemeMapping',
            '<w:doNotIncludeSubdocsInStats',
            '<w:doNotAutoCompressPictures',
            '<w:forceUpgrade',
            '<w:captions',
            '<w:readModeInkLockDown',
            '<w:smartTagType',
            '<w:schemaLibrary',
            '<w:doNotEmbedSmartTags',
            '<w:decimalSymbol',
            '<w:listSeparator',
        ]
        insertion_point = None
        for marker in later_markers:
            idx = xml.find(marker, m_settings.end())
            if idx != -1 and (insertion_point is None or idx < insertion_point):
                insertion_point = idx

        if insertion_point is None:
            # No later marker found — insert just before </w:settings>.
            close = xml.rfind('</w:settings>')
            if close != -1:
                insertion_point = close

        if insertion_point is not None:
            xml = xml[:insertion_point] + desired + xml[insertion_point:]
            report.settings_themefontlang_added += 1
        return xml

    # Path C: no <w:settings> root at all — should never happen in a valid docx
    # but if so, we can't safely create one here without breaking relationships.
    return xml


def _harden_normal_style(xml: str, report: Report) -> str:
    """Ensure the Normal paragraph style (if it exists) carries bidi and right-jc.

    The Normal style is the parent of every other paragraph style by default.
    If it lacks bidi, all derived styles can fall back to LTR when their own
    properties don't override.
    """
    m_style = re.search(
        r'(<w:style\b[^>]*w:styleId="Normal"[^>]*>)(.*?)(</w:style>)',
        xml, re.DOTALL,
    )
    if not m_style:
        return xml

    body = m_style.group(2)
    # Find <w:pPr>...</w:pPr> inside the style body, or self-closing <w:pPr/>
    m_ppr_full = re.search(r'<w:pPr\b[^>]*>(.*?)</w:pPr>', body, re.DOTALL)
    m_ppr_self = re.search(r'<w:pPr\s*/>', body)

    if m_ppr_full:
        inner = m_ppr_full.group(1)
        has_bidi = '<w:bidi/>' in inner or '<w:bidi ' in inner
        has_jc = '<w:jc ' in inner
        if has_bidi and has_jc:
            return xml
        add = ''
        if not has_bidi:
            add += '<w:bidi/>'
        if not has_jc:
            add += '<w:jc w:val="right"/>'
        new_inner = add + inner
        new_body = body[:m_ppr_full.start()] + f'<w:pPr>{new_inner}</w:pPr>' + body[m_ppr_full.end():]
    elif m_ppr_self:
        new_body = body[:m_ppr_self.start()] + f'<w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>' + body[m_ppr_self.end():]
    else:
        # No <w:pPr> at all — insert one after the opening tag (which is between groups 1 and 2).
        new_body = f'<w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>' + body

    new_style = m_style.group(1) + new_body + m_style.group(3)
    xml = xml[:m_style.start()] + new_style + xml[m_style.end():]
    report.normal_style_bidi_added += 1
    return xml


def harden_docx(in_path: Path, out_path: Path | None = None, *, locale: str = "ar-SA") -> Report:
    in_path = Path(in_path)
    if not in_path.exists():
        raise FileNotFoundError(in_path)

    report = Report()
    target = out_path or in_path

    if locale not in _BIDI_LOCALES:
        raise ValueError(f"Unsupported locale {locale!r}. Supported: {', '.join(sorted(_BIDI_LOCALES))}")

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        scratch = tmp / "scratch.docx"

        with zipfile.ZipFile(in_path, "r") as zin, zipfile.ZipFile(
            scratch, "w", compression=zipfile.ZIP_DEFLATED
        ) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                base = os.path.basename(info.filename)
                if _is_target_file(info.filename):
                    text = data.decode("utf-8")
                    before = (
                        report.sect_bidi_added,
                        report.tbl_bidivisual_added,
                        report.p_bidi_added,
                        report.r_rtl_added,
                    )
                    text = _harden_xml(text, report)
                    after = (
                        report.sect_bidi_added,
                        report.tbl_bidivisual_added,
                        report.p_bidi_added,
                        report.r_rtl_added,
                    )
                    if before != after:
                        report.files_changed.append(info.filename)
                    data = text.encode("utf-8")
                elif base == "styles.xml":
                    text = data.decode("utf-8")
                    before = (
                        report.doc_default_lang_added,
                        report.doc_default_ppr_added,
                        report.normal_style_bidi_added,
                    )
                    text = _harden_styles_lang(text, locale, report)
                    text = _harden_styles_ppr_default(text, report)
                    text = _harden_normal_style(text, report)
                    after = (
                        report.doc_default_lang_added,
                        report.doc_default_ppr_added,
                        report.normal_style_bidi_added,
                    )
                    if before != after:
                        report.files_changed.append(info.filename)
                    data = text.encode("utf-8")
                elif base == "settings.xml":
                    text = data.decode("utf-8")
                    before = report.settings_themefontlang_added
                    text = _harden_settings_themefontlang(text, locale, report)
                    if report.settings_themefontlang_added != before:
                        report.files_changed.append(info.filename)
                    data = text.encode("utf-8")
                zout.writestr(info, data)

        shutil.move(str(scratch), str(target))

    return report


# ---------- Content-level validation (SKILL.md rules 5, 7, 8) ----------
#
# harden_docx() fixes *structural* OOXML (themeFontLang, docDefaults, section
# bidi, table bidiVisual, paragraph bidi, run rtl). That is necessary but not
# sufficient: a generator can still emit content-level mistakes that render
# wrong in Word even when the structure is perfect. Those mistakes pass
# silently because LibreOffice's PDF preview infers direction from the runs and
# looks fine. The validator below catches them.
#
# Scope: only the content parts (document.xml, header*.xml, footer*.xml).
# styles.xml is deliberately NOT scanned. The document defaults legitimately
# carry <w:jc w:val="right"/> per SKILL.md rule 0.5 (the harden step injects it
# there on purpose), so flagging those would be a false positive. Rule 5's
# "use start, not right" applies to *body* paragraphs, which live in the
# content parts.

_LATIN_LETTER_RE = re.compile(r"[A-Za-z]")
_LATIN_DIGIT_RE = re.compile(r"[0-9]")
_TRAILING_LATIN_PUNCT = (",", ";", "?")

_PPR_BLOCK_RE = re.compile(r"<w:pPr\b[^>]*>.*?</w:pPr>|<w:pPr\b[^>]*?/>", re.DOTALL)
_SECTPR_BLOCK_RE = re.compile(r"<w:sectPr\b[^>]*>.*?</w:sectPr>", re.DOTALL)
_JC_RIGHT_RE = re.compile(r'<w:jc\b[^>]*\bw:val="right"[^>]*/>')


@dataclass
class Issue:
    severity: str  # "ERROR" or "WARN"
    code: str      # short, machine-readable
    message: str
    where: str = ""

    def __str__(self) -> str:
        loc = f"  [{self.where}]" if self.where else ""
        return f"{self.severity:5s} {self.code}: {self.message}{loc}"


def _has_latin_letter(text: str) -> bool:
    return bool(_LATIN_LETTER_RE.search(text))


def _run_has_rtl_flag(r_xml: str) -> bool:
    return "<w:rtl/>" in r_xml or "<w:rtl " in r_xml


def _paragraph_pPr(p_xml: str) -> str:
    """Return the paragraph's own <w:pPr> block with any nested <w:sectPr>
    stripped, so a section's bidi/jc is never mistaken for the paragraph's.
    (jc never appears inside a sectPr, so removing it is safe for our checks.)
    """
    m = _PPR_BLOCK_RE.search(p_xml)
    if not m:
        return ""
    return _SECTPR_BLOCK_RE.sub("", m.group(0))


def _ppr_has_bidi(ppr: str) -> bool:
    return "<w:bidi/>" in ppr or "<w:bidi " in ppr


def _validate_xml(xml: str, fname: str) -> list[Issue]:
    """Run the content-level checks over one OOXML part. Pure read-only."""
    issues: list[Issue] = []

    # Rule 5 — a paragraph that is bidi but pins alignment to physical "right".
    # Word (notably Word for Mac) can re-interpret "right" as the logical end,
    # which in RTL is the LEFT side. Use logical "start" instead.
    for _, _, p_xml in _iter_blocks(xml, "w:p", "w:p"):
        ppr = _paragraph_pPr(p_xml)
        if _ppr_has_bidi(ppr) and _JC_RIGHT_RE.search(ppr):
            snippet = "".join(_W_T_RE.findall(p_xml))[:40]
            issues.append(Issue(
                "ERROR", "jc-right-in-rtl",
                'bidi paragraph uses jc="right"; use jc="start"',
                f"{fname}: {snippet!r}"))

    # Rules 7 & 8 — per-run content checks.
    for _, _, r_xml in _iter_blocks(xml, "w:r", "w:r"):
        text = _run_text(r_xml)
        if not text:
            continue
        has_rtl = _has_rtl(text)
        has_lat = _has_latin_letter(text)

        # Rule 7 — one run mixing RTL script and Latin letters reorders inside
        # the run. Each script needs its own run.
        if has_rtl and has_lat:
            issues.append(Issue(
                "ERROR", "mixed-script-run",
                f"run mixes RTL and Latin letters; split into separate runs: {text[:50]!r}",
                fname))

        # Rule 8 — Latin digits inside an RTL-flagged run. Fine for
        # IBANs/codes/URLs (which belong in their own LTR run anyway), so this
        # is a warning, not an error.
        if _run_has_rtl_flag(r_xml) and _LATIN_DIGIT_RE.search(text):
            issues.append(Issue(
                "WARN", "latin-digits-in-rtl",
                f"Latin digits in an RTL run; prefer Arabic-Indic ٠-٩ "
                f"unless this is an IBAN/code/URL: {text[:50]!r}",
                fname))

        # Rule 8 — RTL-only run ending in Latin , ; ? — should be Arabic , ; ?.
        if has_rtl and not has_lat:
            stripped = text.rstrip()
            if stripped and stripped[-1] in _TRAILING_LATIN_PUNCT:
                issues.append(Issue(
                    "WARN", "latin-punct-in-rtl",
                    f"RTL run ends with Latin punctuation; use ، ؛ ؟: {text[:50]!r}",
                    fname))

    return issues


def _fix_jc_in_xml(xml: str) -> tuple[str, int]:
    """Rewrite <w:jc w:val="right"/> -> "start" only in paragraphs that are
    unambiguously RTL: the pPr has <w:bidi/>, and every text-bearing run is
    RTL-flagged (no LTR runs). Mixed paragraphs are left untouched — rewriting
    those is a content decision, not a safe mechanical fix.
    """

    def _xform_p(p_xml: str) -> tuple[str, bool]:
        ppr = _paragraph_pPr(p_xml)
        if not _ppr_has_bidi(ppr) or not _JC_RIGHT_RE.search(ppr):
            return p_xml, False
        # Bail if any text-bearing run lacks <w:rtl/> (i.e. an LTR run).
        for _, _, r_xml in _iter_blocks(p_xml, "w:r", "w:r"):
            if _run_text(r_xml).strip() and not _run_has_rtl_flag(r_xml):
                return p_xml, False
        # jc never lives in a sectPr, so we can rewrite directly in the
        # original (un-stripped) pPr block.
        m = _PPR_BLOCK_RE.search(p_xml)
        if not m:
            return p_xml, False
        new_block, n = _JC_RIGHT_RE.subn(
            lambda mm: mm.group(0).replace('w:val="right"', 'w:val="start"'),
            m.group(0), count=1)
        if n == 0:
            return p_xml, False
        return p_xml[: m.start()] + new_block + p_xml[m.end():], True

    return _replace_blocks(xml, "w:p", "w:p", _xform_p)


def validate_docx(
    in_path: Path, out_path: Path | None = None, *, fix_jc: bool = False
) -> tuple[list[Issue], int]:
    """Validate content-level RTL rules. If fix_jc is set, first rewrite safe
    jc="right" -> "start" (writing to out_path or in place), then validate the
    result. Returns (issues, jc_fix_count).
    """
    in_path = Path(in_path)
    if not in_path.exists():
        raise FileNotFoundError(in_path)

    fix_count = 0
    scan_path = in_path

    if fix_jc:
        target = out_path or in_path
        with tempfile.TemporaryDirectory() as tmp:
            scratch = Path(tmp) / "scratch.docx"
            with zipfile.ZipFile(in_path, "r") as zin, zipfile.ZipFile(
                scratch, "w", compression=zipfile.ZIP_DEFLATED
            ) as zout:
                for info in zin.infolist():
                    data = zin.read(info.filename)
                    if _is_target_file(info.filename):
                        text, n = _fix_jc_in_xml(data.decode("utf-8"))
                        fix_count += n
                        data = text.encode("utf-8")
                    zout.writestr(info, data)
            shutil.move(str(scratch), str(target))
        scan_path = target

    issues: list[Issue] = []
    with zipfile.ZipFile(scan_path, "r") as z:
        for info in z.infolist():
            if _is_target_file(info.filename):
                text = z.read(info.filename).decode("utf-8")
                issues.extend(_validate_xml(text, os.path.basename(info.filename)))

    return issues, fix_count


# ---------- CLI ----------

def _main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx", type=Path, help="Path to .docx file to harden or validate")
    ap.add_argument("-o", "--output", type=Path, default=None, help="Output path (defaults to in-place)")
    ap.add_argument("--report", action="store_true", help="Print summary of changes")
    ap.add_argument(
        "--locale", default="ar-SA",
        help=(
            "Complex-script locale for docDefaults w:lang w:bidi. "
            f"One of: {', '.join(sorted(_BIDI_LOCALES))}. Default: ar-SA."
        ),
    )
    ap.add_argument(
        "--validate", action="store_true",
        help=(
            "Validate content-level RTL rules (5, 7, 8) instead of hardening. "
            "Reports issues and exits 1 if any ERROR is found."
        ),
    )
    ap.add_argument(
        "--fix-jc", dest="fix_jc", action="store_true",
        help=(
            'Opt-in: rewrite <w:jc w:val="right"/> -> "start" in paragraphs that '
            "are unambiguously RTL (writes the file). Implies validation."
        ),
    )
    args = ap.parse_args()

    # Validation / jc-fix mode is separate from structural hardening.
    if args.validate or args.fix_jc:
        issues, fix_count = validate_docx(args.docx, args.output, fix_jc=args.fix_jc)
        if args.fix_jc:
            print(f'--fix-jc: rewrote {fix_count} jc="right" -> "start" in RTL paragraphs.')
        errors = [i for i in issues if i.severity == "ERROR"]
        warns = [i for i in issues if i.severity == "WARN"]
        if issues:
            for issue in issues:
                print(issue)
        else:
            print("No content-level RTL issues found.")
        print(f"\n{len(errors)} error(s), {len(warns)} warning(s).")
        return 1 if errors else 0

    report = harden_docx(args.docx, args.output, locale=args.locale)

    if args.report or report.total() == 0:
        print(report.summary())
    else:
        print(f"Hardened {args.docx} ({report.total()} edits across {len(report.files_changed)} parts).")
    return 0


if __name__ == "__main__":
    sys.exit(_main())
