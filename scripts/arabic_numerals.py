#!/usr/bin/env python3
"""arabic_numerals.py — convert Western digits and Latin punctuation
in Arabic prose to Eastern Arabic-Indic digits and Arabic punctuation,
while skipping protected tokens (IBANs, license codes, emails, URLs).

Usage:
    # As a library
    from arabic_numerals import arabize
    arabize("بمبلغ 375,000 ريال (15%)")    # → "بمبلغ ٣٧٥٬٠٠٠ ريال (١٥٪)"

    # CLI: pipe in text, get arabized text out
    echo "بمبلغ 375,000 ريال" | python arabic_numerals.py

    # Skip the email/IBAN-style protected token detection
    python arabic_numerals.py --no-protect

The "protect" heuristic preserves any token matching the patterns:
  • starts with 2 capital letters (likely IBAN/country code)
  • contains '@' or '://'  (email/URL)
  • starts with a Latin prefix followed by '-' and digits (FL-888252203)
"""

from __future__ import annotations

import argparse
import re
import sys

_W2E = str.maketrans('0123456789', '٠١٢٣٤٥٦٧٨٩')
_ARABIC_COMMA = '٬'      # U+066C
_ARABIC_PERCENT = '٪'    # U+066A
_ARABIC_DECIMAL = '٫'    # U+066B

# Heuristic patterns for tokens that should NOT be transliterated.
_PROTECT_PATTERNS = [
    r'\b[A-Z]{2}\d{6,}\b',                  # IBAN-like (SA02..., GB29...)
    r'\b[A-Za-z]+-\d+\b',                   # FL-888252203, ID-1234
    r'\S+@\S+\.\S+',                        # email
    r'https?://\S+',                        # URL
    r'\bwww\.\S+',                          # URL
]
_PROTECT_RE = re.compile('|'.join(_PROTECT_PATTERNS))

# Punctuation conversion done only between Arabic-letter neighbours.
_ARABIC_LETTER_RE = re.compile(r'[؀-ۿ]')


def _convert_segment(seg: str) -> str:
    """Convert digits and Arabic-context punctuation in a single segment."""
    # 1. Grouped numbers: 2,500,000 → ٢٬٥٠٠٬٠٠٠
    def grouped(m: re.Match) -> str:
        return m.group(0).translate(_W2E).replace(',', _ARABIC_COMMA)

    seg = re.sub(r'\b\d{1,3}(?:,\d{3})+\b', grouped, seg)

    # 2. Decimal numbers: 12.5 → ١٢٫٥ (only when surrounded by digits)
    def decimal(m: re.Match) -> str:
        return m.group(0).translate(_W2E).replace('.', _ARABIC_DECIMAL)

    seg = re.sub(r'\b\d+\.\d+\b', decimal, seg)

    # 3. Bare digit runs: 15 → ١٥
    seg = re.sub(r'\d+', lambda m: m.group(0).translate(_W2E), seg)

    # 4. Percent: ١٥% → ١٥٪
    seg = re.sub(r'(?<=[٠-٩])%', _ARABIC_PERCENT, seg)
    seg = re.sub(r'%(?=[)\s])', _ARABIC_PERCENT, seg)

    return seg


def _convert_punct(text: str) -> str:
    """Replace Latin , ; ? with Arabic equivalents in Arabic-letter contexts."""
    # Replace comma with ، when adjacent to an Arabic letter on either side.
    def commareplace(m: re.Match) -> str:
        i = m.start()
        prev = text[i-1] if i > 0 else ''
        nxt = text[i+1] if i+1 < len(text) else ''
        if _ARABIC_LETTER_RE.match(prev) or _ARABIC_LETTER_RE.match(nxt):
            return '،'
        return m.group(0)

    text = re.sub(r',', commareplace, text)
    text = re.sub(r';', lambda m: '؛' if _surrounded_arabic(text, m.start()) else ';', text)
    text = re.sub(r'\?', lambda m: '؟' if _surrounded_arabic(text, m.start()) else '?', text)
    return text


def _surrounded_arabic(text: str, i: int) -> bool:
    prev = text[i-1] if i > 0 else ''
    nxt = text[i+1] if i+1 < len(text) else ''
    return bool(_ARABIC_LETTER_RE.match(prev) or _ARABIC_LETTER_RE.match(nxt))


def arabize(text: str, *, protect: bool = True, convert_punctuation: bool = True) -> str:
    """Convert digits (and optionally punctuation) to Arabic forms.

    `protect` keeps IBAN/email/code-like tokens untouched (default True).
    `convert_punctuation` also replaces , ; ? with ، ؛ ؟ in Arabic context.
    """
    if protect:
        parts: list[str] = []
        last = 0
        for m in _PROTECT_RE.finditer(text):
            parts.append(_convert_segment(text[last:m.start()]))
            parts.append(text[m.start():m.end()])  # preserve as-is
            last = m.end()
        parts.append(_convert_segment(text[last:]))
        out = ''.join(parts)
    else:
        out = _convert_segment(text)

    if convert_punctuation:
        out = _convert_punct(out)

    return out


def _main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--no-protect", action="store_true", help="Disable IBAN/email/code-token protection")
    ap.add_argument("--no-punct", action="store_true", help="Don't replace , ; ? with Arabic equivalents")
    args = ap.parse_args()

    text = sys.stdin.read()
    out = arabize(text, protect=not args.no_protect, convert_punctuation=not args.no_punct)
    sys.stdout.write(out)
    return 0


if __name__ == "__main__":
    sys.exit(_main())
