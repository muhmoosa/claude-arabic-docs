# Changelog — discovery history

Each rule below was added after the previous version of the skill failed to fix
a real Arabic Word document, and a human had to manually correct it in MS Word.
Diffing the manually-fixed file against the programmatically-generated one
revealed the missing piece. Each rule is documented in detail in `SKILL.md`.

## v1.0 — 2026-05-19

Initial release with six-layer hardening, after six debugging iterations on
two production legal documents (~30 pages of Arabic).

### Iteration history

| Iteration | Symptom | Root cause discovered | Fix added |
|---|---|---|---|
| 1 | Tables flow LTR (label column on the left) | docx-js doesn't add `visuallyRightToLeft` to tables by default | Set `visuallyRightToLeft: true` on every Table |
| 2 | Headings render LTR in MS Word even with paragraph bidi | Section `<w:sectPr>` lacks `<w:bidi/>` | Inject section bidi in correct schema position |
| 3 | Headings *still* LTR after section bidi | `docDefaults/rPrDefault/rPr` lacks `<w:lang w:bidi="ar-SA"/>` | Inject complex-script language declaration |
| 4 | Headings *still* LTR after lang fix | `docDefaults/pPrDefault` is empty (self-closing) | Inject `<w:pPr><w:bidi/><w:jc/></w:pPr>` into pPrDefault |
| 5 | Tables RTL but headings/body LTR despite all above | `settings.xml` lacks `<w:themeFontLang w:bidi="ar-SA"/>` — the master switch | Inject themeFontLang in schema-correct position |
| 6 | All flags present but body left-aligned in Word for Mac | Word for Mac re-interprets `<w:jc w:val="right"/>` as logical "end" (= visually left) in RTL paragraphs | Use `start`/`end` (logical) instead of `right`/`left` (physical) |

Each iteration cost ~30 minutes of bisecting plus a screenshot review with the
client. The skill's value proposition is making sure no one ever has to
re-discover these in this order again.

### Bundled scripts

- `scripts/harden_rtl.py` — applies all six XML-level fixes to a finished docx
  - Idempotent: safe to run multiple times
  - Schema-aware: respects OOXML element ordering for `sectPr`, `tblPr`, `settings`
  - Locale-flexible: `--locale ar-SA | he-IL | fa-IR | ur-PK | ar-EG | ar-AE | ar-MA`
- `scripts/arabic_numerals.py` — converts Western digits + Latin punctuation to Arabic forms
  - Protects IBAN, license codes, emails, URLs, two-letter-prefixed codes via regex heuristic
  - Handles thousands separator (`٬` U+066C), decimal separator (`٫` U+066B), percent (`٪` U+066A), comma/semicolon/question mark
