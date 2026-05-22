# Changelog — discovery history

Each rule below was added after the previous version of the skill failed to fix
a real Arabic Word document, and a human had to manually correct it in MS Word.
Diffing the manually-fixed file against the programmatically-generated one
revealed the missing piece. Each rule is documented in detail in `SKILL.md`.

## v1.1.1 — 2026-05-22

Renamed the skill's `name:` field from `claude-arabic-docs` to `arabic-rtl-docs`:
skill names may not contain the reserved word "claude", so the previous name was
rejected at install time. No functional changes — the repository, README, and the
`.skill` filename are unaffected; only the installed skill identifier changed.

## v1.1 — 2026-05-22

Follow-up after a real failure case: a generator produced a multi-page Arabic
document that passed the structural harden and looked correct in the LibreOffice
PDF preview, yet still violated three content-level rules and rendered wrong in
Word. The structural fixes can't catch those, so this release adds a validator
and a python-docx reference, and documents the python-docx alignment trap.

### Added

| Area | Change |
|---|---|
| `scripts/harden_rtl.py` | New `--validate` mode flags content-level rule violations (rules 5, 7, 8): a `bidi` paragraph using `jc="right"`, a single run mixing Arabic + Latin letters (errors), and Latin digits / trailing Latin punctuation in RTL runs (warnings). Exits 1 on errors. |
| `scripts/harden_rtl.py` | New opt-in `--fix-jc` rewrites `jc="right"` → `"start"` only in paragraphs that are unambiguously RTL (have `<w:bidi/>`, no LTR runs). Validation scans content parts only, so the deliberate `jc="right"` in document defaults is never a false positive. |
| `references/python-docx-template.md` | New battle-tested python-docx helper layer (`split_bidi` tokenizer, `add_para`/`add_run`/`make_table_rtl`) that encodes rules 1–8 by construction, plus a worked example. |
| `SKILL.md` | New "python-docx gotcha" section explaining that `WD_ALIGN_PARAGRAPH` has no `START`, so the natural API emits the buggy `jc="right"`; documents the direct-OOXML workaround. |

### Changed

- `SKILL.md` reference links repointed: the docx-js section now points to the real `examples/build_test_arabic.js`, and the per-language notes are folded inline (the previously-linked `references/docx-js-template.md` and `references/per-language.md` never existed in the repo).
- README (EN + AR): documented `--validate`/`--fix-jc` and added `references/` to "What's in the box".

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
