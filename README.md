# claude-arabic-docs

> A Claude skill that fixes a class of bugs no one tells you about until you ship a 30-page Arabic Word document and the client opens it on Mac. Read the Arabic version: [README.ar.md](./README.ar.md)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![Claude Skill](https://img.shields.io/badge/Claude-Skill-blueviolet)](https://docs.claude.com/skills)
[![Languages: Arabic · Hebrew · Persian · Urdu](https://img.shields.io/badge/RTL-AR%20·%20HE%20·%20FA%20·%20UR-green)](#supported-locales)

## The bug this skill exists to fix

You generate an Arabic `.docx` programmatically using `docx-js`, `python-docx`, or any of the popular libraries. Every paragraph is marked RTL, every run carries `<w:rtl/>`, every table has `<w:bidiVisual/>`. You convert to PDF in LibreOffice and it looks perfect. You hand the file to a real user.

**They open it in Microsoft Word and every heading and body paragraph is left-aligned. Tables are correct. Only the body is broken.**

You google this. You add `<w:bidi/>` to every paragraph. Still broken in Word. You add it to the section. Still broken. You set `docDefaults` language to `ar-SA`. Still broken. You add `<w:pPrDefault>` with bidi. Still broken.

After two days of bisecting OOXML XML, you discover that Microsoft Word ignores all of those if `settings.xml` doesn't have `<w:themeFontLang w:bidi="ar-SA"/>`. That setting is never written by any docx generation library — it's only populated by a live Word session, sourced from the OS keyboard layout. So *every* programmatically-generated docx starts off broken in Word.

You fix that and the document *almost* renders correctly — except now headings are still left-aligned. After more bisecting you discover that Word for Mac re-interprets `<w:jc w:val="right"/>` as logical "end" (= left in RTL mode), not physical right. You change everything to `start`/`end` and it finally works.

This skill packages all six of those fixes — discovered the hard way over many hours — into a description that triggers on RTL documents, a hardening script you can run as a post-processor, and a digit/punctuation converter for Arabic typography conventions.

## What it does

When Claude is generating a `.docx` (or a PDF) that contains Arabic, Hebrew, Persian, Urdu, or any RTL script, this skill makes Claude:

1. Use logical (`start`/`end`) alignment instead of physical (`left`/`right`) in generation.
2. Inject the **six RTL layers** Word actually checks (full table below).
3. Convert digits to Eastern Arabic-Indic (٠–٩) and Latin commas/percents to Arabic equivalents (٬ ٪ ، ؛ ؟), while keeping IBAN/email/license codes Latin via heuristics.

## Installation

Every release on the [Releases page](https://github.com/muhmoosa/claude-arabic-docs/releases/latest) has a ready-made `arabic-rtl-docs.skill`. This link always gives you the newest version:

**[⬇ Download arabic-rtl-docs.skill](https://github.com/muhmoosa/claude-arabic-docs/releases/latest/download/arabic-rtl-docs.skill)**

### Claude app (web, desktop and Cowork)

1. Download the `.skill` file from the link above.
2. In Claude, open the Skills settings and upload the file. If an older version is installed, remove it first so only one copy is active.

### Claude Code (CLI)

```bash
curl -fL -o /tmp/arabic-rtl-docs.skill https://github.com/muhmoosa/claude-arabic-docs/releases/latest/download/arabic-rtl-docs.skill
unzip -o /tmp/arabic-rtl-docs.skill -d ~/.claude/skills/   # installs to ~/.claude/skills/arabic-rtl-docs
```

To update, run the same two commands again; `unzip -o` replaces the old files.

## Usage

Once installed, the skill triggers automatically whenever Claude detects an Arabic / RTL request. No manual invocation needed. If you want to be explicit:

> "Make me an Arabic Word document about X, and use the arabic-rtl-docs skill."

If you're using the bundled scripts manually:

```bash
# Harden any .docx (in place, or into a second file)
python3 scripts/harden_rtl.py document.docx
python3 scripts/harden_rtl.py input.docx output.docx

# Check only: lists what would change, exits 1 unless clean. "warn:" lines are content issues
# (mixed Arabic/Latin runs, Latin punctuation, Windows-only fonts) and never fail the check.
python3 scripts/harden_rtl.py document.docx --check

# Another RTL locale
python3 scripts/harden_rtl.py document.docx --lang he-IL   # Hebrew
python3 scripts/harden_rtl.py document.docx --lang fa-IR   # Persian

# PDF from the hardened file (fonts must be installed: fc-list :lang=ar family)
soffice --headless --convert-to pdf document.docx

# Convert digits and punctuation in Arabic text (stdin → stdout)
echo "بمبلغ 375,000 ريال (15%)" | python3 scripts/arabic_numerals.py
# → "بمبلغ ٣٧٥٬٠٠٠ ريال (١٥٪)"
```

The hardening script is **idempotent** — running it twice is a no-op on the second pass.

## The six layers, briefly

| # | Layer | XML | Why |
|---|---|---|---|
| 0.0 | `settings.xml` themeFontLang | `<w:themeFontLang w:bidi="ar-SA"/>` | **Master switch.** Without it, Word doesn't engage its RTL pipeline at all. |
| 0   | docDefaults rPr lang        | `<w:lang w:bidi="ar-SA"/>`         | Tells Word the complex-script language. Tables work without it; nothing else does. |
| 0.5 | docDefaults pPrDefault       | `<w:pPr><w:bidi/><w:jc w:val="start"/></w:pPr>` | Default direction for the `Normal` paragraph style. Without it, headings/body fall back to LTR. |
| 1   | Section bidi                 | `<w:bidi/>` in `<w:sectPr>`        | Section-level reading direction. Schema-position-sensitive (must come before `<w:docGrid>`). |
| 2   | Table bidi-visual            | `<w:bidiVisual/>` in `<w:tblPr>`   | Flips column order so first cell renders on the right. |
| 3   | Paragraph + run flags        | `<w:bidi/>` + `<w:rtl/>`           | Element-level RTL. Cheap, do this from the start. |
| 5   | Logical alignment            | Prefer `w:jc="start"`/`"end"`      | Word for Mac re-interprets physical `left`/`right` as logical in RTL mode. Counter-intuitive — use start/end. |

Full rules with rationale: see [`SKILL.md`](./SKILL.md).

## Supported locales

- `ar-SA` Arabic (Saudi Arabia) — **default**
- `ar-EG`, `ar-AE`, `ar-MA` Arabic (Egypt, UAE, Morocco)
- `he-IL` Hebrew (Israel)
- `fa-IR` Persian (Iran)
- `ur-PK` Urdu (Pakistan)

Other RTL scripts (Syriac, Thaana, NKo, Mandaic, Samaritan) are detected by Unicode range when injecting paragraph/run RTL flags, but the `themeFontLang` master switch needs an explicit locale code Word accepts. PRs to extend the list welcome.

## What's in the box

```
claude-arabic-docs/
├── SKILL.md                       The skill: rules, PDF and verification (read this)
├── README.md                      You are here
├── README.ar.md                   Arabic version of this README
├── LICENSE                        MIT
├── CHANGELOG.md                   Discovery history of each rule
├── scripts/
│   ├── harden_rtl.py              Fix or --check any .docx (structure, complex-script props, schema order)
│   └── arabic_numerals.py         Convert Western digits + Latin punct to Arabic forms
├── references/
│   └── python-docx-template.md    Known-good python-docx helper layer (rules 1–8 by construction)
└── examples/
    ├── build_test_arabic.js       Minimal docx-js example using the skill conventions
    └── sample-output.docx         The generated test document
```

## Building from source

```bash
git clone https://github.com/muhmoosa/claude-arabic-docs
cd claude-arabic-docs

# Pack the .skill file: a plain ZIP holding an arabic-rtl-docs/ folder with SKILL.md, scripts and references
mkdir -p build/arabic-rtl-docs && cp -R SKILL.md scripts references build/arabic-rtl-docs/
(cd build && zip -qr ../arabic-rtl-docs.skill arabic-rtl-docs)

# Test that the hardening script runs
python3 scripts/harden_rtl.py examples/sample-output.docx --check
```

## Testing

After making changes to the hardening script, regenerate the example and confirm the report:

```bash
node examples/build_test_arabic.js
python3 scripts/harden_rtl.py examples/sample-output.docx --check
# Second run with --check must print "clean" (the script is idempotent)
```

## Acknowledgements

This skill exists because Mohammed Almousa (the human who commissioned it) spent enough hours fixing the same Word RTL bug across enough deliverables that codifying it became cheaper than re-debugging it.

The discovery process is documented in [`CHANGELOG.md`](./CHANGELOG.md) — six layers, six iterations, each found by diffing a Word-saved file against the programmatically-generated original.

## License

MIT — see [`LICENSE`](./LICENSE). Use it, fork it, improve it. If you do improve it, please open a PR back here so the next person doesn't have to rediscover what you fixed.
