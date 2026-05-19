---
name: claude-arabic-docs
description: Produce correctly-rendered right-to-left (RTL) Microsoft Office documents — Word (.docx), PowerPoint (.pptx), and Excel (.xlsx) — for Arabic, Hebrew, Persian, Urdu, and other RTL scripts. Use this skill whenever the user asks for a document in Arabic or another RTL language, or whenever the deliverable mixes RTL prose with English/Latin tokens (IBANs, URLs, emails, numbers). It must also be used whenever you produce a docx-js / openpyxl / python-pptx output that contains any RTL text, because those libraries do NOT apply section-level bidi, table direction, or heading bidi automatically — outputs that look correct in LibreOffice can still render as LTR (cells reversed, headings left-aligned, lists flipped) when opened in Microsoft Word. Trigger on phrases like "اكتب", "خطاب", "تقرير عربي", "Arabic letter", "RTL report", "بالعربي", "اللغة العربية", "Hebrew document", "Persian", or any user message written predominantly in an RTL script.
---

# Arabic / RTL Documents

## Why this skill exists

RTL bugs in Office files are sneaky. They almost never throw an error and they often look fine in LibreOffice or in a PDF preview — but a user opening the file in Microsoft Word sees:

- Headings left-aligned even though the body looks right.
- Table cells reversed (the *label* column ends up on the LEFT when it should be on the RIGHT).
- Bullet lists flipped (bullet on the right, indent going the wrong way).
- IBANs and emails embedded in Arabic prose getting their digits reordered.
- Punctuation drifting to the wrong side of a word.

The root cause is almost always the same: the generator emitted *paragraph-level* RTL flags but skipped the *section-level* and *table-level* flags. Word treats those as authoritative when laying out headings, footers, table direction, and auto-numbering. LibreOffice is more forgiving, so problems hide in QA.

This skill bundles a checklist plus a post-build hardening script that fixes the common omissions in a finished `.docx`.

## When to apply this skill

Apply it whenever **any** of the following is true:

1. The document contains Arabic, Hebrew, Persian, Urdu, Pashto, Dari, Sindhi, Kashmiri, or any other RTL script.
2. The document is described as an Arabic letter, memo, report, contract, claim, خطاب, تقرير, عقد, مذكرة, محضر, إفادة, etc.
3. The document mixes RTL prose with LTR tokens (IBANs, URLs, codes, English names).
4. You are about to build a `.docx`/`.pptx`/`.xlsx` with `docx-js`, `python-docx`, `openpyxl`, or `python-pptx` and any text run is RTL.

Even if the user did not say "make sure it's RTL," do this. Users expect Arabic documents to *look right in Word*; "I made it RTL in the runs" is not enough.

## The non-negotiable RTL rules

These are the things that *must* be set. Each is something Microsoft Word and Office 365 actively look at, and each is something the popular libraries forget to set by default.

### 0.0. `settings.xml` themeFontLang (the actual master switch)

**This is the single setting MS Word actually checks** to decide whether the document gets the RTL rendering pipeline. Every other RTL setting below is *necessary* but not *sufficient* on its own.

`word/settings.xml` must contain:

```xml
<w:themeFontLang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>
```

(or with `he-IL`, `fa-IR`, etc. depending on the script.)

What this controls in Word:

- Whether Word engages its RTL layout engine at all (the master switch).
- Whether headings and body paragraphs respect bidi at all.
- Whether `<w:jc w:val="left"/>` and `<w:jc w:val="right"/>` get treated as physical or logical alignment in RTL flow.

Without this, the layout you see is *exactly* the symptom:

> **Tables render correctly RTL, but every single heading and body paragraph renders left-aligned in Microsoft Word — even when the document has all of: `<w:bidi/>` on every paragraph, `<w:rtl/>` on every run, `<w:bidiVisual/>` on every table, `<w:bidi/>` on the section, `<w:lang w:bidi="ar-SA"/>` in `docDefaults`, AND `<w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>` in `pPrDefault`.**

That last sentence describes a file that *looks correct in LibreOffice's PDF export, looks correct in pandoc's HTML conversion, passes Word's XML validator, and still renders left-aligned the moment a real Word user opens it.* The reason is that the other settings all describe *content* characteristics ("this paragraph is RTL", "this run is RTL"). `themeFontLang` describes the *document's identity* — and Word's RTL pipeline activates only when the document is identified as RTL.

`docx-js`, `python-docx`, and the openpyxl-derived converters generate `settings.xml` *without* `themeFontLang` — that element only gets written when a real Word session saves the file, populated from the OS keyboard layout. So any document generated programmatically starts off without it.

The bundled hardening script injects this setting. Run it.

### 0. Document-default complex-script language (THE critical one)

This is the rule that explains the majority of "I made everything RTL and it still renders left-aligned in Word" bug reports.

`styles.xml` contains a `<w:docDefaults><w:rPrDefault><w:rPr>…</w:rPr></w:rPrDefault></w:docDefaults>` block. That `<w:rPr>` must include a `<w:lang>` element with a `w:bidi="…"` attribute set to an RTL locale:

```xml
<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>
```

Without this declaration, Word treats the document as a Western (LTR) document that *happens to contain* some Arabic glyphs, and it applies LTR layout rules everywhere — paragraph alignment falls back to LEFT, headings render LTR, list bullets land on the wrong side. None of the paragraph-level `<w:bidi/>` or run-level `<w:rtl/>` flags rescue this. Word needs to know, at the *document defaults* level, that the complex-script language is Arabic (or Hebrew, Persian, Urdu, etc.).

LibreOffice doesn't need this — it infers RTL from the runs. Which is why the bug is invisible until someone opens the file in Word.

`docx-js`, `python-docx`, and `openpyxl→docx` converters all default to an empty `<w:rPr>` in `docDefaults` and never add the language declaration. The bundled hardening script (`scripts/harden_rtl.py`) injects it. If you're writing your own generator, do this once when you build `styles.xml` and you're done.

Locale codes Word accepts: `ar-SA` (Saudi), `ar-EG`, `ar-AE`, `ar-MA`, `he-IL`, `fa-IR`, `ur-PK`. The skill defaults to `ar-SA`.

### 0.5. Document-default paragraph properties AND Normal-style bidi

The companion to Rule #0. `styles.xml` also contains `<w:docDefaults><w:pPrDefault>…</w:pPrDefault></w:docDefaults>` — the default *paragraph* properties for the whole document. If this is empty or self-closing (`<w:pPrDefault/>`), Word treats the document's *default* paragraph reading direction as LTR. The result is the most confusing RTL bug of all:

> **Tables render correctly RTL, but headings and body paragraphs render left-aligned in Microsoft Word — even though every single `<w:p>` in the file has `<w:bidi/>` and `<w:jc w:val="right"/>` set explicitly.**

The symptom is unique because tables go through `<w:bidiVisual/>` (which is unaffected), but headings/body inherit their *default* direction from `pPrDefault`. When that's empty, MS Word's renderer applies the LTR fallback to anything that uses the Normal style or any style derived from it — which in practice is everything that isn't inside a table cell with its own bidi.

The fix:

```xml
<w:pPrDefault>
  <w:pPr>
    <w:bidi/>
    <w:jc w:val="right"/>
  </w:pPr>
</w:pPrDefault>
```

Apply the same to the `Normal` style if it's explicitly defined in the document (`<w:style w:type="paragraph" w:styleId="Normal">`). The `Normal` style is the parent of every paragraph style by default; if it lacks bidi, derived styles can silently fall back to LTR.

LibreOffice ignores this entire class of bug — it infers RTL from the runs, so PDF previews look correct. The bug is invisible until someone opens the file in Word. **Always run the bundled hardening script, which patches both `pPrDefault` and the `Normal` style.**

### 1. Section-level bidi

Every `<w:sectPr>` must contain `<w:bidi/>`. This sets the default reading direction for the whole section. Without it, Word renders headings, footers, page numbers, and auto-numbered lists left-to-right even when individual paragraphs are bidi.

`docx-js` does **not** expose this directly — you have to post-process the XML, or use the hardening script bundled with this skill (`scripts/harden_rtl.py`).

### 2. Paragraph-level bidi on every Arabic paragraph

Every `<w:p>` containing RTL text must have `<w:bidi/>` inside its `<w:pPr>`. In `docx-js`, this is the `bidirectional: true` option on `Paragraph`. In `python-docx`, set `paragraph.paragraph_format.element.get_or_add_pPr().append(OxmlElement('w:bidi'))`.

### 3. Run-level RTL on every Arabic run

Every `<w:r>` containing RTL text must have `<w:rtl/>` inside its `<w:rPr>`. In `docx-js`, that's `rightToLeft: true` on `TextRun`. Without this, mixed Arabic+number runs ("بمبلغ 375,000 ريال") will reorder incorrectly.

### 4. Table direction: `visuallyRightToLeft` (a.k.a. `<w:bidiVisual/>`)

Every `<w:tbl>` containing RTL content must have `<w:bidiVisual/>` inside its `<w:tblPr>`. This flips the visual order of columns so that the **first cell** in your code renders on the **right**.

Without this, a 2-column "label | value" table renders with the label on the LEFT — exactly the opposite of what an Arabic reader expects. This is the single most common RTL bug.

In `docx-js`, pass `visuallyRightToLeft: true` to `new Table({...})`. The hardening script also injects this if it's missing.

### 5. Use logical `start`/`end` alignment, NOT physical `left`/`right`

**The rule:** for any RTL body paragraph, use `AlignmentType.START` (`<w:jc w:val="start"/>`), NOT `AlignmentType.RIGHT`. For LTR content (English contact lines, IBANs on their own line), use `AlignmentType.LEFT` is fine because those paragraphs aren't bidi.

**Why this matters in MS Word — especially Word for Mac.** OOXML defines two alignment families:

- **Physical**: `left`, `right`, `center`, `both` — fixed visual side regardless of text direction.
- **Logical**: `start`, `end` — direction-aware. In LTR, `start = left`; in RTL, `start = right`.

The OOXML spec says `left`/`right` should be physical. **Microsoft Word does not consistently follow that.** In a paragraph that resolves to RTL direction (whether via explicit `<w:bidi/>` or inherited from `pPrDefault`), Word — particularly Word for Mac — often re-interprets `w:jc="right"` as the *logical end*, which in RTL is the LEFT side. The exact symptom is:

> Your XML has `<w:jc w:val="right"/>` on every body paragraph. LibreOffice's PDF preview shows it correctly right-aligned. The user opens it in MS Word and every paragraph is left-aligned. They manually fix it. You inspect their saved file and the alignment value is now `<w:jc w:val="left"/>` — yet the document renders right-aligned in Word. Word flipped your "right" to its left-rendered, and the user's "left" to its right-rendered.

That non-intuitive flip is real, undocumented, and Word-version-specific. The reliable fix is to use `start` and `end` everywhere, which are direction-aware and unambiguously rendered: `start` always means "leading edge of the line" (right in RTL, left in LTR), `end` always means "trailing edge."

**Practical mapping:**

| You want… | Use in docx-js | OOXML output |
|---|---|---|
| RTL body, right-aligned (visually) | `AlignmentType.START` | `<w:jc w:val="start"/>` |
| LTR body, left-aligned (visually) | `AlignmentType.START` | `<w:jc w:val="start"/>` |
| Centered (titles, invocations) | `AlignmentType.CENTER` | `<w:jc w:val="center"/>` |
| Force LTR (e.g. an Arabic doc's English contact line) | `AlignmentType.LEFT` *plus* set the paragraph and run to LTR (`bidirectional: false`, `rightToLeft: false`) | `<w:jc w:val="left"/>` (no `<w:bidi/>`, no `<w:rtl/>`) |
| Force right side regardless of direction | `AlignmentType.RIGHT` — only when you really mean physical right | `<w:jc w:val="right"/>` |

The bundled hardening script does **not** rewrite `right` to `start` automatically — that's a content choice, not a layout-bug fix. But if you're generating a new document, default every RTL paragraph to `START`.

### 6. Heading styles must carry bidi

If you override the built-in heading styles (`Heading1`, `Heading2`, …) for an Arabic document, you must include `bidi: true` in the paragraph properties of the style itself, not just on instances. Otherwise the style's defaults overrule per-paragraph settings in some Word versions.

In `docx-js`:

```js
{ id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal",
  run: { size: 32, bold: true, font: "Arial", rightToLeft: true },
  paragraph: { spacing: { before: 240, after: 240 }, alignment: AlignmentType.RIGHT,
               bidirectional: true, outlineLevel: 0 } }
```

### 7. LTR tokens inside RTL paragraphs

For an IBAN, email, URL, or English name embedded in an Arabic paragraph, *split it into its own run* and leave that run's `rightToLeft` flag **off**. Otherwise the digit order can flip. For an entire cell that's LTR (e.g., a cell that contains only an IBAN), set the paragraph's `bidirectional: false` and the run's `rightToLeft: false`, and use `AlignmentType.LEFT`.

### 8. Numbers and punctuation

**Punctuation.** In Arabic prose, use Arabic punctuation, not Latin:

| Use | Codepoint | Don't use |
|---|---|---|
| `،` comma | U+060C | `,` |
| `؛` semicolon | U+061B | `;` |
| `؟` question mark | U+061F | `?` |
| `«…»` quotation | U+00AB / U+00BB | `"…"` |
| `٪` percent | U+066A | `%` |
| `٬` thousands separator | U+066C | `,` |
| `٫` decimal separator | U+066B | `.` |
| `–` en-dash for ranges | U+2013 | `-` |

**Digits — default to Eastern Arabic-Indic.** Use ٠١٢٣٤٥٦٧٨٩ in Arabic prose, dates, currency, percentages, and clause numbers. Examples:

- Money: `(٣٧٥٬٠٠٠) ريال سعودي`
- Percentage: `١٥٪`
- Date: `١٩/٠٥/٢٠٢٦م`
- Clause reference: `البند الرابع/فقرة (١)`

**Exceptions that stay in Latin digits** even inside Arabic prose:

- **IBAN** numbers (`SA0280000301608016014488`) — international standard, must stay Latin.
- **License / registration codes** with a Latin prefix (`FL-888252203`).
- **Email addresses, URLs, software identifiers.**
- **Western brand or product codes** the user explicitly wants preserved.

When mixing, put the Latin token in its own run with `rightToLeft: false` so digit order doesn't reorder.

Some Gulf documents use Western digits throughout for executive/financial reports — confirm with the user if the deliverable is for a mixed Arabic/English audience. The default for traditional Saudi business writing is Eastern Arabic-Indic.

### 9. Fonts

Default to **Arial** for portability — it ships with every Office install and renders Arabic correctly. Other safe choices: **Calibri**, **Tahoma**, **Sakkal Majalla** (Word default for Arabic), **Traditional Arabic** for body, **Amiri** for literary work. Avoid fonts with no `Arabic` block (e.g., Times New Roman renders Arabic, but loses some shaping nuance).

Always specify the font on the run (`font: "Arial"`) — relying on document defaults is unreliable across Office versions.

### 10. Dates

Saudi/Gulf business documents conventionally use dual dating: `04 / 01 / 1447 هـ  الموافق  29 / 06 / 2025 م`. Always include the era marker (`هـ` for Hijri, `م` for Gregorian).

## The recommended build flow

1. **Build the document normally** with `docx-js` (or the lib of your choice). Apply rules 2–7 from the start — they're cheap and prevent the worst problems.
2. **Save the file.**
3. **Run the hardening script** (`scripts/harden_rtl.py`) to catch what the library missed:
   - Injects `<w:bidi/>` into every `<w:sectPr>`.
   - Injects `<w:bidiVisual/>` into every `<w:tblPr>`.
   - Injects `<w:bidi/>` into every paragraph that contains RTL text but lacks the flag.
   - Injects `<w:rtl/>` into every run that contains RTL text but lacks the flag.
   - Leaves explicitly-LTR paragraphs alone (those with `<w:jc w:val="left"/>` *and* no RTL text in their runs).
4. **Convert to PDF and visually verify** the *first time* you produce a given layout; once the template is known good, you can skip the visual check.

## Reference: hardening script

The bundled `scripts/harden_rtl.py` takes a `.docx` and produces a hardened copy in-place (or to a new path):

```bash
python scripts/harden_rtl.py path/to/document.docx               # in place
python scripts/harden_rtl.py path/to/in.docx -o path/to/out.docx # to new file
python scripts/harden_rtl.py path/to/doc.docx --report           # print what it changed
```

Read the script's docstring for the exact list of transforms.

## Reference: docx-js snippets

A complete docx-js template for an RTL Arabic letter is in `references/docx-js-template.md`. Read it when you're about to generate an Arabic `.docx` and want a known-good starting point.

## Reference: language-specific notes

`references/per-language.md` has notes on Arabic (Saudi business conventions, common honorifics, dual dating), Hebrew (calendar formatting, niqqud), Persian (Farsi vs Dari, Persian digits), and Urdu (Nastaliq fonts, Pakistani date conventions). Read it when working in any of those languages.

## A few things that look like bugs but aren't

- **Numbers inside Arabic prose look "backwards" in plain text editors.** That's the editor, not the file. Word and PDF render correctly.
- **Mixed LTR/RTL paragraph alignment can look ragged in LibreOffice but clean in Word.** Don't chase rendering differences between viewers — verify in Word if the deliverable is for Word users.
- **Heading numbering (1, 2, 3) appears on the wrong side after applying RTL.** That's correct — in RTL, the number naturally appears on the right of the heading text. If the user wants Latin numbering on the left, that's a design override, not a bug.
