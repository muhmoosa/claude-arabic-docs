---
name: arabic-rtl-docs
description: Produce and fix right-to-left (RTL) Word documents (.docx) and PDFs in Arabic, Hebrew, Persian, Urdu and other RTL scripts so they render correctly in Microsoft Word on Windows and Mac, Google Docs, and LibreOffice on Linux. Use whenever a document is in an RTL language or mixes RTL prose with Latin tokens (IBANs, URLs, emails, numbers), whenever docx-js, python-docx or pandoc output or a PDF contains RTL text, and whenever an RTL document shows left-aligned headings or paragraphs, reversed table columns, flipped lists, wrong font, lost bold or wrong size in Word, Google Docs or a PDF. Trigger on "اكتب", "خطاب", "تقرير عربي", "بالعربي", "Arabic letter", "RTL report", "Arabic PDF", "Hebrew document", "Persian", or a message written mostly in an RTL script.
---

# Arabic / RTL documents (.docx and PDF)

## Why this needs care

RTL faults almost never raise an error, and the apps people open the file in disagree:

| Reader | What goes wrong |
|---|---|
| **LibreOffice** (Linux, and PDFs made with it) | Infers direction from the text and tolerates broken markup, so **its preview proves nothing about Word or Google Docs**. Missing fonts are silently replaced (often by DejaVu Sans). |
| **Word, Windows and Mac** | Needs RTL marked in settings and defaults. **Ignores property elements out of schema order.** Styles RTL text from the complex-script slots (`bCs`, `iCs`, `szCs`, `rFonts w:cs`), so `bold`/`size`/`font` set only on the Latin slots do nothing to Arabic. Mac Word may show `jc="right"` on the left. |
| **Google Docs** | **Ignores `jc="start"`** and uses the physical alignment inherited from styles and `pPrDefault`; an inherited `right` shows on the **left**. Substitutes fonts it does not host (Traditional Arabic, Sakkal Majalla…). |

So: mark RTL at every level, use logical alignment through the whole style chain, keep schema order, fill the complex-script properties, pick fonts the reader has, and verify in the reader's app.

Apply this skill for any RTL script (Arabic, Hebrew, Persian, Urdu, Pashto, Dari, Sindhi, Kashmiri…), any Arabic letter, memo, report, contract or record (خطاب، تقرير، عقد، مذكرة، محضر، إفادة), and any `.docx` or PDF with an RTL run, even if the user never said "RTL".

## Workflow

1. Build the `.docx` following the rules below (or take the file you were asked to fix).
2. Harden it with the bundled script, then check it. `--check` must print `clean`; read every `warn:` line and fix the content it points to.
   ```bash
   python3 <this-skill-dir>/scripts/harden_rtl.py in.docx out.docx   # omit out.docx to fix in place; --lang he-IL etc.
   python3 <this-skill-dir>/scripts/harden_rtl.py out.docx --check
   ```
3. If a PDF is wanted, make it from the hardened file (see PDF).
4. Verify in the reader's app (see Verification).

The script handles all the structural rules below, so do not hand-edit XML for them. It is idempotent and needs `lxml`. It does **not** rewrite text: rules 6–8 are yours.

## Structural rules (the script enforces these)

1. **Document is RTL.** `settings.xml` `<w:themeFontLang … w:bidi="ar-SA"/>`; `docDefaults` `<w:lang … w:bidi="ar-SA"/>`. Locales: `ar-SA` (default), `ar-EG`, `ar-AE`, `ar-MA`, `he-IL`, `fa-IR`, `fa-AF`, `ur-PK`.
2. **Defaults and styles use logical alignment.** `pPrDefault` is `<w:bidi/><w:jc w:val="start"/>`, **never `right`** (that is what makes Google Docs show everything on the left). `Normal` carries `bidi`; no paragraph style carries `jc="right"`/`"left"` (including headings and pandoc's `BodyText`, `FirstParagraph`, `Compact`); use `start`, `end`, `center`. Mac Word in particular may show `right` on the left.
3. **Every level is marked.** RTL paragraph: `<w:bidi/>`. RTL run: `<w:rtl/>` (without it "بمبلغ 375,000 ريال" reorders). RTL table: `<w:bidiVisual/>` (first cell in code is on the right). Every `sectPr` and every numbering level: `<w:bidi/>`. A paragraph with Latin text and no RTL text gets `<w:bidi w:val="0"/>`, otherwise it inherits RTL from the defaults.
4. **Complex-script properties are filled.** Every `b`, `i`, `sz` has its `bCs`, `iCs`, `szCs`, and `rFonts` has `w:cs`. python-docx's `run.bold`, `font.size` and `font.name` write only the Latin slots, so in Word the Arabic shows unbolded, at the default size, in the fallback font.
5. **Schema order.** Word drops property children written out of order, without warning (`pPr`: `… numPr … bidi … spacing ind … jc … rPr`; `rPr`: `rFonts b bCs i iCs … sz szCs … rtl … lang`; `sectPr`: `… bidi rtlGutter docGrid`; `tblPr`: `tblStyle … bidiVisual … tblW jc …`). Appending in any order is fine because the script re-sorts.

When writing with a library anyway: docx-js `bidirectional: true`, `rightToLeft: true`, `visuallyRightToLeft: true`, `alignment: AlignmentType.START`, and the same on heading styles. python-docx has no API for any of this and its `WD_ALIGN_PARAGRAPH.RIGHT` writes the bad `jc="right"`, so leave alignment unset and let the script add the rest.

## Content rules (yours; the script only warns)

6. **Latin tokens get their own run.** An IBAN, email, URL, code or English name goes in a separate run without `rtl`. Keep brackets, quotes and punctuation **in the RTL run** next to the token, not in the Latin run, or Word mirrors or strands them: `(` `SA0280000301608016014488` `)` is three runs, and a numeric range such as `10–19` stays in one run. Do not use bidi control characters (U+2066–2069, U+202A–202E): they show as boxes in many Word fonts.
7. **Arabic punctuation and digits.** Use `،` U+060C, `؛` U+061B, `؟` U+061F, `«…»`, `٪` U+066A, `٬` U+066C (thousands), `٫` U+066B (decimal) and `–` U+2013 for ranges, never Latin `,` `;` `?` in Arabic prose. Saudi/Gulf business writing defaults to Arabic-Indic digits ٠–٩ (`(٣٧٥٬٠٠٠) ريال`, `١٥٪`, `١٩/٠٥/٢٠٢٦م`); keep Latin digits for IBANs, Latin-prefixed codes (`FL-888252203`), emails and URLs. For mixed Arabic/English reports, ask. Dual dating: `٠٤/٠١/١٤٤٧هـ الموافق ٢٩/٠٦/٢٠٢٥م`. Common honorifics: `سعادة`، `المكرم`.
8. **Fonts the reader has.** Set the font on every run, and put the **same face in `ascii`, `hAnsi` and `cs`**: Google Docs uses only the `ascii`/`hAnsi` face, so `ascii=Arial, cs=Amiri` shows Arial there (tested), while Word shows Amiri.

   | Font | Word Windows | Word Mac | Google Docs | Linux / LibreOffice |
   |---|---|---|---|---|
   | Arial, Tahoma | yes | yes | yes | only if installed |
   | Traditional Arabic, Simplified Arabic, Sakkal Majalla | yes | may be substituted | substituted | substituted |
   | Google fonts: Amiri, Noto Naskh Arabic, Noto Sans Arabic, IBM Plex Sans Arabic, Cairo, Tajawal | only if installed | only if installed | yes (tested) | install (`fonts-hosny-amiri`, `fonts-noto-core`) |

   Default to Arial: it is everywhere. A Google font looks better but Word shows it only where it is installed, so use one when the reader is on Google Docs, when the file goes out as a PDF, or when the user accepts that Word readers without the font see a substitute. Traditional Arabic suits Windows-only readers who expect it. A font the user asked for wins; when it will be substituted, say so.

   Other scripts: Hebrew (`he-IL`) uses Western digits and keeps niqqud in the same run as its consonants; Hebrew-calendar dates are usually written out. Persian (`fa-IR`) uses ۰–۹ (U+06F0–06F9), which differ from Arabic-Indic for 4, 5 and 6; Dari (`fa-AF`) uses Arabic-Indic. Urdu (`ur-PK`) needs a Nastaliq font (Jameel Noori Nastaleeq, Noto Nastaliq Urdu; Arial flattens it) and `dd/mm/yyyy` Gregorian dates.

**Helpers** (open only when useful): `scripts/arabic_numerals.py` converts digits and punctuation in Arabic text while leaving IBANs, codes, emails and URLs alone (`from arabic_numerals import arabize`, or pipe text through it). `references/python-docx-template.md` has python-docx helpers (`split_bidi`, `add_para`, `make_table_rtl`) that split mixed text into correct runs and apply rule 6 by construction.

## PDF

- **Best:** export from the reader's own app (Word: File › Save As › PDF; Google Docs: File › Download › PDF) so its fonts and layout are used.
- **Headless:** `soffice --headless --convert-to pdf --outdir out/ out.docx` on the **hardened** file. LibreOffice gets direction right even on a broken file, so a good PDF never proves the `.docx` is good.
- **Fonts first.** Check `fc-list :lang=ar family`. If the document's fonts are missing, install them or switch fonts before converting, then confirm with `pdffonts out.pdf`: seeing DejaVu Sans means a substitution happened.
- **Directly to PDF, without Word:** use HTML through Chromium (Playwright `page.pdf()`) or WeasyPrint with `<html lang="ar" dir="rtl">` and a real Arabic font in CSS. ReportLab and similar canvas libraries do not shape Arabic: letters come out disconnected and in reverse order.

## Verification

- `--check` prints `clean`. That covers the Word schema trap, the complex-script trap and the Google Docs alignment trap, none of which a preview shows.
- **Google Docs**, when that is the reader and a Drive connector is available: upload a copy without conversion, open it, and check headings, body, one table and one list; trash the copy afterwards. For files over the upload limit, swap the images in `word/media/` for tiny placeholders first.
- **Word**: you usually cannot open it. Ask the user to check heading and body alignment, table column order, bullets, and that bold and size apply to the Arabic. Say plainly that a LibreOffice preview is not evidence.
- **PDF**: render a page to an image and look at it; check `pdffonts`.

## Not bugs

- Numbers in Arabic prose look reversed in plain-text editors, and `pdftotext` output of Arabic is often reordered or splits ligatures (`الاسم` → `االسم`). The rendered page is what matters.
- Heading numbers sit on the right of RTL headings. Correct.
- LibreOffice and Word wrap mixed-direction lines slightly differently. Judge in the reader's app.
