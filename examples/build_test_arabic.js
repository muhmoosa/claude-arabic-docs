// build_test_arabic.js — minimal Arabic lorem-ipsum test document
// to demonstrate the claude-arabic-docs skill end-to-end:
//   1. Generate with docx-js using AlignmentType.START (skill Rule #5)
//   2. Apply harden_rtl.py to inject themeFontLang, docDefaults bidi, etc.
//   3. Apply arabic_numerals.py to convert digits + punctuation
//
// Covers every common element: titles, headings (H1/H2/H3), body paragraphs,
// bullet list, numbered list, a two-column table, and mixed Arabic+Latin
// content (IBAN, email) to verify the LTR-token isolation works.

const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
} = require("docx");

const FONT = "Arial";
const NAVY = "1F3864";
const DARK = "2E2E2E";
const GREY = "595959";
const LIGHT = "F2F2F2";

const border = { style: BorderStyle.SINGLE, size: 6, color: "BFBFBF" };
const cellBorders = { top: border, bottom: border, left: border, right: border };

// helpers
function p(text, opts = {}) {
  const {
    bold = false, size = 24, color = "000000",
    align = AlignmentType.START,
    spacingBefore = 0, spacingAfter = 120,
    rtl = true,
  } = opts;
  return new Paragraph({
    alignment: align,
    bidirectional: rtl,
    spacing: { before: spacingBefore, after: spacingAfter, line: 360 },
    children: [new TextRun({ text, font: FONT, size, bold, color, rightToLeft: rtl })],
  });
}

function h1(text) { return p(text, { bold: true, size: 32, color: NAVY, spacingBefore: 280, spacingAfter: 140 }); }
function h2(text) { return p(text, { bold: true, size: 26, color: NAVY, spacingBefore: 200, spacingAfter: 100 }); }
function h3(text) { return p(text, { bold: true, size: 24, color: DARK, spacingBefore: 160, spacingAfter: 80 }); }
function spacer(after = 120) { return new Paragraph({ bidirectional: true, spacing: { after }, children: [new TextRun({ text: "", font: FONT })] }); }

function bullet(text) {
  return new Paragraph({
    alignment: AlignmentType.START,
    bidirectional: true,
    spacing: { after: 80, line: 360 },
    indent: { right: 360 },
    children: [new TextRun({ text: "•  " + text, font: FONT, size: 24, rightToLeft: true })],
  });
}

function numbered(idx, text) {
  return new Paragraph({
    alignment: AlignmentType.START,
    bidirectional: true,
    spacing: { after: 80, line: 360 },
    indent: { right: 360 },
    children: [
      new TextRun({ text: idx + ". ", font: FONT, size: 24, bold: true, rightToLeft: true }),
      new TextRun({ text, font: FONT, size: 24, rightToLeft: true }),
    ],
  });
}

// Simple 2-column table demonstrating RTL flow
function sampleTable() {
  const cell = (t, opts = {}) => new TableCell({
    borders: cellBorders,
    width: { size: opts.w, type: WidthType.DXA },
    shading: opts.shade ? { fill: opts.shade, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: AlignmentType.START, bidirectional: true,
      children: [new TextRun({ text: t, font: FONT, size: 22, bold: !!opts.bold, rightToLeft: true })],
    })],
  });
  return new Table({
    visuallyRightToLeft: true,
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [3000, 6360],
    rows: [
      new TableRow({ children: [cell("الحقل", { w: 3000, bold: true, shade: LIGHT }), cell("القيمة", { w: 6360, bold: true, shade: LIGHT })] }),
      new TableRow({ children: [cell("الاسم", { w: 3000, bold: true, shade: LIGHT }), cell("محمد بن عبد الرحمن الموسى", { w: 6360 })] }),
      new TableRow({ children: [cell("المدينة", { w: 3000, bold: true, shade: LIGHT }), cell("المدينة المنورة", { w: 6360 })] }),
      new TableRow({ children: [cell("التاريخ", { w: 3000, bold: true, shade: LIGHT }), cell("١٩ / ٠٥ / ٢٠٢٦ م", { w: 6360 })] }),
      new TableRow({ children: [cell("الآيبان", { w: 3000, bold: true, shade: LIGHT }), cell("SA0280000301608016014488", { w: 6360 })] }),
    ],
  });
}

// Arabic "lorem ipsum" — classic Naseej / educational placeholder text
const lorem1 = "هذا النص هو مثال لنصٍّ يمكن أن يُستبدل في نفس المساحة، لقد تَمّ توليد هذا النص من مولِّد النص العربيِّ، حيث يمكنك أن تُولِّد مثل هذا النص أو العديد من النصوص الأخرى إضافةً إلى زيادة عدد الحروف التي يولِّدها التطبيق.";
const lorem2 = "إذا كنت تحتاج إلى عددٍ أكبر من الفقرات يتيح لك مولِّد النص العربي زيادة عدد الفقرات كما تريد، النص لن يبدو مقسَّماً ولا يحوي أخطاءً لغويةً، مولِّد النص العربي مفيد لمصمِّمي المواقع على وجه الخصوص، حيث يحتاج العميل في كثير من الأحيان أن يطَّلع على صورةٍ حقيقيةٍ لتصميم الموقع.";
const lorem3 = "ومن هنا وجب على المصمِّم أن يضع نصوصاً مؤقَّتة على التصميم ليُظهر للعميل الشكل كاملاً، دور مولِّد النص العربي هو أن يوفِّر على المصمِّم عناء البحث عن نصٍّ بديل لا علاقة له بالموضوع الذي يتحدَّث عنه التصميم فيظهر بشكلٍ لا يُوحي بالاحترافيَّة.";

const children = [
  // Title (centered)
  p("نموذج اختبار للنصوص العربية", { bold: true, size: 36, color: NAVY, align: AlignmentType.CENTER, spacingAfter: 200 }),
  p("مولَّد بواسطة إضافة claude-arabic-docs — للتحقق من اتجاه النص ومحاذاته في برنامج Microsoft Word", { italics: false, size: 20, color: GREY, align: AlignmentType.CENTER, spacingAfter: 320 }),

  // H1
  h1("أولاً: العناوين والفقرات"),

  h2("المقدِّمة"),
  p(lorem1),
  p(lorem2),

  h3("نقطةٌ فرعيةٌ ضمن المقدِّمة"),
  p(lorem3),

  // H1
  h1("ثانياً: القوائم النقطية والمرقَّمة"),

  h3("قائمة نقطية"),
  bullet("العنصر الأول من القائمة النقطية في نصٍّ عربيٍّ تجريبي."),
  bullet("العنصر الثاني، وفيه قيمة عددية كـ ٣٧٥٬٠٠٠ ريال سعودي."),
  bullet("العنصر الثالث، يحتوي على نسبةٍ مئوية: ١٥٪ من القيمة الكلية."),
  bullet("العنصر الرابع، يتضمَّن تاريخاً: ١٩/٠٥/٢٠٢٦م."),

  spacer(80),
  h3("قائمة مرقَّمة"),
  numbered("١", "الخطوة الأولى من العملية، تتمثَّل في جمع البيانات اللازمة."),
  numbered("٢", "الخطوة الثانية، يتم خلالها تحليل البيانات وفقَ المعايير المعتمدة."),
  numbered("٣", "الخطوة الثالثة، إعداد التقرير النهائي ورفعه للجهة المختصة."),

  // H1
  h1("ثالثاً: الجداول"),
  p("الجدول التالي يستعرض بياناتٍ تجريبيةً بسيطةً للتحقق من اتجاه الجداول في المستند:"),
  spacer(80),
  sampleTable(),
  spacer(120),

  // H1 — mixed content
  h1("رابعاً: محتوى عربيٌّ مختلطٌ بالأرقام والمعرِّفات اللاتينية"),
  p("يحتوي هذا القسم على نصٍّ يمزج بين العربية والمعرِّفات اللاتينية، للتأكُّد من أن المعرِّفات لا تتعرَّض لإعادة الترتيب:"),
  bullet("البريد الإلكتروني للتواصل: muh.moosa@gmail.com — يجب أن يبقى كما هو."),
  bullet("رقم الآيبان الدولي: SA0280000301608016014488 — لا يتحوَّل إلى أرقام هندية."),
  bullet("رقم الترخيص: FL-888252203 — يبقى لاتينياً لأنه معرِّف نظامي."),
  bullet("أما المبلغ النقدي ١٬٢٣٤٬٥٦٧ ريالاً والنسبة ١٥٪ فيُكتبان بالأرقام الهندية."),

  // H1 — closing
  h1("خامساً: الخاتمة"),
  p(lorem1),
  p("هذا المستند مولَّدٌ بالكامل من خلال سكربت docx-js، ثم مرَّ على إضافة claude-arabic-docs التي طبَّقت ستَّ طبقات من إعدادات RTL لضمان عرضٍ صحيحٍ في برنامج Microsoft Word على جميع المنصات (Windows / Mac / Online). إذا ظهرت هذه الفقرة محاذاةً إلى اليمين فإن الإضافة تعمل بشكلٍ صحيح. وإذا ظهرت إلى اليسار، يرجى إبلاغ المطوِّر فوراً."),

  spacer(240),
  p("•   •   •", { align: AlignmentType.CENTER, color: GREY }),
  p("انتهى المستند", { align: AlignmentType.CENTER, color: GREY, size: 20 }),
];

const doc = new Document({
  creator: "claude-arabic-docs skill test",
  title: "نموذج اختبار للنصوص العربية",
  styles: { default: { document: { run: { font: FONT, size: 24 } } } },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
      },
      bidirectional: true,
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  const outPath = "/sessions/amazing-clever-planck/mnt/استشارات/نموذج اختبار - النصوص العربية.docx";
  fs.writeFileSync(outPath, buffer);
  console.log("Wrote:", outPath, "size:", buffer.length, "bytes");
});
