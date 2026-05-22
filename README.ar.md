# claude-arabic-docs

> إضافةٌ (skill) للذكاء الاصطناعي Claude تُعالج فئةً من الأخطاء لا يُحدِّثك عنها أحدٌ حتى تُسلِّم تقريراً عربياً من ٣٠ صفحة، ثم يفتحه العميل على Word for Mac وتجده محاذىً إلى اليسار. النسخة الإنجليزية: [README.md](./README.md)

[![الترخيص: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![Claude Skill](https://img.shields.io/badge/Claude-Skill-blueviolet)](https://docs.claude.com/skills)
[![اللغات: ar · he · fa · ur](https://img.shields.io/badge/RTL-AR%20·%20HE%20·%20FA%20·%20UR-green)](#اللغات-المدعومة)

## الخطأ الذي وُجدت هذه الإضافة لإصلاحه

تُولِّد مستنداً عربياً بصيغة `.docx` برمجياً عبر `docx-js` أو `python-docx` أو أيٍّ من المكتبات الشائعة. كلُّ فقرةٍ مُعلَّمةٌ بـ RTL، وكلُّ مقطع نصٍّ يحمل `<w:rtl/>`، وكلُّ جدولٍ يحوي `<w:bidiVisual/>`. تُحوِّله إلى PDF عبر LibreOffice فيظهر مثاليّاً. تُسلِّمه للعميل.

**يفتحه العميل في Microsoft Word ويجد كل العناوين والفقرات مُحاذاةً إلى اليسار. الجداول صحيحة. فقط النص الأساسي مكسور.**

تبحث في Google. تضيف `<w:bidi/>` لكل فقرة. يبقى مكسوراً في Word. تضيفه للقسم (section). يبقى مكسوراً. تضبط لغة `docDefaults` على `ar-SA`. يبقى مكسوراً. تضيف `<w:pPrDefault>` مع bidi. يبقى مكسوراً.

بعد يومَين من تفكيك XML سطراً سطراً، تكتشف أن Microsoft Word يتجاهل كلَّ ما سبق إذا لم يحتوِ `settings.xml` على `<w:themeFontLang w:bidi="ar-SA"/>`. وهذا الإعداد لا تكتبه أيُّ مكتبة توليد مستندات — يُكتب فقط حين يحفظ Word الحيُّ الملف، مأخوذاً من تخطيط لوحة المفاتيح في نظام التشغيل. لذا فإن **كلَّ** مستند مولَّد برمجياً يبدأ مكسوراً في Word.

تُصلح ذلك فيظهر المستند صحيحاً... تقريباً — إلا أن العناوين لا تزال مُحاذاةً إلى اليسار. بعد مزيدٍ من التفكيك تكتشف أن Word for Mac يُعيد تفسير `<w:jc w:val="right"/>` على أنها "نهاية منطقية" (= يسار في وضع RTL)، لا "يمين فيزيائي". تُغيِّر كل شيءٍ إلى `start`/`end` فينجح أخيراً.

هذه الإضافة تَحزم الإصلاحات الستَّة كلَّها — المكتشفة بالطريقة الصعبة على مدى ساعاتٍ طويلة — في وصفٍ يَنشَط على المستندات RTL، وسكربت تقوية (hardening) يُمكن تشغيله كمعالجة لاحقة، ومحوِّل أرقام وترقيم للوفاء بأعراف الطباعة العربية.

## ما الذي تفعله

عندما يُولِّد Claude مستنداً بصيغة `.docx` (أو `.pptx` / `.xlsx`) يحتوي على عربية أو عبرية أو فارسية أو أردية أو أيِّ سكربتٍ يُكتب من اليمين لليسار، تَجعل هذه الإضافة Claude يقوم بـ:

1. استخدام المحاذاة المنطقية (`start`/`end`) بدل الفيزيائية (`left`/`right`) في التوليد.
2. حقن **الطبقات الست** الفعليَّة التي يفحصها Word (الجدول الكامل أدناه).
3. تحويل الأرقام إلى الهندية (٠–٩) والفواصل والنسب اللاتينية إلى نظائرها العربية (٬ ٪ ، ؛ ؟)، مع الإبقاء على رموز الآيبان/البريد الإلكتروني/التراخيص بالأرقام اللاتينية عبر استدلالٍ ذكي.

## التثبيت

### Cowork (تطبيق Claude لسطح المكتب)

1. نزِّل [`claude-arabic-docs.skill`](./claude-arabic-docs.skill) (أو ابنِه: راجع "البناء من المصدر" أدناه).
2. في Cowork: انقر منتقي الإضافات ← تثبيت إضافة ← اختر ملف `.skill`.

### Claude Code (واجهة الأوامر)

```bash
# انسخ مجلد الإضافة إلى مسار إضافات Claude Code
cp -r claude-arabic-docs ~/.claude/skills/
```

### Claude.ai (الويب)

الإضافات حالياً غير قابلة للتثبيت عبر واجهة الويب. استخدم Cowork أو Claude Code.

## الاستخدام

بعد التثبيت، تَنشَط الإضافة تلقائياً عند رصد Claude لأي طلبٍ يتعلَّق بمستندٍ عربي/RTL. لا حاجة لاستدعاءٍ يدوي. إن أردت أن تكون صريحاً:

> "أنشئ لي مستند Word عربياً عن كذا، واستخدم إضافة claude-arabic-docs."

إن أردت تشغيل السكربتات المضمَّنة يدوياً:

```bash
# تقوية أي ملف .docx موجود (في مكانه)
python scripts/harden_rtl.py document.docx

# أو إلى ملف مختلف
python scripts/harden_rtl.py input.docx -o output.docx

# تقرير مفصَّل عمَّا تغيَّر
python scripts/harden_rtl.py document.docx --report

# فحص قواعد RTL على مستوى المحتوى — يُنهي بالرمز 1 عند وجود أخطاء
python scripts/harden_rtl.py document.docx --validate

# الفحص مع إعادة كتابة jc="right" إلى "start" حيث يكون ذلك آمناً بلا لبس
python scripts/harden_rtl.py document.docx --validate --fix-jc

# لغة RTL مختلفة
python scripts/harden_rtl.py document.docx --locale he-IL   # عبرية
python scripts/harden_rtl.py document.docx --locale fa-IR   # فارسية

# تحويل الأرقام والترقيم في نصٍّ عربي
echo "بمبلغ 375,000 ريال (15%)" | python scripts/arabic_numerals.py
# → "بمبلغ ٣٧٥٬٠٠٠ ريال (١٥٪)"
```

سكربت التقوية **متعدِّد التشغيل بأمان (idempotent)** — تشغيله مرتَّين يعني أن المرة الثانية لا تفعل شيئاً.

## الطبقات الست باختصار

| # | الطبقة | XML | لماذا |
|---|---|---|---|
| ٠.٠ | `settings.xml` themeFontLang | `<w:themeFontLang w:bidi="ar-SA"/>` | **المفتاح الرئيسي.** بدونه لا يَنشُط محرِّك RTL في Word أصلاً. |
| ٠   | docDefaults rPr lang        | `<w:lang w:bidi="ar-SA"/>`         | يُخبر Word بلغة السكربت المركَّب. الجداول تعمل بدونه؛ لا شيء غيرها يعمل. |
| ٠.٥ | docDefaults pPrDefault       | `<w:pPr><w:bidi/><w:jc w:val="start"/></w:pPr>` | الاتجاه الافتراضي لنمط فقرة `Normal`. بدونه تَعود العناوين/المتن إلى LTR. |
| ١   | bidi على مستوى القسم         | `<w:bidi/>` في `<w:sectPr>`        | اتجاه القراءة على مستوى القسم. حسَّاسٌ لموقعه في schema (يجب أن يأتي قبل `<w:docGrid>`). |
| ٢   | bidi-visual للجدول           | `<w:bidiVisual/>` في `<w:tblPr>`   | يَقلب ترتيب الأعمدة بحيث تَظهر أوَّل خليةٍ على اليمين. |
| ٣   | علامات الفقرة والمقطع        | `<w:bidi/>` + `<w:rtl/>`           | RTL على مستوى العنصر. رخيصةٌ — افعلها من البداية. |
| ٥   | محاذاة منطقية                 | فضِّل `w:jc="start"`/`"end"`       | Word for Mac يُعيد تفسير `left`/`right` الفيزيائية على أنها منطقية في وضع RTL. مضادٌّ للحدس — استخدم start/end. |

القواعد الكاملة مع التبرير: راجع [`SKILL.md`](./SKILL.md).

## اللغات المدعومة

- `ar-SA` العربية (السعودية) — **الافتراضي**
- `ar-EG`, `ar-AE`, `ar-MA` العربية (مصر، الإمارات، المغرب)
- `he-IL` العبرية (إسرائيل)
- `fa-IR` الفارسية (إيران)
- `ur-PK` الأردية (باكستان)

السكربتات الأخرى التي تُكتب من اليمين لليسار (السريانية، الثاآنا، النكو، المندائية، السامرية) تُكتشَف بنطاق Unicode عند حقن علامات الفقرات/المقاطع، لكنَّ المفتاح الرئيسي `themeFontLang` يحتاج رمز لغةٍ صريحاً يَقبَله Word. PRات لتوسيع القائمة مرحَّبٌ بها.

## محتويات المستودع

```
claude-arabic-docs/
├── SKILL.md                       فهرس القواعد الكامل مع التبرير (اقرأ هذا)
├── README.md                      النسخة الإنجليزية
├── README.ar.md                   أنت هنا
├── LICENSE                        MIT
├── CHANGELOG.md                   تاريخ اكتشاف كل قاعدة
├── scripts/
│   ├── harden_rtl.py              معالجة لاحقة لأي .docx — تطبق الطبقات الست + وضع الفحص --validate
│   └── arabic_numerals.py         تحويل الأرقام الغربية والترقيم اللاتيني إلى صيغ عربية
├── references/
│   └── python-docx-template.md    طبقة مساعِدة جاهزة لـ python-docx (تُطبّق القواعد ١–٨ بالبناء)
└── examples/
    ├── build_test_arabic.js       مثال docx-js يستخدم أعراف الإضافة
    └── sample-output.docx         المستند التجريبي المولَّد
```

## البناء من المصدر

```bash
git clone https://github.com/<your-username>/claude-arabic-docs
cd claude-arabic-docs

# حزم ملف .skill (هو فقط ZIP عادي يحتوي SKILL.md والسكربتات والمراجع)
python -m zipfile -c claude-arabic-docs.skill SKILL.md scripts/ references/

# اختبار تشغيل سكربت التقوية
python scripts/harden_rtl.py examples/sample-output.docx --report
```

## الترخيص

MIT — راجع [`LICENSE`](./LICENSE). استخدمها، فرِّعها، طوِّرها. وإن طوَّرتها، رجاءً افتح PR هنا حتى لا يضطر الشخص التالي إلى إعادة اكتشاف ما أصلحته.
