#!/usr/bin/env python3
"""Make an RTL .docx render right-to-left in Word (Windows/Mac), Google Docs and LibreOffice.

    python3 harden_rtl.py in.docx [out.docx] [--lang ar-SA] [--check]

Fixes in place (or into out.docx). --check changes nothing, lists what would be fixed and
exits 1 unless the file is structurally clean. Idempotent. Needs lxml.
Lines starting with "warn:" are content issues the script cannot fix safely; they never fail --check.
"""
import argparse, re, shutil, sys, zipfile
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
q = lambda t: '{%s}%s' % (W, t)
# Hebrew, Arabic, Syriac, Thaana, NKo, Samaritan, Arabic Ext-A, presentation forms. Escaped on purpose:
# U+FEFF is invisible and gets lost when the file is copied, which turns the class end into a literal "-".
RTL = re.compile('[֐-ࣿיִ-﷿ﹰ-﻿]')
LATIN = re.compile('[A-Za-z]')
ISOLATES = re.compile('[⁦-⁩‪-‮]')
WINDOWS_ONLY_FONTS = {'traditional arabic', 'simplified arabic', 'simplified arabic fixed', 'sakkal majalla',
                      'arabic typesetting', 'andalus', 'aldhabi', 'microsoft uighur', 'urdu typesetting',
                      'arabic transparent', 'times new roman (arabic)'}
LATIN_ONLY_FONTS = {'calibri', 'calibri light', 'cambria', 'aptos', 'aptos display', 'consolas', 'georgia',
                    'verdana', 'garamond', 'helvetica', 'liberation sans', 'liberation serif', 'carlito', 'caladea'}
CONTENT_PART = re.compile(r'word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml')

ORDER = {k: v.split() for k, v in {
    'pPr': 'pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs '
           'suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd '
           'snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment '
           'textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange',
    'rPr': 'rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid '
           'vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs '
           'em lang eastAsianLayout specVanish oMath rPrChange',
    'sectPr': 'headerReference footerReference footnotePr endnotePr type pgSz pgMar paperSrc pgBorders lnNumType '
              'pgNumType cols formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid printerSettings '
              'sectPrChange',
    'tblPr': 'tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize tblW jc tblCellSpacing '
             'tblInd tblBorders shd tblLayout tblCellMar tblLook tblCaption tblDescription tblPrChange'}.items()}
SETTINGS_AFTER = ('clrSchemeMapping doNotIncludeSubdocsInStats doNotAutoCompressPictures forceUpgrade captions '
                  'readModeInkLockDown smartTagType schemaLibrary shapeDefaults doNotEmbedSmartTags decimalSymbol '
                  'listSeparator').split()
CS_PAIRS = (('b', 'bCs'), ('i', 'iCs'), ('sz', 'szCs'))


def local(e): return etree.QName(e).localname
def text_of(el): return ''.join(t.text or '' for t in el.iter(q('t')))


def reorder(el):
    order = ORDER[local(el)]
    known = [k for k in el if isinstance(k.tag, str) and local(k) in order]
    want = sorted(known, key=lambda k: order.index(local(k)))
    if want == known: return 0
    for k in known: el.remove(k)
    for i, k in enumerate(want): el.insert(i, k)
    return 1


def child(parent, tag, first=False):
    e = parent.find(q(tag))
    if e is None:
        e = etree.Element(q(tag))
        parent.insert(0, e) if first else parent.append(e)
    return e


def ensure(parent, tag, log, what):
    if parent.find(q(tag)) is None:
        parent.append(etree.Element(q(tag))); log.append(what)


def jc_logical(ppr, log, what):
    jc = ppr.find(q('jc'))
    if jc is not None and jc.get(q('val')) in ('right', 'left'):
        log.append(f"{what}: jc {jc.get(q('val'))} -> start"); jc.set(q('val'), 'start')


def mirror_cs(rpr):
    """Word styles RTL text from bCs/iCs/szCs and rFonts@cs, not b/i/sz/ascii. Copy what is missing."""
    changed = 0
    for latin, cs in CS_PAIRS:
        src = rpr.find(q(latin))
        if src is not None and rpr.find(q(cs)) is None:
            e = etree.SubElement(rpr, q(cs))
            for k, v in src.attrib.items(): e.set(k, v)
            changed += 1
    fonts = rpr.find(q('rFonts'))
    if fonts is not None and fonts.get(q('cs')) is None and fonts.get(q('cstheme')) is None:
        face = fonts.get(q('ascii')) or fonts.get(q('hAnsi'))
        if face and face.lower() not in LATIN_ONLY_FONTS: fonts.set(q('cs'), face); changed += 1
    return changed


def fix_settings(root, lang, log):
    tfl = root.find(q('themeFontLang'))
    if tfl is None:
        tfl = etree.Element(q('themeFontLang')); tfl.set(q('val'), 'en-US')
        nxt = next((k for k in root if isinstance(k.tag, str) and local(k) in SETTINGS_AFTER), None)
        nxt.addprevious(tfl) if nxt is not None else root.append(tfl)
    if tfl.get(q('bidi')) is None:
        tfl.set(q('bidi'), lang); log.append(f'settings: themeFontLang bidi={lang}')


def fix_styles(root, lang, log):
    dd = root.find(q('docDefaults'))
    if dd is None: dd = etree.Element(q('docDefaults')); root.insert(0, dd)
    rpr = child(child(dd, 'rPrDefault', first=True), 'rPr')
    lg = child(rpr, 'lang')
    if lg.get(q('bidi')) is None: lg.set(q('bidi'), lang); log.append(f'docDefaults: lang bidi={lang}')
    if mirror_cs(rpr): log.append('docDefaults: complex-script size/bold/font mirrored')
    ppr = child(child(dd, 'pPrDefault'), 'pPr')
    ensure(ppr, 'bidi', log, 'pPrDefault: bidi')
    if ppr.find(q('jc')) is None:
        etree.SubElement(ppr, q('jc')).set(q('val'), 'start'); log.append('pPrDefault: jc start')
    else:
        jc_logical(ppr, log, 'pPrDefault')
    for st in root.iter(q('style')):
        sid, kind = st.get(q('styleId')), st.get(q('type'))
        srpr = st.find(q('rPr'))
        if srpr is not None and kind in ('paragraph', 'character') and mirror_cs(srpr):
            log.append(f'style {sid}: complex-script size/bold/font mirrored')
        if kind != 'paragraph': continue
        sppr = st.find(q('pPr'))
        if sid == 'Normal':
            if sppr is None: sppr = etree.SubElement(st, q('pPr'))
            ensure(sppr, 'bidi', log, 'style Normal: bidi')
        if sppr is not None: jc_logical(sppr, log, f'style {sid}')


def fix_content(root, log, warn, name):
    n = dict(p=0, ltr=0, r=0, cs=0, tbl=0, jc=0)
    for p in root.iter(q('p')):
        txt = text_of(p)
        ppr = p.find(q('pPr'))
        if not RTL.search(txt):
            # Defaults are RTL, so a Latin-only paragraph must opt out explicitly or it inherits RTL.
            if LATIN.search(txt):
                if ppr is None: ppr = etree.Element(q('pPr')); p.insert(0, ppr)
                if ppr.find(q('bidi')) is None:
                    etree.SubElement(ppr, q('bidi')).set(q('val'), '0'); n['ltr'] += 1
            continue
        if ppr is None: ppr = etree.Element(q('pPr')); p.insert(0, ppr)
        bidi = ppr.find(q('bidi'))
        if bidi is None: ppr.append(etree.Element(q('bidi'))); n['p'] += 1
        elif bidi.get(q('val')) in ('0', 'false', 'off'): del bidi.attrib[q('val')]; n['p'] += 1
        jc = ppr.find(q('jc'))
        if jc is not None and jc.get(q('val')) in ('right', 'left'): jc.set(q('val'), 'start'); n['jc'] += 1
        mark = ppr.find(q('rPr'))  # paragraph-mark run properties (affect list numbers and empty lines)
        if mark is not None: n['cs'] += bool(mirror_cs(mark))
    for r in root.iter(q('r')):
        txt = text_of(r)
        if not RTL.search(txt): continue
        rpr = r.find(q('rPr'))
        if rpr is None: rpr = etree.Element(q('rPr')); r.insert(0, rpr)
        if rpr.find(q('rtl')) is None: rpr.append(etree.Element(q('rtl'))); n['r'] += 1
        n['cs'] += bool(mirror_cs(rpr))
        short = txt.strip()[:40]
        if LATIN.search(txt): warn.add(f'{name}: run mixes RTL and Latin letters (split the Latin token into its own non-rtl run): "{short}"')
        if re.search(r'[,;?]\s*$', txt): warn.add(f'{name}: RTL run ends in Latin , ; or ? (use ، ؛ ؟): "{short}"')
    for t in root.iter(q('t')):
        if t.text and ISOLATES.search(t.text):
            warn.add(f'{name}: bidi control characters (U+2066-2069/U+202A-202E) render as boxes in some Word fonts; use separate runs'); break
    for t in root.iter(q('tbl')):
        if not RTL.search(text_of(t)): continue
        tp = t.find(q('tblPr'))
        if tp is None: tp = etree.Element(q('tblPr')); t.insert(0, tp)
        if tp.find(q('bidiVisual')) is None: tp.append(etree.Element(q('bidiVisual'))); n['tbl'] += 1
    for s in root.iter(q('sectPr')): ensure(s, 'bidi', log, 'sectPr: bidi')
    names = dict(p='paragraph bidi added', ltr='Latin-only paragraph set LTR (bidi 0)', r='run rtl added', cs='runs got complex-script size/bold/font',
                 tbl='table bidiVisual added', jc='paragraph jc -> start')
    log.extend(f'{name}: {v} {names[k]}' for k, v in n.items() if v)


def fix_numbering(root, log):
    for lvl in root.iter(q('lvl')):
        ppr = lvl.find(q('pPr'))
        if ppr is None: ppr = etree.SubElement(lvl, q('pPr'))
        ensure(ppr, 'bidi', log, 'numbering level: bidi')


def font_warnings(roots, warn):
    faces = {f.get(q('cs')) for r in roots for f in r.iter(q('rFonts')) if f.get(q('cs'))}
    odd = sorted(f for f in faces if f.lower() in WINDOWS_ONLY_FONTS)
    if odd:
        warn.add(f'fonts {", ".join(odd)} ship with Windows/Office only: Google Docs, LibreOffice on Linux and '
                 'some Mac setups substitute them. Fine for Windows Word readers; otherwise prefer Arial, Amiri, '
                 'Noto Naskh Arabic or Noto Sans Arabic, or deliver a PDF made where the font is installed')


def process(src, dst, lang, check):
    log, warn, parts, roots = [], set(), {}, []
    with zipfile.ZipFile(src) as zin:
        for it in zin.infolist():
            name, data = it.filename, zin.read(it.filename)
            kind = ('settings' if name == 'word/settings.xml' else 'styles' if name == 'word/styles.xml'
                    else 'numbering' if name == 'word/numbering.xml'
                    else 'content' if CONTENT_PART.fullmatch(name) else None)
            if kind:
                root = etree.fromstring(data); roots.append(root)
                if kind == 'settings': fix_settings(root, lang, log)
                elif kind == 'styles': fix_styles(root, lang, log)
                elif kind == 'numbering': fix_numbering(root, log)
                else: fix_content(root, log, warn, name)
                moved = sum(reorder(e) for e in root.iter() if isinstance(e.tag, str) and local(e) in ORDER)
                if moved: log.append(f'{name}: {moved} property blocks put in schema order')
                data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
            parts[it] = data
    font_warnings(roots, warn)
    print('\n'.join(log) if log else 'clean')
    for w in sorted(warn): print('warn:', w)
    if check: return 1 if log else 0
    with zipfile.ZipFile(dst + '.tmp', 'w', zipfile.ZIP_DEFLATED) as z:
        for it, data in parts.items(): z.writestr(it, data)
    shutil.move(dst + '.tmp', dst)
    return 0


if __name__ == '__main__':
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('src'); ap.add_argument('dst', nargs='?')
    ap.add_argument('--lang', default='ar-SA', help='ar-SA (default), ar-EG, ar-AE, ar-MA, he-IL, fa-IR, fa-AF, ur-PK')
    ap.add_argument('--check', action='store_true')
    a = ap.parse_args()
    sys.exit(process(a.src, a.dst or a.src, a.lang, a.check))
