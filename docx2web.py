#!/usr/bin/env python3
"""
docx2web.py - turn a Word (.docx) manual into ONE searchable, clickable HTML file.

Requirements: Python 3.6+ ONLY. No packages to install (a .docx is a zip of XML files,
read here with the standard library).

Usage:
    python docx2web.py manual.docx -o manual.html
    python docx2web.py manual.docx -o manual.html --config configs/my_manual.json
    python docx2web.py manual.docx --title "Safety Manual" --no-embed

What it keeps from Word:
  * Heading 1-9 styles (or outline levels) -> sidebar table of contents
  * Word's own list and heading numbering (A., 1., a., 7302.12 ...) computed exactly
  * bold / italic / underline / superscript, line breaks
  * tables, including merged cells and repeating header rows
  * pictures (PNG/JPG/GIF/SVG), embedded in the file
  * hyperlinks, cross-references and bookmarks, footnotes and endnotes
  * text boxes
Skipped: Word's own Table of Contents (the sidebar replaces it), headers/footers, page numbers,
deleted tracked changes, comments.
"""
import argparse
import base64
import html
import os
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import webviewer as wv  # noqa: E402

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
V = 'urn:schemas-microsoft-com:vml'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'


def w(t):
    return '{%s}%s' % (W, t)


def wattr(el, name, default=None):
    if el is None:
        return default
    return el.get(w(name), default)


def on(el):
    """True for <w:b/>, <w:b w:val="1"/>, <w:b w:val="true"/>; False for val=0/false/none."""
    if el is None:
        return False
    v = el.get(w('val'))
    return v is None or v.lower() not in ('0', 'false', 'off', 'none')


DEFAULTS = {
    "title": None, "subtitle": "", "org": "", "meta": "",
    "chapter_level": "auto",       # heading level used for chapters (1 = Heading 1)
    "chapter_regex": r"^(?:CHAPTER|Chapter|PART|Part)\s+(\w+)[\s:.-]*(.*)$",
    "section_regex": r"^(\d+(?:\.\d+)*)\s+(.*)$",    # pulls "7302.12" out of typed heading text
    "exhibit_regex": None,         # headings matching this go under "Exhibits", e.g. "^Exhibit\\s+(\\S+)\\s*(.*)$"
    "procedure_regex": None,       # headings matching this go under "Procedures & tasks", e.g. "^((?:PRO|TSK)-[\\d.]+[A-Z]?\\d?)\\s*(.*)$"
    "autolink": [],                # same rules as pdf2web (see configs/lom_2021.json)
    "fallback_headings": "auto",   # if the document uses no Heading styles: detect bold numbered lines
    "heading_patterns": [          # used by the fallback; first match wins
        {"regex": r"^(?:CHAPTER|Chapter)\s+\d+", "level": 1},
        {"regex": r"^(?:Exhibit|EXHIBIT)\s+\S+", "level": 2},
        {"regex": r"^\d+\s+\S", "level": 2},
        {"regex": r"^\d+\.\d\s+\S", "level": 3},
        {"regex": r"^\d+\.\d{2,}\s+\S", "level": 4},
    ],
}

SYMBOL_MAP = {'': '•', '': '▪', '': '➢', '': '❖', '': '✓', '': '➔',
              '': '■', '': '❑', '§': '▪', 'o': '◦', '': '–'}
IMG_MIME = {'.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif',
            '.bmp': 'image/bmp', '.svg': 'image/svg+xml', '.webp': 'image/webp'}


def log(*a):
    print(*a, file=sys.stderr)


# =============================================================================
# Number formats
# =============================================================================
def roman(n):
    vals = [(1000, 'm'), (900, 'cm'), (500, 'd'), (400, 'cd'), (100, 'c'), (90, 'xc'),
            (50, 'l'), (40, 'xl'), (10, 'x'), (9, 'ix'), (5, 'v'), (4, 'iv'), (1, 'i')]
    out = ''
    for v, s in vals:
        while n >= v:
            out += s
            n -= v
    return out


def letters(n):
    # Word style: a..z, aa..zz, aaa..
    if n <= 0:
        return ''
    ch = chr(ord('a') + (n - 1) % 26)
    return ch * ((n - 1) // 26 + 1)


def fmt_num(n, fmt):
    if fmt in ('upperLetter',):
        return letters(n).upper()
    if fmt == 'lowerLetter':
        return letters(n)
    if fmt == 'upperRoman':
        return roman(n).upper()
    if fmt == 'lowerRoman':
        return roman(n)
    if fmt == 'decimalZero':
        return '%02d' % n
    if fmt == 'none':
        return ''
    if fmt == 'ordinal':
        return '%d%s' % (n, 'th' if 10 <= n % 100 <= 20 else {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th'))
    return str(n)


# =============================================================================
# The document package
# =============================================================================
class Docx(object):
    def __init__(self, path):
        self.z = zipfile.ZipFile(path)
        self.names = set(self.z.namelist())
        main = 'word/document.xml'
        # find the main part through [Content_Types] / _rels if it has an unusual name
        if main not in self.names:
            rels = ET.fromstring(self.z.read('_rels/.rels'))
            for r in rels:
                if r.get('Type', '').endswith('/officeDocument'):
                    main = r.get('Target').lstrip('/')
        self.main = main
        self.base = posixpath.dirname(main)
        self.doc = ET.fromstring(self.z.read(main))
        self.rels = self._rels(main)
        self.styles = self._styles()
        self.abstract, self.nums, self.num_override = self._numbering()
        self.notes = {'footnote': self._notes('footnotes'), 'endnote': self._notes('endnotes')}

    def read(self, part):
        return self.z.read(part) if part in self.names else None

    def _rels(self, part):
        rp = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
        out = {}
        data = self.read(rp)
        if data:
            for r in ET.fromstring(data):
                out[r.get('Id')] = (r.get('Target'), r.get('TargetMode'), r.get('Type', ''))
        return out

    def _styles(self):
        st = {}
        data = self.read(posixpath.join(self.base, 'styles.xml'))
        if not data:
            return st
        for s in ET.fromstring(data).iter(w('style')):
            sid = wattr(s, 'styleId')
            name_el = s.find(w('name'))
            ppr, rpr = s.find(w('pPr')), s.find(w('rPr'))
            numpr = ppr.find(w('numPr')) if ppr is not None else None
            ol = ppr.find(w('outlineLvl')) if ppr is not None else None
            ind = ppr.find(w('ind')) if ppr is not None else None
            st[sid] = {
                'name': (wattr(name_el, 'val') or sid or '').strip(),
                'type': wattr(s, 'type'),
                'based': wattr(s.find(w('basedOn')), 'val'),
                'outline': int(wattr(ol, 'val')) if ol is not None and wattr(ol, 'val', '').isdigit() else None,
                'numId': wattr(numpr.find(w('numId')), 'val') if numpr is not None and numpr.find(w('numId')) is not None else None,
                'ilvl': wattr(numpr.find(w('ilvl')), 'val') if numpr is not None and numpr.find(w('ilvl')) is not None else None,
                'b': on(rpr.find(w('b'))) if rpr is not None and rpr.find(w('b')) is not None else None,
                'i': on(rpr.find(w('i'))) if rpr is not None and rpr.find(w('i')) is not None else None,
                'indLeft': (wattr(ind, 'left') or wattr(ind, 'start')) if ind is not None else None,
            }
        return st

    def style_get(self, sid, key, depth=0):
        while sid and depth < 20:
            s = self.styles.get(sid)
            if not s:
                return None
            if s.get(key) is not None:
                return s[key]
            sid = s['based']
            depth += 1
        return None

    def style_name(self, sid):
        s = self.styles.get(sid)
        return s['name'] if s else (sid or '')

    def _numbering(self):
        abstract, nums, override = {}, {}, {}
        data = self.read(posixpath.join(self.base, 'numbering.xml'))
        if not data:
            return abstract, nums, override
        root = ET.fromstring(data)
        links = {}
        for an in root.findall(w('abstractNum')):
            aid = wattr(an, 'abstractNumId')
            lv = {}
            for l in an.findall(w('lvl')):
                il = int(wattr(l, 'ilvl', '0'))
                lv[il] = {
                    'start': int(wattr(l.find(w('start')), 'val', '1') or 1),
                    'fmt': wattr(l.find(w('numFmt')), 'val', 'decimal'),
                    'text': wattr(l.find(w('lvlText')), 'val', ''),
                    'restart': wattr(l.find(w('lvlRestart')), 'val'),
                    'legal': l.find(w('isLgl')) is not None,
                }
            abstract[aid] = lv
            sl = an.find(w('numStyleLink'))
            if sl is not None:
                links[aid] = wattr(sl, 'val')
        for n in root.findall(w('num')):
            nid = wattr(n, 'numId')
            nums[nid] = wattr(n.find(w('abstractNumId')), 'val')
            ov = {}
            for o in n.findall(w('lvlOverride')):
                so = o.find(w('startOverride'))
                if so is not None:
                    ov[int(wattr(o, 'ilvl', '0'))] = int(wattr(so, 'val', '1'))
                lvl = o.find(w('lvl'))
                if lvl is not None:   # full level override: keep it simple, copy format
                    il = int(wattr(o, 'ilvl', '0'))
                    abstract.setdefault('ov-%s' % nid, dict(abstract.get(nums[nid], {})))
                    abstract['ov-%s' % nid][il] = {
                        'start': int(wattr(lvl.find(w('start')), 'val', '1') or 1),
                        'fmt': wattr(lvl.find(w('numFmt')), 'val', 'decimal'),
                        'text': wattr(lvl.find(w('lvlText')), 'val', ''), 'restart': None, 'legal': False}
                    nums[nid] = 'ov-%s' % nid
            override[nid] = ov
        # abstract numbering that just points at a numbering style ("numStyleLink")
        for aid, stname in links.items():
            if abstract.get(aid):
                continue
            for sid, s in self.styles.items():
                if sid == stname and s.get('numId') in nums:
                    abstract[aid] = abstract.get(nums[s['numId']], {})
        return abstract, nums, override

    def _notes(self, kind):
        data = self.read(posixpath.join(self.base, kind + '.xml'))
        out = {}
        if not data:
            return out
        for n in ET.fromstring(data):
            t = wattr(n, 'type')
            if t in ('separator', 'continuationSeparator', 'continuationNotice'):
                continue
            out[wattr(n, 'id')] = n
        return out

    def media(self, rid):
        r = self.rels.get(rid)
        if not r or r[1] == 'External':
            return None, None
        path = posixpath.normpath(posixpath.join(self.base, r[0])) if not r[0].startswith('/') else r[0].lstrip('/')
        return path, self.read(path)


# =============================================================================
# Converter
# =============================================================================
class Converter(object):
    def __init__(self, dx, cfg):
        self.dx = dx
        self.cfg = cfg
        self.linker = wv.RefLinker(cfg.get('autolink'))
        self.counters = {}     # abstractNumId -> [count per level]
        self.num_started = set()
        self.headings = []     # dicts in document order
        self.used_ids = set()
        self.notes_used = []   # (kind, id, number)
        self.images = self.skipped_images = 0
        self.has_heading_styles = self._detect_heading_styles()

    # ---------------- numbering ----------------
    def numbering(self, p_pr, style_id):
        num_id = ilvl = None
        np_ = p_pr.find(w('numPr')) if p_pr is not None else None
        if np_ is not None:
            num_id = wattr(np_.find(w('numId')), 'val') if np_.find(w('numId')) is not None else None
            ilvl = wattr(np_.find(w('ilvl')), 'val') if np_.find(w('ilvl')) is not None else None
        if num_id is None:
            num_id = self.dx.style_get(style_id, 'numId')
        if ilvl is None:
            ilvl = self.dx.style_get(style_id, 'ilvl')
            if ilvl is None and num_id is not None:
                # style-linked numbering: find which level points at this style
                ilvl = '0'
        if not num_id or num_id == '0' or num_id not in self.dx.nums:
            return None
        ilvl = int(ilvl or 0)
        aid = self.dx.nums[num_id]
        levels = self.dx.abstract.get(aid) or {}
        if ilvl not in levels:
            return None
        c = self.counters.setdefault(aid, [None] * 10)
        if num_id not in self.num_started:
            self.num_started.add(num_id)
            for lv, start in self.dx.num_override.get(num_id, {}).items():
                c[lv] = start - 1
                for d in range(lv + 1, 10):
                    c[d] = None
        lvdef = levels[ilvl]
        c[ilvl] = lvdef['start'] if c[ilvl] is None else c[ilvl] + 1
        for d in range(ilvl + 1, 10):
            c[d] = None
        if lvdef['fmt'] == 'bullet':
            sym = lvdef['text'] or '•'
            return {'label': SYMBOL_MAP.get(sym, sym if sym.strip() else '•'), 'ilvl': ilvl, 'bullet': True}

        def sub(m):
            k = int(m.group(1)) - 1
            ld = levels.get(k, {'start': 1, 'fmt': 'decimal'})
            n = c[k] if c[k] is not None else ld['start']
            f = 'decimal' if (lvdef.get('legal') and k < ilvl) else ld['fmt']
            return fmt_num(n, f)
        label = re.sub(r'%(\d)', sub, lvdef['text'] or '')
        return {'label': label, 'ilvl': ilvl, 'bullet': False}

    # ---------------- headings ----------------
    def _detect_heading_styles(self):
        for p in self.dx.doc.iter(w('p')):
            ppr = p.find(w('pPr'))
            sid = wattr(ppr.find(w('pStyle')), 'val') if ppr is not None and ppr.find(w('pStyle')) is not None else None
            if sid and self.heading_level(ppr, sid):
                return True
        return False

    def heading_level(self, ppr, sid):
        name = self.dx.style_name(sid).lower()
        m = re.match(r'^heading\s*(\d)$', name)
        if m:
            return int(m.group(1))
        ol = ppr.find(w('outlineLvl')) if ppr is not None else None
        if ol is not None and (wattr(ol, 'val') or '').isdigit() and int(wattr(ol, 'val')) < 9:
            return int(wattr(ol, 'val')) + 1
        lvl = self.dx.style_get(sid, 'outline')
        if lvl is not None and lvl < 9:
            return lvl + 1
        return 0

    # ---------------- inline content ----------------
    def runs_html(self, p, extra_blocks):
        """Render the inline content of a paragraph. Returns (html, plain_text, all_bold)."""
        toks = []          # (html, link, plain, bold)
        fields = []        # stack of {'instr', 'phase', 'link'}
        state = {'link': None}

        def cur_link():
            for f in reversed(fields):
                if f.get('link'):
                    return f['link']
            return state['link']

        def in_instr():
            return any(f['phase'] == 'instr' for f in fields)

        def add_text(text, rpr_fmt):
            if not text or in_instr():
                return
            b, i, u, sup, sub, strike = rpr_fmt
            h = self.linker.mark(html.escape(text, quote=False))
            if sup:
                h = '<sup>%s</sup>' % h
            elif sub:
                h = '<sub>%s</sub>' % h
            if strike:
                h = '<s>%s</s>' % h
            if u:
                h = '<u>%s</u>' % h
            if i:
                h = '<em>%s</em>' % h
            if b:
                h = '<strong>%s</strong>' % h
            toks.append((h, cur_link(), text, b))

        def add_raw(h, plain=''):
            if not in_instr():
                toks.append((h, cur_link(), plain, None))

        def run_fmt(r):
            rpr = r.find(w('rPr'))
            rsid = wattr(rpr.find(w('rStyle')), 'val') if rpr is not None and rpr.find(w('rStyle')) is not None else None

            def get(tag, key):
                el = rpr.find(w(tag)) if rpr is not None else None
                if el is not None:
                    return on(el)
                v = self.dx.style_get(rsid, key) if rsid else None
                return bool(v)
            va = rpr.find(w('vertAlign')) if rpr is not None else None
            va = wattr(va, 'val')
            u = rpr.find(w('u')) if rpr is not None else None
            return (get('b', 'b'), get('i', 'i'), u is not None and wattr(u, 'val') not in (None, 'none') and cur_link() is None,
                    va == 'superscript', va == 'subscript', rpr is not None and on(rpr.find(w('strike'))))

        def do_run(r):
            f = run_fmt(r)
            for ch in r:
                tag = ch.tag
                if tag == w('t'):
                    add_text(ch.text or '', f)
                elif tag == w('tab') or tag == w('ptab'):
                    add_text(' ', f)
                elif tag in (w('br'), w('cr')):
                    if wattr(ch, 'type') not in ('page', 'column'):
                        add_raw('<br>', ' ')
                elif tag == w('noBreakHyphen'):
                    add_text('-', f)
                elif tag == w('sym'):
                    c = wattr(ch, 'char') or ''
                    try:
                        s = chr(int(c, 16))
                    except ValueError:
                        s = ''
                    add_text(SYMBOL_MAP.get(s, s if ord(s or ' ') < 0xF000 else '•'), f)
                elif tag == w('fldChar'):
                    t = wattr(ch, 'fldCharType')
                    if t == 'begin':
                        fields.append({'instr': '', 'phase': 'instr', 'link': None})
                    elif t == 'separate' and fields:
                        fields[-1]['phase'] = 'result'
                        fields[-1]['link'] = self.field_link(fields[-1]['instr'])
                    elif t == 'end' and fields:
                        fields.pop()
                elif tag == w('instrText'):
                    if fields:
                        fields[-1]['instr'] += ch.text or ''
                elif tag in (w('footnoteReference'), w('endnoteReference')):
                    kind = 'footnote' if tag == w('footnoteReference') else 'endnote'
                    nid = wattr(ch, 'id')
                    num = len(self.notes_used) + 1
                    self.notes_used.append((kind, nid, num))
                    add_raw('<sup class="fn"><a href="#%s-%s" id="%sref-%s">%d</a></sup>' % (kind, nid, kind, nid, num))
                elif tag in (w('drawing'), w('pict'), w('object')) or tag == '{%s}AlternateContent' % MC:
                    self.drawing(ch, add_raw, extra_blocks)

        def walk(el):
            for ch in el:
                tag = ch.tag
                if tag == w('r'):
                    do_run(ch)
                elif tag == w('hyperlink'):
                    rid = ch.get('{%s}id' % RNS)
                    anchor = wattr(ch, 'anchor')
                    link = None
                    if anchor:
                        link = '#' + anchor
                    elif rid and rid in self.dx.rels:
                        link = self.dx.rels[rid][0]
                    old = state['link']
                    state['link'] = link
                    walk(ch)
                    state['link'] = old
                elif tag == w('fldSimple'):
                    old = state['link']
                    state['link'] = self.field_link(wattr(ch, 'instr') or '') or old
                    walk(ch)
                    state['link'] = old
                elif tag == w('bookmarkStart'):
                    name = wattr(ch, 'name') or ''
                    if name and name != '_GoBack':
                        add_raw('<span id="%s"></span>' % html.escape(name))
                elif tag in (w('ins'), w('moveTo'), w('smartTag'), w('customXml'), w('sdtContent'), w('sdt')):
                    walk(ch)
                elif tag in (w('del'), w('moveFrom'), w('pPr'), w('rPr'), w('sdtPr'), w('proofErr'),
                             w('commentRangeStart'), w('commentRangeEnd')):
                    continue
                elif tag in ('{%s}oMath' % M, '{%s}oMathPara' % M):
                    txt = ''.join(t.text or '' for t in ch.iter('{%s}t' % M))
                    add_text(txt, (False, True, False, False, False, False))
        walk(p)

        # group consecutive tokens with the same link into <a>
        out, plain = [], []
        i = 0
        while i < len(toks):
            link = toks[i][1]
            j = i
            while j < len(toks) and toks[j][1] == link:
                j += 1
            inner = ''.join(t[0] for t in toks[i:j])
            if link:
                ext = not link.startswith('#')
                inner = re.sub('[\x02\x03\x04]', '', re.sub('\x02[^\x03]*\x03', '', inner))  # no auto-links inside real links
                out.append('<a href="%s"%s>%s</a>' % (html.escape(link), ' target="_blank" rel="noopener"' if ext else '', inner))
            else:
                out.append(inner)
            i = j
        for t in toks:
            plain.append(t[2])
        h = ''.join(out)
        h = re.sub(r'</strong><strong>', '', h)
        h = re.sub(r'</em><em>', '', h)
        text = ''.join(plain)
        txt_toks = [t for t in toks if t[2].strip() and t[3] is not None]
        all_bold = bool(txt_toks) and all(t[3] for t in txt_toks)
        return h.strip(), re.sub(r'\s+', ' ', text).strip(), all_bold

    def field_link(self, instr):
        instr = instr.strip()
        m = re.match(r'HYPERLINK\s+(?:\\l\s+)?"([^"]+)"(?:\s+\\l\s+"([^"]+)")?', instr)
        if m:
            if '\\l' in instr.split('"')[0]:
                return '#' + m.group(1)
            return m.group(1) + ('#' + m.group(2) if m.group(2) else '')
        m = re.match(r'(?:REF|NOTEREF)\s+(\S+)', instr)
        if m:
            return '#' + m.group(1)
        return None

    def drawing(self, el, add_raw, extra_blocks):
        # AlternateContent: use the modern "Choice" only
        if el.tag == '{%s}AlternateContent' % MC:
            ch = el.find('{%s}Choice' % MC)
            if ch is None:
                ch = el.find('{%s}Fallback' % MC)
            if ch is not None:
                for sub in ch:
                    self.drawing(sub, add_raw, extra_blocks)
            return
        # text boxes -> their paragraphs become blocks after this one
        for tb in el.iter(w('txbxContent')):
            extra_blocks.extend(self.blocks(tb))
        rids = [b.get('{%s}embed' % RNS) or b.get('{%s}link' % RNS) for b in el.iter('{%s}blip' % A)]
        rids += [b.get('{%s}id' % RNS) for b in el.iter('{%s}imagedata' % V)]
        alt = ''
        for dp in el.iter('{%s}docPr' % WP):
            alt = dp.get('descr') or dp.get('title') or ''
        width = None
        for ex in el.iter('{%s}extent' % WP):
            try:
                width = int(int(ex.get('cx')) / 9525)
            except (TypeError, ValueError):
                pass
        for rid in rids:
            if not rid:
                continue
            path, data = self.dx.media(rid)
            if not data:
                continue
            ext = os.path.splitext(path)[1].lower()
            mime = IMG_MIME.get(ext)
            if not mime:
                self.skipped_images += 1
                add_raw('<span class="img-missing">[picture not shown: %s format]</span>' % ext.lstrip('.').upper(), '')
                continue
            self.images += 1
            add_raw('<img src="data:%s;base64,%s" alt="%s"%s loading="lazy">'
                    % (mime, base64.b64encode(data).decode('ascii'), html.escape(alt),
                       ' width="%d"' % width if width else ''), alt)

    # ---------------- block content ----------------
    def blocks(self, container):
        """Convert body-level children to a list of block dicts."""
        out = []
        for el in container:
            tag = el.tag
            if tag == w('p'):
                out.extend(self.paragraph(el))
            elif tag == w('tbl'):
                out.append({'type': 'table', 'html': self.table(el)})
            elif tag == w('sdt'):
                pr = el.find(w('sdtPr'))
                gal = pr.find('.//' + w('docPartGallery')) if pr is not None else None
                if gal is not None and 'Table of Contents' in (wattr(gal, 'val') or ''):
                    continue  # Word's own TOC
                c = el.find(w('sdtContent'))
                if c is not None:
                    out.extend(self.blocks(c))
            elif tag in (w('customXml'), w('ins'), w('moveTo')):
                out.extend(self.blocks(el))
        return out

    def paragraph(self, p):
        ppr = p.find(w('pPr'))
        sid = wattr(ppr.find(w('pStyle')), 'val') if ppr is not None and ppr.find(w('pStyle')) is not None else None
        sname = self.dx.style_name(sid).lower()
        if sname.startswith('toc') or sname in ('table of figures', 'index 1'):
            return []
        extra = []
        level = self.heading_level(ppr, sid) if sid or ppr is not None else 0
        num = self.numbering(ppr, sid)
        h, text, all_bold = self.runs_html(p, extra)
        blocks = []
        if not text and '<img' not in h and '<span class="img-missing"' not in h:
            # keep bookmarks that sit on empty paragraphs
            if '<span id=' in h:
                blocks.append({'type': 'raw', 'html': h})
            return blocks + extra
        if sname == 'title':
            blocks.append({'type': 'title', 'text': text, 'html': h})
            return blocks + extra
        if not level and not self.has_heading_styles and self.cfg['fallback_headings'] and all_bold and len(text) <= 160:
            for hp in self.cfg['heading_patterns']:
                if re.match(hp['regex'], text):
                    level = int(hp['level'])
                    break
        if level:
            label = num['label'] if num and not num['bullet'] else ''
            blocks.append({'type': 'heading', 'level': level, 'label': label.strip().rstrip('.') if label else '',
                           'text': text, 'html': h})
            return blocks + extra
        if num and (num['label'] or num['bullet']):
            blocks.append({'type': 'li', 'l': num['ilvl'], 'marker': num['label'], 'html': h})
            return blocks + extra
        ind = ppr.find(w('ind')) if ppr is not None else None
        left = (wattr(ind, 'left') or wattr(ind, 'start')) if ind is not None else self.dx.style_get(sid, 'indLeft')
        try:
            lvl = int(round(int(left or 0) / 720.0))
        except ValueError:
            lvl = 0
        blocks.append({'type': 'p', 'l': max(0, min(lvl, 6)), 'html': h})
        return blocks + extra

    def table(self, tbl):
        rows = []
        for tr in tbl.findall(w('tr')):
            trpr = tr.find(w('trPr'))
            is_hdr = trpr is not None and on(trpr.find(w('tblHeader')))
            cells = []
            col = 0
            for tc in tr.findall(w('tc')) + [x for sdt in tr.findall(w('sdt')) for x in sdt.iter(w('tc'))]:
                tcpr = tc.find(w('tcPr'))
                span = int(wattr(tcpr.find(w('gridSpan')), 'val', '1') or 1) if tcpr is not None and tcpr.find(w('gridSpan')) is not None else 1
                vm = tcpr.find(w('vMerge')) if tcpr is not None else None
                vmerge = None if vm is None else (wattr(vm, 'val') or 'continue')
                inner = self.render(self.blocks(tc), in_cell=True)
                cells.append({'col': col, 'span': span, 'vmerge': vmerge, 'html': inner, 'rowspan': 1})
                col += span
            rows.append({'hdr': is_hdr, 'cells': cells})
        # vertical merges -> rowspan
        for ri, r in enumerate(rows):
            for c in r['cells']:
                if c['vmerge'] == 'restart':
                    k = ri + 1
                    while k < len(rows):
                        nxt = [x for x in rows[k]['cells'] if x['col'] == c['col'] and x['vmerge'] == 'continue']
                        if not nxt:
                            break
                        c['rowspan'] += 1
                        k += 1
        h = ['<div class="tbl"><table>']
        head = [r for r in rows if r['hdr']]
        body = [r for r in rows if not r['hdr']]

        def row_html(r, tagname):
            cs = []
            for c in r['cells']:
                if c['vmerge'] == 'continue':
                    continue
                attrs = ''
                if c['span'] > 1:
                    attrs += ' colspan="%d"' % c['span']
                if c['rowspan'] > 1:
                    attrs += ' rowspan="%d"' % c['rowspan']
                cs.append('<%s%s>%s</%s>' % (tagname, attrs, c['html'], tagname))
            return '<tr>%s</tr>' % ''.join(cs)
        if head:
            h.append('<thead>%s</thead>' % ''.join(row_html(r, 'th') for r in head))
        h.append('<tbody>%s</tbody></table></div>' % ''.join(row_html(r, 'td') for r in body))
        return ''.join(h)

    def render(self, blocks, in_cell=False):
        out = []
        for b in blocks:
            t = b['type']
            if t == 'p':
                if in_cell and b['l'] == 0:
                    out.append('<p>%s</p>' % b['html'])
                else:
                    out.append('<p%s>%s</p>' % (' class="ind" style="--l:%d"' % b['l'] if b['l'] else '', b['html']))
            elif t == 'li':
                out.append('<div class="li" style="--l:%d"><span class="m">%s</span><div class="lt">%s</div></div>'
                           % (b['l'], html.escape(b['marker']), b['html']))
            elif t in ('table', 'raw'):
                out.append(b['html'])
            elif t in ('heading', 'title'):
                out.append('<p><strong>%s</strong></p>' % b['html'])
        return '\n'.join(out)

    def notes_html(self):
        if not self.notes_used:
            return ''
        items = []
        for kind, nid, num in self.notes_used:
            el = self.dx.notes[kind].get(nid)
            if el is None:
                continue
            inner = self.render(self.blocks(el))
            inner = re.sub(r'^<p>', '', inner, count=1)
            items.append('<div class="li" id="%s-%s"><span class="m"><a href="#%sref-%s">%d</a></span><div class="lt">%s</div></div>'
                         % (kind, nid, kind, nid, num, inner))
        return '<section class="notes"><h5 class="minor">Notes</h5>%s</section>' % '\n'.join(items)


# =============================================================================
# Main
# =============================================================================
def convert(path, out_path, cfg, embed=True):
    log('Reading %s ...' % os.path.basename(path))
    dx = Docx(path)
    cv = Converter(dx, cfg)
    body_el = dx.doc.find(w('body'))
    blocks = cv.blocks(body_el)
    if not cv.has_heading_styles:
        log('  note: no Heading styles found; detecting bold numbered lines as headings')

    heads = [b for b in blocks if b['type'] == 'heading']
    title_block = next((b for b in blocks if b['type'] == 'title'), None)
    if not heads:
        log('  warning: no headings found - the sidebar will be empty. Use Heading styles in Word, '
            'or set "heading_patterns" in a config file.')

    # ---- which level is "chapter"
    lvls = sorted(set(b['level'] for b in heads))
    chap = cfg['chapter_level']
    if chap == 'auto':
        chap = lvls[0] if lvls else 1
        if len([b for b in heads if b['level'] == chap]) == 1 and len(lvls) > 1:
            only = next(b for b in heads if b['level'] == chap)
            if title_block is None:
                title_block = {'text': only['text']}
            only['as_title'] = True
            chap = lvls[1]
    chap = int(chap)

    chap_rx = re.compile(cfg['chapter_regex'])
    sec_rx = re.compile(cfg['section_regex'])
    ex_rx = re.compile(cfg['exhibit_regex']) if cfg.get('exhibit_regex') else None
    pr_rx = re.compile(cfg['procedure_regex']) if cfg.get('procedure_regex') else None
    targets = {'code': {}, 'section': {}, 'exhibit': {}}
    nav, used = [], cv.used_ids
    cur = None
    html_out = []
    for b in blocks:
        if b['type'] == 'title':
            continue
        if b['type'] != 'heading' or b.get('as_title'):
            if b['type'] == 'heading':
                continue
            html_out.append(cv.render([b]))
            continue
        text, label, lvl = b['text'], b['label'], b['level']
        # number: from Word numbering, or typed at the start of the heading text
        num = label
        title = text
        m = sec_rx.match(text)
        if not num and m:
            num, title = m.group(1), m.group(2)
        if lvl < chap:
            lvl = chap
        if lvl == chap:
            mc = chap_rx.match(text)
            cnum, ctitle = (mc.group(1), mc.group(2).strip()) if mc else (num, title)
            cid = 'ch' + re.sub(r'\W+', '-', cnum) if cnum else wv.slug(ctitle, used)
            if cid in used and cnum:
                cid = wv.slug('ch-' + ctitle, used)
            used.add(cid)
            cur = {'id': cid, 'num': cnum, 't': wv.smart_title(ctitle or text), 'secs': [], 'exs': [], 'procs': []}
            nav.append(cur)
            html_out.append('<h1 id="%s" class="chapter">%s%s</h1>'
                            % (cid, '<span class="num">Chapter %s</span> ' % html.escape(cnum) if cnum else '',
                               html.escape(wv.smart_title(ctitle or text))))
            continue
        if cur is None:
            cur = {'id': 'front', 'num': '', 't': 'Introduction', 'secs': [], 'exs': [], 'procs': []}
            used.add('front')
            nav.append(cur)
            html_out.append('<span class="anchor" id="front"></span>')
        depth = lvl - chap          # 1 = first level under a chapter
        mex = ex_rx.match(text) if ex_rx else None
        mpr = pr_rx.match(text) if pr_rx else None
        if mex:
            hid = wv.num_id(mex.group(1), used, 'ex')
            targets['exhibit'][mex.group(1)] = hid
            cur['exs'].append({'id': hid, 'num': mex.group(1), 't': wv.smart_title(mex.group(2) or text), 'l': 2})
            html_out.append('<h2 id="%s" class="exhibit"><span class="num">%s</span> %s</h2>'
                            % (hid, html.escape(text[:mex.end(1)]), html.escape(wv.smart_title(mex.group(2)))))
            continue
        if mpr:
            hid = wv.num_id(mpr.group(1), used, '')
            targets['code'][mpr.group(1)] = hid
            cur['procs'].append({'id': hid, 'num': mpr.group(1), 't': mpr.group(2) or text, 'u': False})
            html_out.append('<h2 id="%s" class="proc"><span class="num">%s</span> %s</h2>'
                            % (hid, html.escape(mpr.group(1)), html.escape(mpr.group(2) or text)))
            continue
        hid = wv.num_id(num, used) if num else wv.slug(title, used)
        if num:
            targets['section'][num] = hid
        hl = min(depth + 1, 5)
        # keep bookmarks that were inside the heading so Word cross-references still land here
        bms = ''.join(re.findall(r'<span id="[^"]+"></span>', b['html']))
        if depth <= 3:
            cur['secs'].append({'id': hid, 'num': num, 't': wv.smart_title(title), 'l': max(1, depth)})
        tag = 'h%d' % hl if hl <= 4 else 'h5'
        html_out.append('<%s id="%s"%s>%s%s%s</%s>' % (
            tag, hid, ' class="minor"' if tag == 'h5' else '', bms,
            '<span class="num">%s</span> ' % html.escape(num) if num else '', html.escape(wv.smart_title(title)), tag))
    html_out.append(cv.notes_html())
    body = '\n'.join(x for x in html_out if x)
    body, nlinks = wv.resolve_refs(body, targets)

    props = {}
    core = dx.read('docProps/core.xml')
    if core:
        for el in ET.fromstring(core):
            props[el.tag.split('}')[-1]] = (el.text or '').strip()
    title = (cfg['title'] or (title_block['text'] if title_block else '') or props.get('title')
             or os.path.splitext(os.path.basename(path))[0].replace('_', ' '))
    orig = open(path, 'rb').read() if embed else None
    size = wv.assemble(out_path, nav, body, title, cfg['subtitle'], cfg['org'], cfg['meta'],
                       original_bytes=orig, original_name=os.path.basename(path), original_kind='docx')
    nsec = sum(len(c['secs']) for c in nav)
    log('  %d chapters, %d sections, %d exhibits, %d procedures; %d pictures%s; %d footnotes; %d cross-reference links'
        % (len(nav), nsec, sum(len(c['exs']) for c in nav), sum(len(c['procs']) for c in nav), cv.images,
           (' (%d in formats browsers cannot show, e.g. EMF)' % cv.skipped_images) if cv.skipped_images else '',
           len(cv.notes_used), nlinks))
    log('Wrote %s (%.1f MB)' % (out_path, size / 1e6))


def main():
    ap = argparse.ArgumentParser(description='Convert a Word (.docx) manual into one searchable HTML file.')
    ap.add_argument('docx')
    ap.add_argument('-o', '--out', help='output .html (default: next to the .docx)')
    ap.add_argument('-c', '--config', help='JSON config (see configs/)')
    ap.add_argument('--title')
    ap.add_argument('--subtitle')
    ap.add_argument('--no-embed', action='store_true', help="don't embed the original .docx for download")
    a = ap.parse_args()
    if a.docx.lower().endswith('.doc'):
        sys.exit('This is an old .doc file. Open it in Word and "Save As" .docx first.')
    cfg = dict(DEFAULTS)
    cfg.update(wv.load_config(a.config))
    if a.title:
        cfg['title'] = a.title
    if a.subtitle is not None:
        cfg['subtitle'] = a.subtitle
    out = a.out or os.path.splitext(a.docx)[0] + '.html'
    convert(a.docx, out, cfg, embed=not a.no_embed)


if __name__ == '__main__':
    main()
