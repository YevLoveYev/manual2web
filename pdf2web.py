#!/usr/bin/env python3
"""
pdf2web.py - turn a PDF manual into ONE searchable, clickable HTML file.

Requirements: Python 3.6+ and PyMuPDF (pip install pymupdf). Nothing else.

Usage:
    python pdf2web.py manual.pdf -o manual.html
    python pdf2web.py manual.pdf -o manual.html --config configs/lom_2021.json
    python pdf2web.py manual.pdf -o manual.html --title "Safety Manual" --no-embed

How it finds the structure (in this order):
  1. Links on the PDF's own "Table of Contents" pages (most precise; Word-made PDFs have them)
  2. PDF bookmarks (the outline panel in Acrobat)
  3. Bold, numbered headings in the text ("7302.12 Receiving Evidence") as a last resort
Repeated page headers/footers (page numbers, "Date Posted ...") are detected and removed automatically.
"""
import argparse
import collections
import html
import os
import re
import sys

try:
    import pymupdf as fitz          # PyMuPDF >= 1.24
except ImportError:
    try:
        import fitz                 # older PyMuPDF
    except ImportError:
        sys.exit('PyMuPDF is required:  pip install pymupdf   (or install the .whl file offline)')

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import webviewer as wv  # noqa: E402

DEFAULTS = {
    "title": None, "subtitle": "", "org": "", "meta": "",
    "skip_pages": [],                  # 1-based pages to leave out (e.g. [1] for a cover)
    "skip_toc_pages": True,            # leave out "Table of Contents" pages (the sidebar replaces them)
    "page_label_offset": 0,            # printed page number = PDF page - offset
    "chapter_bookmark_level": "auto",  # which bookmark level holds chapters
    "chapter_regex": r"^(?:CHAPTER|Chapter|PART|Part)\s+(\w+)[\s:.-]*(.*)$",
    "sections_from": "auto",           # "toc_links" | "bookmarks" | "headings" | "auto"
    "section_regex": r"^(\d+(?:\.\d+)*)\s*(.*)$",
    "level_rule": "dots",              # "dots": 1.2.3 depth | "lom": 7002 / 7002.1 / 7002.11 | "bookmark"
    "exhibit_regex": None,             # e.g. "^(?:Exhibit\\s*)?(\\d+/\\d+)\\s*(.*)$"
    "procedures": None,                # see configs/lom_2021.json
    "autolink": [],                    # cross-reference rules, see configs/lom_2021.json
    "footer_auto": True,
    "drop_line_regex": [],             # extra lines to delete everywhere
    "section_titles": {},              # fix titles the PDF splits badly: {"7308": "Maintaining ..."}
}

LIST_MARK = re.compile(r'^(\(?(?:[A-Z]|[a-z]|\d{1,3}[a-z]?|[ivxIVX]{1,4})[.)]|\(\w{1,3}\)|[•●▪◦▫■□➢\-–o])\s+(.*)$', re.S)
BULLETS = '•●▪◦▫■□➢-–o'


def log(*a):
    print(*a, file=sys.stderr)


# =============================================================================
# 1. Read the PDF into simple blocks
# =============================================================================
class Block(object):
    __slots__ = ('kind', 'page', 'x0', 'y0', 'x1', 'y1', 'segs', 'text', 'bold', 'nlines', 'table', 'size', 'ylast')

    def __init__(self, **kw):
        for k in self.__slots__:
            setattr(self, k, kw.get(k))


def span_fmt(s):
    f = s['flags']
    font = s.get('font', '')
    bold = bool(f & 16) or 'Bold' in font or 'Black' in font or 'Heavy' in font
    ital = bool(f & 2) or 'Italic' in font or 'Oblique' in font
    sup = bool(f & 1)
    return bold, ital, sup


LINE_MARK = re.compile(r'^(\(?(?:[A-Z]|[a-z]|\d{1,3}[a-z]?|[ivx]{1,4})[.)]|\(\w{1,3}\)|[•●▪◦▫■□➢])\s+\S')


def make_block(lines, pno, r):
    """lines: [y, spans, x1, x0] -> one text Block."""
    segs = []
    for i, (y, spans, _, _) in enumerate(lines):
        if i > 0:
            prev = segs[-1][0] if segs else ''
            sep = '' if prev.endswith('-') and not prev.endswith(' -') else ' '
            segs.append((sep, False, False, False))
        for s in spans:
            bold, ital, sup = span_fmt(s)
            segs.append((s['text'], bold, ital, sup))
    text = ''.join(s[0] for s in segs)
    if not text.strip():
        return None
    nb = sum(len(t.strip()) for t, bo, _, _ in segs if bo)
    nt = sum(len(t.strip()) for t, _, _, _ in segs)
    size = max(s['size'] for _, sp, _, _ in lines for s in sp)
    y0 = lines[0][0]
    return Block(kind='text', page=pno, x0=lines[0][3], y0=y0, x1=max(l[2] for l in lines), y1=r.y1, segs=segs,
                 text=re.sub(r'\s+', ' ', text).strip(), bold=(nt > 0 and nb >= 0.9 * nt),
                 nlines=len(lines), size=size, ylast=lines[-1][0])


def read_page(page, pno):
    """Return a list of Blocks (text or table) in reading order."""
    tables = []
    try:
        tf = page.find_tables()
        tables = [t for t in tf.tables if t.row_count >= 1 and t.col_count >= 2]
    except Exception:
        pass
    trects = [fitz.Rect(t.bbox) for t in tables]
    d = page.get_text('dict', sort=True)
    blocks = []
    for b in d['blocks']:
        if b.get('type') != 0:
            continue
        r = fitz.Rect(b['bbox'])
        if any((r & tr).get_area() > 0.5 * max(r.get_area(), 1) for tr in trects):
            continue
        # split a PyMuPDF block into visual lines; join lines that share a baseline
        lines = []
        for l in b['lines']:
            spans = [s for s in l['spans'] if s['text']]
            if not spans:
                continue
            y = round(l['bbox'][1])
            if lines and abs(lines[-1][0] - y) <= 2 and l['bbox'][0] > lines[-1][2]:
                lines[-1][1].extend([{'text': ' ', 'flags': spans[0]['flags'], 'font': spans[0].get('font', ''), 'size': spans[0]['size']}] + spans)
                lines[-1][2] = l['bbox'][2]
            else:
                lines.append([y, list(spans), l['bbox'][2], l['bbox'][0]])
        if not lines:
            continue
        # a new list item ("A.", "2.", "•") starting on its own line begins a new block
        groups = [[lines[0]]]
        for ln in lines[1:]:
            t0 = ''.join(s['text'] for s in ln[1]).lstrip()
            if LINE_MARK.match(t0) and ln[3] <= groups[-1][0][3] + 40:
                groups.append([ln])
            else:
                groups[-1].append(ln)
        for g in groups:
            blk = make_block(g, pno, r)
            if blk:
                blocks.append(blk)
    words = page.get_text('words') if tables else []
    for t, tr in zip(tables, trects):
        # rebuild each cell's text from the page's words (keeps the spaces table.extract() can lose)
        try:
            rows = []
            for row in t.rows:
                cells = []
                for c in row.cells:
                    if c is None:
                        cells.append(None)
                        continue
                    cr = fitz.Rect(c)
                    ws = [w for w in words if cr.contains(fitz.Point((w[0] + w[2]) / 2, (w[1] + w[3]) / 2))]
                    ws.sort(key=lambda w: (w[5], w[6], w[7]))   # block, line, word order
                    lines, key = [], None
                    for w in ws:
                        if (w[5], w[6]) != key:
                            lines.append([])
                            key = (w[5], w[6])
                        lines[-1].append(w[4])
                    cells.append('\n'.join(' '.join(l) for l in lines))
                rows.append(cells)
        except Exception:
            try:
                rows = t.extract()
            except Exception:
                continue
        hdr = False
        try:
            hdr = (not t.header.external) and bool(t.header.names)
        except Exception:
            pass
        blocks.append(Block(kind='table', page=pno, x0=tr.x0, y0=tr.y0, x1=tr.x1, y1=tr.y1,
                            table=(rows, hdr), text=' '.join(c or '' for row in rows for c in row)))
    blocks.sort(key=lambda b: (round(b.y0 / 3), b.x0))
    # glue a lone list marker ("A.", "1.", "•") to the text block printed right after it on the same line
    out = []
    for b in blocks:
        if (out and b.kind == 'text' and out[-1].kind == 'text' and abs(out[-1].y0 - b.y0) < 4
                and b.x0 > out[-1].x1 - 2 and re.match(r'^(\(?\w{1,4}[.)]|[%s])$' % re.escape(BULLETS), out[-1].text)):
            m = out[-1]
            m.segs = list(m.segs) + [(' ', False, False, False)] + list(b.segs)
            m.text = m.text + ' ' + b.text
            m.x1, m.y1, m.bold, m.nlines = b.x1, max(m.y1, b.y1), b.bold, max(m.nlines, b.nlines)
            continue
        out.append(b)
    return out


def find_repeating_lines(doc, pages_blocks):
    """Headers/footers: text in the top/bottom 9% of the page that repeats on many pages."""
    cnt = collections.Counter()
    for pno, blocks in pages_blocks.items():
        h = doc[pno - 1].rect.height
        seen = set()
        for b in blocks:
            if b.kind == 'text' and (b.y1 < h * 0.09 or b.y0 > h * 0.91):
                seen.add(re.sub(r'\d+', '#', b.text))
        cnt.update(seen)
    n = max(len(pages_blocks), 1)
    return {k for k, v in cnt.items() if v >= max(3, 0.3 * n)}


# =============================================================================
# 2. Fix words that the PDF glued together ("andBermuda" -> "and Bermuda")
# =============================================================================
class WordFixer(object):
    def __init__(self, doc):
        self.vocab, self.raw = set(), set()
        for pg in doc:
            for w in pg.get_text('words'):
                self.raw.add(re.sub(r'[^A-Za-z-]', '', w[4]).lower())
                for tok in re.findall(r'[A-Za-z]+', w[4]):
                    self.vocab.add(tok)
        self.low = {v.lower() for v in self.vocab}
        self.fixed = {}

    def _segment(self, tok):
        n = len(tok)
        best = [None] * (n + 1)
        best[0] = []
        for i in range(1, n + 1):
            for j in range(max(0, i - 25), i):
                if best[j] is None:
                    continue
                piece = tok[j:i]
                if (len(piece) >= 2 or piece in ('a', 'A', 'I')) and piece in self.vocab:
                    cand = best[j] + [piece]
                    if best[i] is None or len(cand) < len(best[i]):
                        best[i] = cand
        return best[n]

    def fix(self, text):
        def rep(m):
            tok = m.group(0)
            if tok in self.vocab or tok.lower() in self.low:
                return tok
            seg = self._segment(tok)
            if seg and 2 <= len(seg) <= 4:
                hy = '-'.join(seg)
                new = hy if hy.lower() in self.raw else ' '.join(seg)
                self.fixed[tok] = new
                return new
            return tok
        return re.sub(r'[A-Za-z]{5,}', rep, text)


# =============================================================================
# 3. Structure: chapters, sections, exhibits, procedures
# =============================================================================
def toc_page_set(doc, toc):
    pages = set()
    for i, t in enumerate(toc):
        if re.match(r'^\s*(table of )?contents\s*$', t[1], re.I):
            end = toc[i + 1][2] if i + 1 < len(toc) else t[2] + 1
            pages.update(range(t[2], max(end, t[2] + 1)))
    if not pages:  # no bookmark: look for a page that says TABLE OF CONTENTS near the top
        for i in range(min(len(doc), 30)):
            if re.search(r'^\s*TABLE OF CONTENTS\s*$', doc[i].get_text()[:400], re.M | re.I):
                pages.add(i + 1)
    return pages


def toc_link_entries(doc, tocpages):
    """[(text, target_page_1based, y)] from internal links on the contents pages."""
    out = []
    for p in sorted(tocpages):
        page = doc[p - 1]
        words = page.get_text('words')
        for l in page.get_links():
            if l.get('page') is None or l['page'] < 0:
                continue
            r = l['from']
            cy = (r.y0 + r.y1) / 2
            ws = [w for w in words if w[1] <= cy <= w[3] and r.y0 <= (w[1] + w[3]) / 2 <= r.y1]
            txt = ' '.join(w[4] for w in sorted(ws, key=lambda w: w[0])).strip()
            to = l.get('to')
            y = to.y if to is not None and to.y > 0 else 0
            if txt:
                out.append((txt, l['page'] + 1, y))
    return out


def level_of(num, rule, bm_level=None):
    if rule == 'bookmark' and bm_level:
        return max(1, min(3, bm_level))
    if rule == 'lom':
        if '.' not in num:
            return 1
        return 2 if len(num.split('.')[1]) == 1 else 3
    return max(1, min(3, num.count('.') + 1))


def build_structure(doc, cfg, tocpages):
    toc = doc.get_toc(simple=False)
    simple = [(t[0], t[1], t[2]) for t in toc]
    chapter_rx = re.compile(cfg['chapter_regex'])

    # ---- chapters from bookmarks
    lvl = cfg['chapter_bookmark_level']
    if lvl == 'auto':
        lvl = None
        for L in (1, 2, 3):
            items = [t for t in simple if t[0] == L]
            if len([t for t in items if chapter_rx.match(t[1].strip())]) >= 2:
                lvl = L
                break
        if lvl is None:
            l1 = [t for t in simple if t[0] == 1]
            lvl = 2 if len(l1) == 1 and any(t[0] == 2 for t in simple) else 1
    chapters = []
    for t in simple:
        if t[0] == lvl and not re.match(r'^\s*(table of )?contents\s*$', t[1], re.I):
            m = chapter_rx.match(t[1].strip())
            num, title = (m.group(1), m.group(2).strip()) if m else ('', t[1].strip())
            chapters.append({'num': num, 'title': wv.smart_title(title or t[1].strip()), 'page': max(t[2], 1)})

    # ---- sections
    src = cfg['sections_from']
    links = toc_link_entries(doc, tocpages) if src in ('auto', 'toc_links') else []
    if src == 'auto':
        src = 'toc_links' if len(links) >= 5 else ('bookmarks' if any(t[0] > lvl for t in simple) else 'headings')
    sec_rx = re.compile(cfg['section_regex'])
    ex_rx = re.compile(cfg['exhibit_regex']) if cfg.get('exhibit_regex') else None
    proc_rx = re.compile(cfg['procedures']['code_regex']) if cfg.get('procedures') else None
    raw = []
    if src == 'toc_links':
        raw = [(txt, pg, y, None) for txt, pg, y in links]
    elif src == 'bookmarks':
        for t in toc:
            if t[0] > lvl:
                to = t[3].get('to') if len(t) > 3 and isinstance(t[3], dict) else None
                y = to.y if to is not None and hasattr(to, 'y') else 0
                if 'TABLE OF CONTENTS' in t[1].upper():
                    continue
                raw.append((t[1].strip(), max(t[2], 1), y, t[0] - lvl))
    entries, seen = [], set()
    for txt, pg, y, bml in raw:
        if proc_rx and proc_rx.match(txt):
            continue  # procedures are handled separately
        m_ex = ex_rx.match(txt) if ex_rx else None
        m = sec_rx.match(txt)
        if m_ex:
            e = {'kind': 'ex', 'num': m_ex.group(1), 'title': m_ex.group(2).strip(), 'level': 1}
        elif m and m.group(1):
            e = {'kind': 'sec', 'num': m.group(1), 'title': m.group(2).strip(), 'level': level_of(m.group(1), cfg['level_rule'], bml)}
        elif src == 'bookmarks':
            e = {'kind': 'sec', 'num': '', 'title': txt, 'level': max(1, min(3, bml or 1))}
        else:
            continue
        key = (e['kind'], e['num'] or e['title'])
        if key in seen:
            continue
        seen.add(key)
        if e['num'] and e['num'] in cfg.get('section_titles', {}):
            e['title'] = cfg['section_titles'][e['num']]
            e['fixed_title'] = True
        e.update(page=pg, y=y or 0, found=False)
        entries.append(e)
    return chapters, entries, src


def build_procedures(doc, cfg, tocpages, chapters):
    """Procedure documents that start with e.g. 'See Also:' and whose title block is missing.
    Titles come from the contents pages; each procedure is matched to its start page."""
    pc = cfg.get('procedures')
    if not pc:
        return {}
    code_rx = re.compile(pc['code_regex'])
    start_rx = re.compile(pc['start_regex'])
    rel_rx = re.compile(pc.get('related_regex') or r'(?!x)x')

    def chap_of(p):
        c = None
        for ch in chapters:
            if p >= ch['page']:
                c = ch['page']
        return c

    titles, order, linked = {}, [], {}
    for p in sorted(tocpages):
        words = doc[p - 1].get_text('words')
        for w in words:
            m = code_rx.match(w[4])
            if m and m.group(0) == w[4].rstrip(',.'):
                cy = (w[1] + w[3]) / 2
                same = [x for x in words if x[1] <= cy <= x[3] and x[0] > w[0]]
                code = m.group(0)
                if code not in titles:
                    titles[code] = (' '.join(x[4] for x in sorted(same, key=lambda x: x[0])), chap_of(p))
                    order.append(code)
    for txt, pg, _ in toc_link_entries(doc, tocpages):
        m = code_rx.match(txt)
        if m:
            linked.setdefault(m.group(0), pg)
    starts = {}
    for i in range(len(doc)):
        t = ' '.join(doc[i].get_text().split())
        if start_rx.match(t):
            starts[i + 1] = t
    assign, used = {}, set()
    for code, p in linked.items():
        if p in starts and code not in starts[p][:400] and p not in assign:
            assign[p] = code
            used.add(code)
    for p, t in sorted(starts.items()):
        if p in assign:
            continue
        refs = rel_rx.findall(t[:300])
        cands = [c for c in order if c not in used and titles[c][1] == chap_of(p) and c not in t[:300]
                 and any(re.sub(r'^\D+', '', c).startswith(r) for r in refs)]
        if cands:
            assign[p] = cands[0]
            used.add(cands[0])
    for p, code in (pc.get('overrides') or {}).items():
        assign[int(p)] = code
        used.add(code)
    redw = pc.get('redacted_word')
    procs = {}
    for p, t in sorted(starts.items()):
        code = assign.get(p)
        if code:
            procs[p] = {'id': code.replace('.', '-'), 'code': code, 'title': titles.get(code, ('', 0))[0] or code}
        else:
            rel = rel_rx.search(t[:200])
            if redw and t.count(redw) > 8:
                title = 'Redacted procedure'
            elif rel:
                title = 'Procedure related to %s' % rel.group(0)
            else:
                title = 'Untitled procedure'
            procs[p] = {'id': 'proc-p%d' % p, 'code': '', 'title': title, 'untitled': True}
    missing = [c for c in order if c not in used]
    if missing:
        log('  note: procedures listed in contents but not located:', ', '.join(missing))
    return procs


# =============================================================================
# 4. Render
# =============================================================================
def render_segs(segs, linker, fixer):
    """Text segments -> inline HTML (bold/italic/superscript), with reference sentinels."""
    out, cur, text = [], None, []

    def flush():
        if not text:
            return
        s = fixer.fix(''.join(text))
        s = linker.mark(html.escape(s, quote=False))
        b, i, sup = cur
        if sup and len(s.strip()) <= 4:
            s = '<sup>%s</sup>' % s
        if i and s.strip():
            s = '<em>%s</em>' % s
        if b and s.strip():
            s = '<strong>%s</strong>' % s
        out.append(s)
    for t, b, i, sup in segs:
        fmt = (b, i, sup)
        if fmt != cur and t.strip():
            flush()
            text, cur = [], fmt
        if cur is None:
            cur = fmt
        text.append(t)
    flush()
    h = ''.join(out)
    h = re.sub(r'</strong>(\s*)<strong>', r'\1', h)
    h = re.sub(r'(<(strong|em)>)(\s+)', r'\3\1', h)        # move spaces outside tags
    h = re.sub(r'(\s+)(</(strong|em)>)', r'\2\1', h)
    return re.sub(r'\s+', ' ', h).strip()


def strip_marker_html(h):
    """Remove the list marker from the start of inline HTML (it may sit inside <strong>)."""
    return re.sub(r'^((?:<[^>]+>)*)\s*(\(?(?:[A-Za-z]|\d{1,3}[a-z]?|[ivxIVX]{1,4})[.)]|\(\w{1,3}\)|[%s])\s+' % re.escape(BULLETS),
                  r'\1', h, count=1)


def render_table(rows, header, linker, fixer):
    def cell(c):
        if not c:
            return ''
        parts = [p.strip() for p in c.split('\n')]
        out = []
        for p in parts:
            if not p:
                continue
            piece = linker.mark(html.escape(fixer.fix(p), quote=False))
            if out and (re.match(r'^[%s]' % re.escape(BULLETS), p) or re.match(r'^\d+[a-z]?\.\s', p)):
                out.append('<br>' + piece)
            else:
                out.append((' ' if out else '') + piece)
        return ''.join(out)
    rows = [r for r in rows if any((c or '').strip() for c in r)]
    if not rows:
        return ''
    h = ['<div class="tbl"><table>']
    start = 0
    if header:
        h.append('<thead><tr>%s</tr></thead>' % ''.join('<th>%s</th>' % cell(c) for c in rows[0]))
        start = 1
    h.append('<tbody>')
    for r in rows[start:]:
        h.append('<tr>%s</tr>' % ''.join('<td>%s</td>' % cell(c) for c in r))
    h.append('</tbody></table></div>')
    return ''.join(h)


def heading(level, hid, label, title, cls=''):
    lab = '<span class="num">%s</span> ' % html.escape(label) if label else ''
    c = ' class="%s"' % cls if cls else ''
    return '<h%d id="%s"%s>%s%s</h%d>' % (level, hid, c, lab, html.escape(title), level)


def convert(pdf_path, out_path, cfg, embed=True):
    doc = fitz.open(pdf_path)
    n = len(doc)
    log('Reading %s (%d pages)...' % (os.path.basename(pdf_path), n))
    tocpages = toc_page_set(doc, doc.get_toc()) if cfg['skip_toc_pages'] else set()
    chapters, entries, src = build_structure(doc, cfg, toc_page_set(doc, doc.get_toc()))
    procs = build_procedures(doc, cfg, toc_page_set(doc, doc.get_toc()), chapters)
    log('  structure from %s: %d chapters, %d sections/exhibits, %d procedures' % (src, len(chapters), len(entries), len(procs)))
    skip = set(cfg['skip_pages']) | tocpages

    pages_blocks = {}
    for p in range(1, n + 1):
        if p not in skip:
            pages_blocks[p] = read_page(doc[p - 1], p)
    furniture = find_repeating_lines(doc, pages_blocks) if cfg['footer_auto'] else set()
    drop_rx = [re.compile(r) for r in cfg['drop_line_regex']]
    fixer = WordFixer(doc)
    linker = wv.RefLinker(cfg['autolink'])

    # left margin of body text (most common x0) -> list indentation levels
    xs = collections.Counter(round(b.x0) for bl in pages_blocks.values() for b in bl if b.kind == 'text')
    left = xs.most_common(1)[0][0] if xs else 72

    used_ids = set()
    for e in entries:
        e['id'] = (wv.num_id(e['num'], used_ids, 'ex' if e['kind'] == 'ex' else 's') if e['num']
                   else wv.slug(e['title'], used_ids))
    by_page = collections.defaultdict(list)
    for e in entries:
        by_page[e['page']].append(e)
    for p in by_page:
        by_page[p].sort(key=lambda e: e['y'])

    chapter_start = {}
    for i, ch in enumerate(chapters):
        p = ch['page']
        while p in skip and p <= n:
            p += 1
        ch['id'] = 'ch' + (re.sub(r'\W+', '-', ch['num']) if ch['num'] else str(i + 1))
        chapter_start.setdefault(p, []).append(ch)

    heading_fallback = (src == 'headings')
    sec_rx = re.compile(cfg['section_regex'])
    out = []           # list of [kind, html]
    prev_block = [None]
    last_text = None   # index of last paragraph/list item, for joining text split across pages
    for p in range(1, n + 1):
        if p in skip:
            continue
        out.append(['pg', '<div class="pg" id="p%d" data-p="%d"><span>Page %s</span></div>'
                    % (p, p, max(p - cfg['page_label_offset'], 0) or 'cover')])
        first_on_page = True
        for ch in chapter_start.get(p, []):
            out.append(['h', heading(1, ch['id'], ('Chapter ' + ch['num']) if ch['num'] else '', ch['title'], 'chapter')])
            ch['first'] = True
            first_on_page = False
        if p in procs:
            pr = procs[p]
            out.append(['h', heading(2, pr['id'], pr['code'], pr['title'], 'proc' + (' untitled' if pr.get('untitled') else ''))])
            first_on_page = False
        pending = list(by_page.get(p, []))
        for b in pages_blocks[p]:
            key = re.sub(r'\d+', '#', b.text)
            if b.kind == 'text' and (key in furniture or any(r.search(b.text) for r in drop_rx)):
                continue
            if b.kind == 'table':
                out.append(['table', render_table(b.table[0], b.table[1], linker, fixer)])
                first_on_page = False
                continue
            # --- is this block one of the known section/exhibit headings?
            hit = None
            if len(b.text) <= 200:
                for e in pending:
                    if e['kind'] == 'ex':
                        if re.match(r'^Exhibit\s*' + re.escape(e['num']) + r'\b', b.text):
                            hit = e
                    elif e['num']:
                        mm = re.match(r'^' + re.escape(e['num']) + r'(?![\d]|\.\d)\s*(.*)$', b.text)
                        if mm:
                            hit = e
                            if mm.group(1) and not e.get('fixed_title'):
                                e['title'] = fixer.fix(mm.group(1).strip())
                    elif b.text.lower().startswith(e['title'].lower()[:40]):
                        hit = e
                    if hit:
                        break
            if hit:
                pending.remove(hit)
                hit['found'] = True
                if hit['kind'] == 'ex':
                    out.append(['h', heading(2, hit['id'], 'Exhibit ' + hit['num'], wv.smart_title(hit['title']), 'exhibit')])
                else:
                    out.append(['h', heading(hit['level'] + 1, hit['id'], hit['num'], hit['title'])])
                first_on_page = False
                continue
            if heading_fallback and b.bold and len(b.text) <= 160:
                mm = sec_rx.match(b.text)
                if mm and mm.group(1) and mm.group(2):
                    e = {'kind': 'sec', 'num': mm.group(1), 'title': mm.group(2), 'level': level_of(mm.group(1), cfg['level_rule']),
                         'page': p, 'y': b.y0, 'found': True}
                    e['id'] = wv.num_id(e['num'], used_ids)
                    entries.append(e)
                    out.append(['h', heading(e['level'] + 1, e['id'], e['num'], e['title'])])
                    continue
            # --- other short bold lines become minor headings
            if b.bold and len(b.text) <= 120 and b.nlines <= 2 and not LIST_MARK.match(b.text):
                out.append(['h5', '<h5 class="minor">%s</h5>' % linker.mark(html.escape(fixer.fix(b.text), quote=False))])
                first_on_page = False
                continue
            inline = render_segs(b.segs, linker, fixer)
            lvl = max(0, int(round((b.x0 - left) / 18.0)))
            m = LIST_MARK.match(b.text)
            # a wrapped line that PyMuPDF put in its own block: glue it to the block right above
            pb = prev_block[0]
            if (not m and pb is not None and pb.page == b.page and last_text is not None and last_text == len(out) - 1
                    and 0 < b.y0 - pb.ylast <= 1.45 * (pb.size or 11) + 1 and b.x0 >= pb.x0 - 2 and not b.bold):
                out[last_text][1] = re.sub(r'(</p>|</div></div>)$', ' ' + inline + r'\1', out[last_text][1])
                pb.ylast = b.ylast
                continue
            prev_block[0] = b
            # text that continues a sentence from the previous page
            if first_on_page and last_text is not None and re.match(r'^[a-z(]', b.text):
                out[last_text][1] = re.sub(r'(</p>|</div></div>)$', ' ' + inline + r'\1', out[last_text][1])
                first_on_page = False
                continue
            first_on_page = False
            if m:
                marker = m.group(1)
                if marker in BULLETS or marker in ('', '', ''):
                    marker = '•'
                body = strip_marker_html(inline)
                out.append(['li', '<div class="li" style="--l:%d"><span class="m">%s</span><div class="lt">%s</div></div>'
                            % (min(lvl, 6), html.escape(marker), body)])
            elif lvl >= 1:
                out.append(['p', '<p class="ind" style="--l:%d">%s</p>' % (min(lvl, 6), inline)])
            else:
                out.append(['p', '<p>%s</p>' % inline])
            last_text = len(out) - 1
        for e in pending:   # not found on the page: anchor at top of page so links still work
            e['found'] = 'page'
            anchor = (heading(2, e['id'], 'Exhibit ' + e['num'], wv.smart_title(e['title']), 'exhibit') if e['kind'] == 'ex'
                      else '<span class="anchor" id="%s"></span>' % e['id'])
            idx = max(i for i, o in enumerate(out) if o[0] == 'pg')
            out.insert(idx + 1, ['h', anchor])
            if last_text is not None and last_text > idx:
                last_text += 1

    body = '\n'.join(o[1] for o in out)
    targets = {'code': {}, 'section': {}, 'exhibit': {}}
    for e in entries:
        if e['num']:
            targets['exhibit' if e['kind'] == 'ex' else 'section'][e['num']] = e['id']
    for pr in procs.values():
        if pr['code']:
            targets['code'][pr['code']] = pr['id']
    body, nlinks = wv.resolve_refs(body, targets)

    # ---- navigation model
    if not chapters:
        chapters = [{'id': 'top', 'num': '', 'title': 'Contents', 'page': 1}]
    nav = []
    for i, ch in enumerate(chapters):
        end = chapters[i + 1]['page'] if i + 1 < len(chapters) else n + 1
        secs, exs, prs = [], [], []
        for e in entries:
            if ch['page'] <= e['page'] < end and not re.match(r'^\W*REDACTED', e['title'] or ''):
                item = {'id': e['id'], 'num': e['num'], 't': wv.smart_title(e['title']) or e['num'], 'l': e['level']}
                (exs if e['kind'] == 'ex' else secs).append(item)
        for p in sorted(procs):
            if ch['page'] <= p < end:
                pr = procs[p]
                prs.append({'id': pr['id'], 'num': pr['code'], 't': pr['title'], 'p': p, 'u': bool(pr.get('untitled'))})
        nav.append({'id': ch['id'], 'num': ch['num'], 't': ch['title'], 'secs': secs, 'exs': exs, 'procs': prs})

    meta = doc.metadata or {}
    title = cfg['title'] or meta.get('title') or os.path.splitext(os.path.basename(pdf_path))[0].replace('_', ' ')
    orig = open(pdf_path, 'rb').read() if embed else None
    size = wv.assemble(out_path, nav, body, title, cfg['subtitle'], cfg['org'], cfg['meta'],
                       original_bytes=orig, original_name=os.path.basename(pdf_path), original_kind='pdf',
                       pages=n, page_offset=cfg['page_label_offset'])
    not_found = [e for e in entries if e['found'] == 'page']
    log('  %d cross-reference links, %d glued words fixed' % (nlinks, len(fixer.fixed)))
    if not_found:
        log('  %d headings not located in the text (anchored to their page instead): %s'
            % (len(not_found), ', '.join((e['num'] or e['title'])[:30] for e in not_found[:10])))
    log('Wrote %s (%.1f MB)' % (out_path, size / 1e6))


def main():
    ap = argparse.ArgumentParser(description='Convert a PDF manual into one searchable HTML file.')
    ap.add_argument('pdf')
    ap.add_argument('-o', '--out', help='output .html (default: next to the PDF)')
    ap.add_argument('-c', '--config', help='JSON config with structure rules (see configs/)')
    ap.add_argument('--title')
    ap.add_argument('--subtitle')
    ap.add_argument('--no-embed', action='store_true', help="don't embed the original PDF (smaller file, no 'Original' button)")
    a = ap.parse_args()
    cfg = dict(DEFAULTS)
    cfg.update(wv.load_config(a.config))
    if a.title:
        cfg['title'] = a.title
    if a.subtitle is not None:
        cfg['subtitle'] = a.subtitle
    out = a.out or os.path.splitext(a.pdf)[0] + '.html'
    convert(a.pdf, out, cfg, embed=not a.no_embed)


if __name__ == '__main__':
    main()
