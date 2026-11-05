#!/usr/bin/env python3
"""
links2excel.py - list every hyperlink in a set of Word (.docx) files in an Excel workbook,
so you can see which links are good, broken, or risky.

Requirements: Python 3.6+ ONLY. No packages to install (writes a real .xlsx with the standard library).

Usage:
    python links2excel.py "AM chapters"                 -> "AM chapters - links.xlsx"
    python links2excel.py "AM chapters" -o links.xlsx
    python links2excel.py AM6211.docx AM6212.docx -o links.xlsx
    python links2excel.py "AM chapters" --check         also try to open web links (see below)
    python links2excel.py "AM chapters" --include-toc   also list Word table-of-contents links

What it finds (body text, tables, text boxes, page headers/footers, footnotes, endnotes):
  * normal hyperlinks (web, SharePoint, email, network paths, other files)
  * links to places inside the same document (bookmarks / headings)
  * HYPERLINK fields and Word cross-references (REF fields)
  * hyperlinks on pictures

What "Status" means:
  OK               - link inside the document whose target exists; file found next to the document
  BROKEN           - target bookmark missing, Word shows "Error! Reference source not found", file not found,
                     or (with --check) the web server said 404 / not found
  CHECK MANUALLY   - can't be verified from here (SharePoint needs your login, network path not reachable)
  NOT CHECKED      - web link, run with --check to test it
  EMAIL            - mailto: link
About --check: SharePoint and most intranet sites need your Windows login, which this script can't use,
so they come back as "CHECK MANUALLY". Public web links are tested properly.
"""
import argparse
import collections
import datetime
import os
import re
import sys
import zipfile
import posixpath
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape as xesc

try:                                    # Python 3
    from urllib.parse import urlparse, unquote
    import urllib.request as urlreq
    import urllib.error as urlerr
except ImportError:                     # pragma: no cover
    sys.exit('Python 3 is required.')

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import docx2web as dw  # noqa: E402  (reuses the Word reader: styles, header block, relationships)

W = dw.W
A = dw.A
MC = dw.MC
RNS = dw.RNS
w = dw.w
wattr = dw.wattr


def log(*a):
    print(*a, file=sys.stderr)


# =============================================================================
# 1. Reading links out of one Word file
# =============================================================================
PART_LABELS = (('word/document.xml', 'Body'), ('header', 'Page header'), ('footer', 'Page footer'),
               ('footnotes', 'Footnote'), ('endnotes', 'Endnote'), ('comments', 'Comment'))


def part_label(name):
    base = posixpath.basename(name)
    for key, label in PART_LABELS:
        if name == key or base.startswith(key):
            return label
    return base


def text_of(el):
    return re.sub(r'\s+', ' ', ''.join(t.text or '' for t in el.iter(w('t')))).strip()


def parse_field(instr):
    """HYPERLINK "url" \\l "anchor"  /  HYPERLINK \\l "anchor"  /  REF _Ref123 \\h  ->  (kind, url, anchor)"""
    s = instr.strip()
    m = re.match(r'^HYPERLINK\s+(.*)$', s, re.I | re.S)
    if m:
        rest = m.group(1)
        url = anchor = ''
        ma = re.search(r'\\l\s+"([^"]*)"', rest)
        if ma:
            anchor = ma.group(1)
            rest = rest[:ma.start()] + rest[ma.end():]
        mu = re.search(r'"([^"]*)"', rest) or re.search(r'^\s*(\S+)', rest)
        if mu and not mu.group(1).startswith('\\'):
            url = mu.group(1)
        return 'HYPERLINK', url, anchor
    m = re.match(r'^(REF|NOTEREF)\s+(\S+)', s, re.I)
    if m:
        return 'REF', '', m.group(2)
    m = re.match(r'^PAGEREF\s+(\S+)', s, re.I)
    if m:
        return 'PAGEREF', '', m.group(1)
    return None, '', ''


def links_in_file(path, cfg, include_toc=False):
    """Return (rows, info). Each row is a dict describing one link."""
    dx = dw.Docx(path)
    cv = dw.Converter(dx, cfg)
    hdr, _ = dw.find_header_block(dx, cfg) if cfg.get('header_block') else (None, [])
    info = {'number': '', 'title': ''}
    if hdr:
        info = {'number': ('Appendix ' if hdr.get('kind') == 'appendix' else '') + hdr['number'], 'title': hdr['title']}

    parts = [dx.main] + sorted(n for n in dx.names if re.match(
        r'^%s/(header\d*|footer\d*|footnotes|endnotes|comments)\.xml$' % re.escape(dx.base), n))
    # bookmarks anywhere in the document (targets for internal links)
    bookmarks = set()
    trees = {}
    for part in parts:
        data = dx.read(part)
        if not data:
            continue
        root = ET.fromstring(data) if part != dx.main else dx.doc
        trees[part] = root
        for b in root.iter(w('bookmarkStart')):
            bookmarks.add(wattr(b, 'name') or '')

    rows = []
    for part in parts:
        root = trees.get(part)
        if root is None:
            continue
        rels = dx._rels(part) if part != dx.main else dx.rels
        label = part_label(part)
        parent = {c: p for p in root.iter() for c in p}

        # elements to ignore: the duplicate "Fallback" copy of text boxes, and Word's own TOC block
        skip = set()
        for fb in root.iter('{%s}Fallback' % MC):
            skip.update(fb.iter())
        toc_sdts = []
        for sdt in root.iter(w('sdt')):
            gal = sdt.find('.//' + w('docPartGallery'))
            if gal is not None and 'Table of Contents' in (wattr(gal, 'val') or ''):
                toc_sdts.append(sdt)
        toc_set = set()
        for sdt in toc_sdts:
            toc_set.update(sdt.iter())

        def para_of(el):
            while el is not None and el.tag != w('p'):
                el = parent.get(el)
            return el

        def is_toc(el):
            if el in toc_set:
                return True
            p = para_of(el)
            if p is None:
                return False
            ppr = p.find(w('pPr'))
            sid = wattr(ppr.find(w('pStyle')), 'val') if ppr is not None and ppr.find(w('pStyle')) is not None else None
            return dx.style_name(sid).lower().startswith('toc')

        section = ''
        fields = []   # stack: {'instr', 'phase', 'text', 'el'}

        def add(kind_src, text, url, anchor, el):
            if not include_toc and is_toc(el):
                return
            rows.append({'file': os.path.basename(path), 'chapter': info['number'], 'chapter_title': info['title'],
                         'section': section if label == 'Body' else '', 'where': label, 'text': text,
                         'url': url or '', 'anchor': anchor or '', 'source': kind_src, 'bookmarks': bookmarks,
                         'docdir': os.path.dirname(os.path.abspath(path))})

        for el in root.iter():
            if el in skip:
                continue
            tag = el.tag
            if tag == w('p'):
                ppr = el.find(w('pPr'))
                sid = wattr(ppr.find(w('pStyle')), 'val') if ppr is not None and ppr.find(w('pStyle')) is not None else None
                if label == 'Body' and (cv.heading_level(ppr, sid) if (sid or ppr is not None) else 0):
                    t = text_of(el)
                    if t:
                        section = t[:120]
            elif tag == w('hyperlink'):
                rid = el.get('{%s}id' % RNS)
                anchor = wattr(el, 'anchor') or ''
                url = rels[rid][0] if rid and rid in rels else ''
                add('Hyperlink', text_of(el), url, anchor, el)
            elif tag == w('fldSimple'):
                kind, url, anchor = parse_field(wattr(el, 'instr') or '')
                if kind in ('HYPERLINK', 'REF'):
                    add('Field ' + kind, text_of(el), url, anchor, el)
            elif tag == w('fldChar'):
                t = wattr(el, 'fldCharType')
                if t == 'begin':
                    fields.append({'instr': '', 'phase': 'instr', 'text': '', 'el': el})
                elif t == 'separate' and fields:
                    fields[-1]['phase'] = 'result'
                elif t == 'end' and fields:
                    f = fields.pop()
                    kind, url, anchor = parse_field(f['instr'])
                    if kind in ('HYPERLINK', 'REF'):
                        add('Field ' + kind if kind == 'HYPERLINK' else 'Cross-reference',
                            re.sub(r'\s+', ' ', f['text']).strip(), url, anchor, f['el'])
            elif tag == w('instrText'):
                if fields and fields[-1]['phase'] == 'instr':
                    fields[-1]['instr'] += el.text or ''
            elif tag == w('t'):
                for f in fields:
                    if f['phase'] == 'result':
                        f['text'] += el.text or ''
            elif tag == '{%s}hlinkClick' % A:
                rid = el.get('{%s}id' % RNS)
                url = rels[rid][0] if rid and rid in rels else ''
                if url:
                    pd = parent.get(el)
                    alt = (pd.get('descr') or pd.get('name') or '') if pd is not None else ''
                    add('Picture link', alt or '[picture]', url, '', el)
    return rows, info


# =============================================================================
# 2. Classifying and checking
# =============================================================================
SP_RX = re.compile(r'sharepoint\.(com|us|de|cn)|sharepoint-mil\.us|/sites/|/teams/|/_layouts/|/Forms/AllItems\.aspx', re.I)
SHARING_RX = re.compile(r'/:[a-z]:/[rgs]/|[?&]e=[A-Za-z0-9]{6}|/guestaccess\.aspx|/_layouts/15/(guestaccess|share)', re.I)
DOC_EXT = ('.docx', '.doc', '.docm', '.pdf', '.xlsx', '.xls', '.xlsm', '.pptx', '.ppt', '.txt', '.msg', '.vsdx', '.zip')


def sp_site(url):
    """https://agency.sharepoint.com/sites/Policy/Shared Documents/x.docx -> agency.sharepoint.com/sites/Policy"""
    u = urlparse(url)
    m = re.search(r'/(?:sites|teams|personal)/[^/]+', u.path or '', re.I)
    return (u.netloc + (unquote(m.group(0)) if m else '')).lower() if u.netloc else ''


def target_filename(url):
    u = urlparse(url)
    p = unquote(u.path or url).replace('\\', '/')
    base = posixpath.basename(p.rstrip('/'))
    return base if base.lower().endswith(DOC_EXT) else ''


def classify(row, all_files):
    url, anchor, src = row['url'].strip(), row['anchor'], row['source']
    notes = []
    low = url.lower()
    row['site'] = ''
    if not url:
        if src == 'Cross-reference':
            row['type'] = 'Cross-reference (inside document)'
        else:
            row['type'] = 'Inside this document'
        row['target'] = '#' + anchor if anchor else ''
        if re.search(r'Error!\s+(Reference source not found|Bookmark not defined)', row['text'], re.I):
            row['status'] = 'BROKEN'
            notes.append('Word shows "%s"' % row['text'][:60])
        elif not anchor:
            row['status'] = 'BROKEN'
            notes.append('Link has no target')
        elif anchor in row['bookmarks'] or anchor in ('_top',):
            row['status'] = 'OK'
        else:
            row['status'] = 'BROKEN'
            notes.append('Target "%s" no longer exists in this document' % anchor)
        if anchor.startswith('_Toc') and row['status'] == 'OK':
            notes.append('Points to a heading')
        return notes
    row['target'] = url + ('#' + anchor if anchor else '')
    if low.startswith('mailto:'):
        row['type'] = 'Email'
        row['status'] = 'EMAIL'
        return notes
    if low.startswith(('http://', 'https://')):
        if SP_RX.search(url):
            row['site'] = sp_site(url)
            if '-my.sharepoint' in low or '/personal/' in low:
                row['type'] = 'OneDrive (personal)'
                notes.append("Someone's personal OneDrive - breaks if they leave or change sharing")
            elif SHARING_RX.search(url):
                row['type'] = 'SharePoint sharing link'
                notes.append('Sharing link - may expire or only work for certain people')
            else:
                row['type'] = 'SharePoint'
            if '/forms/allitems.aspx' in low or '/_layouts/' in low and 'doc.aspx' not in low:
                notes.append('Points to a library view/page, not a document')
        else:
            row['type'] = 'Web'
        row['status'] = 'NOT CHECKED'
        if low.startswith('http://'):
            notes.append('Uses http (not https)')
        host = urlparse(url).hostname or ''
        if host in ('localhost', '127.0.0.1') or re.match(r'^\d+\.\d+\.\d+\.\d+$', host):
            notes.append('Points to an IP address/this computer')
        if re.search(r'\bsp20(07|10|13)\b|/sites/[^/]+/default\.aspx|\.aspx\?id=', low):
            notes.append('Looks like an old SharePoint address')
    elif low.startswith('file:') or url.startswith('\\\\') or re.match(r'^[a-zA-Z]:[\\/]', url):
        row['type'] = 'Network / file path'
        p = unquote(re.sub(r'^file:/*', '', url, flags=re.I)).replace('/', os.sep) if low.startswith('file:') else url
        if low.startswith('file:') and url[7:9] != '//' and re.match(r'^file:/*[a-zA-Z]:', url):
            p = unquote(re.sub(r'^file:/*', '', url, flags=re.I))
        elif low.startswith('file://') and not re.match(r'^file:/*[a-zA-Z]:', url):
            p = '\\\\' + unquote(url[7:]).lstrip('/').replace('/', '\\')
        row['path_to_check'] = p
        row['status'] = 'CHECK MANUALLY'
        if re.match(r'^[a-zA-Z]:', p):
            notes.append('Drive letter (%s) - may differ on other computers' % p[:2])
    else:
        # relative link: another file next to the document
        row['type'] = 'Other file (relative)'
        p = os.path.normpath(os.path.join(row['docdir'], unquote(url)))
        if os.path.exists(p):
            row['status'] = 'OK'
        else:
            row['status'] = 'BROKEN'
            notes.append('File not found next to the document')
    fname = target_filename(url)
    if fname:
        if fname.lower() in all_files:
            notes.append('Links to %s, which is one of the files in this set' % all_files[fname.lower()])
            if row['type'] in ('SharePoint', 'Web', 'Network / file path', 'Other file (relative)'):
                row['type'] = row['type'] + ' -> chapter file'
        elif row['type'].startswith(('SharePoint', 'Other file')):
            notes.append('Document: %s' % fname)
    # displayed text looks like an address but differs from the real target
    t = row['text'].strip().rstrip('/').lower()
    if re.match(r'^(https?://|www\.|\\\\)', t):
        if t.replace('https://', '').replace('http://', '') not in unquote(low).rstrip('/'):
            notes.append('Text shows a different address than the link goes to')
    if len(url) > 260:
        notes.append('Very long address (%d characters)' % len(url))
    return notes


def check_web(url, timeout=12, cache={}):
    if url in cache:
        return cache[url]

    class NoRedirect(urlreq.HTTPRedirectHandler):
        pass
    req = urlreq.Request(url, method='HEAD', headers={'User-Agent': 'Mozilla/5.0 (link checker)'})
    res = None
    for attempt in ('HEAD', 'GET'):
        try:
            if attempt == 'GET':
                req = urlreq.Request(url, headers={'User-Agent': 'Mozilla/5.0 (link checker)'})
            r = urlreq.urlopen(req, timeout=timeout)
            final = r.geturl()
            if re.search(r'login\.microsoftonline|/_forms/|/adfs/|login\.windows|/sso/|signin', final, re.I):
                res = ('CHECK MANUALLY', 'Asks for a login - open it yourself to check')
            else:
                res = ('OK', 'Opened (%d)' % r.getcode())
            break
        except urlerr.HTTPError as e:
            if e.code in (404, 410):
                res = ('BROKEN', 'Server says not found (%d)' % e.code)
                break
            if e.code in (401, 403) and attempt == 'GET':
                res = ('CHECK MANUALLY', 'Needs a login (%d) - open it yourself to check' % e.code)
                break
            if attempt == 'HEAD':          # some servers refuse quick HEAD requests - retry as a normal page load
                continue
            res = ('CHECK MANUALLY', 'Server answered %d' % e.code)
            break
        except Exception as e:  # DNS, timeout, SSL, proxy...
            res = ('CHECK MANUALLY', 'Could not reach it from this computer (%s)' % type(e).__name__)
            break
    cache[url] = res
    return res


def check_path(p):
    """Network/file path: OK if found; BROKEN only if the share/drive is reachable but the file isn't there."""
    try:
        if os.path.exists(p):
            return 'OK', 'Found'
        m = re.match(r'^(\\\\[^\\]+\\[^\\]+)', p)          # \\server\share
        root = m.group(1) if m else (p[:3] if re.match(r'^[a-zA-Z]:', p) else '')
        if root and os.path.exists(root):
            return 'BROKEN', 'Not found (the share/drive %s is reachable)' % root
        return 'CHECK MANUALLY', 'Share/drive not reachable from this computer'
    except Exception as e:
        return 'CHECK MANUALLY', 'Could not check (%s)' % type(e).__name__


# =============================================================================
# 3. A small .xlsx writer (standard library only)
# =============================================================================
ILLEGAL_XML = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]')
STATUS_STYLE = {'OK': 4, 'BROKEN': 5, 'CHECK MANUALLY': 6, 'NOT CHECKED': 7, 'EMAIL': 7}


def col_letter(i):
    s = ''
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def write_xlsx(path, sheets):
    """sheets: list of dicts {name, headers, rows, widths, link_col (index or None), status_col (index or None)}"""
    styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="4"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/></font><font><b/><sz val="14"/><name val="Calibri"/></font></fonts>
<fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E8C"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFEDEDED"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="9">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

    def cell(ref, v, style=0):
        st = ' s="%d"' % style if style else ''
        if v is None or v == '':
            return '<c r="%s"%s/>' % (ref, st) if style else ''
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return '<c r="%s"%s><v>%s</v></c>' % (ref, st, v)
        s = ILLEGAL_XML.sub('', str(v))[:32000]
        return '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, st, xesc(s))

    z = zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED)
    sheet_entries, rel_entries, ct = [], [], []
    for n, sh in enumerate(sheets, 1):
        headers, rows = sh['headers'], sh['rows']
        title_rows = sh.get('title_rows') or []
        out, links = [], []
        r = 0
        for tr in title_rows:                                   # optional title lines above the table (Summary)
            r += 1
            out.append('<row r="%d">%s</row>' % (r, cell('A%d' % r, tr[0], tr[1] if len(tr) > 1 else 0)))
        if headers:
            r += 1
            hdr_row = r
            out.append('<row r="%d">%s</row>' % (r, ''.join(cell('%s%d' % (col_letter(i), r), h, 1) for i, h in enumerate(headers))))
        for row in rows:
            r += 1
            cs = []
            for i, v in enumerate(row):
                ref = '%s%d' % (col_letter(i), r)
                style = 0
                if sh.get('status_col') == i:
                    style = STATUS_STYLE.get(str(v), 0)
                elif sh.get('link_col') == i and isinstance(v, str) and re.match(r'^(https?://|mailto:|file:)', v, re.I) \
                        and len(v) < 2000:
                    style = 3
                    links.append((ref, v))
                elif sh.get('wrap_cols') and i in sh['wrap_cols']:
                    style = 2
                if isinstance(row, dict):
                    pass
                cs.append(cell(ref, v, style))
            out.append('<row r="%d">%s</row>' % (r, ''.join(cs)))
        ncols = max([len(headers)] + [len(x) for x in rows]) if (headers or rows) else 1
        dim = 'A1:%s%d' % (col_letter(max(ncols - 1, 0)), max(r, 1))
        widths = sh.get('widths') or []
        cols = ''.join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i + 1, i + 1, wd) for i, wd in enumerate(widths))
        freeze = ''
        if headers:
            freeze = ('<sheetViews><sheetView workbookViewId="0"%s><pane ySplit="%d" topLeftCell="A%d" activePane="bottomLeft" '
                      'state="frozen"/><selection pane="bottomLeft" activeCell="A%d" sqref="A%d"/></sheetView></sheetViews>'
                      % (' tabSelected="1"' if n == 1 else '', hdr_row, hdr_row + 1, hdr_row + 1, hdr_row + 1))
        else:
            freeze = '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        afilter = '<autoFilter ref="A%d:%s%d"/>' % (hdr_row, col_letter(len(headers) - 1), max(r, hdr_row)) if headers and rows else ''
        hl_xml, srels = '', []
        if links:
            items = []
            for k, (ref, url) in enumerate(links, 1):
                items.append('<hyperlink ref="%s" r:id="rId%d"/>' % (ref, k))
                srels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
                             'Target="%s" TargetMode="External"/>' % (k, xesc(url.replace(' ', '%20'), {'"': '&quot;'})))
            hl_xml = '<hyperlinks>%s</hyperlinks>' % ''.join(items)
        xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
               '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
               'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
               '<dimension ref="%s"/>%s<sheetFormatPr defaultRowHeight="15"/>%s<sheetData>%s</sheetData>%s%s</worksheet>'
               % (dim, freeze, '<cols>%s</cols>' % cols if cols else '', ''.join(out), afilter, hl_xml))
        z.writestr('xl/worksheets/sheet%d.xml' % n, xml)
        if srels:
            z.writestr('xl/worksheets/_rels/sheet%d.xml.rels' % n,
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s</Relationships>'
                       % ''.join(srels))
        name = re.sub(r'[\[\]\*\?/\\:]', '', sh['name'])[:31]
        sheet_entries.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (xesc(name, {'"': '&quot;'}), n, n))
        rel_entries.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (n, n))
        ct.append('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % n)
    ns = len(sheets)
    rel_entries.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' % (ns + 1))
    z.writestr('[Content_Types].xml',
               '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
               '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
               '<Default Extension="xml" ContentType="application/xml"/>'
               '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
               '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
               '%s</Types>' % ''.join(ct))
    z.writestr('_rels/.rels',
               '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
               '</Relationships>')
    z.writestr('xl/workbook.xml',
               '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
               'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>%s</sheets></workbook>'
               % ''.join(sheet_entries))
    z.writestr('xl/_rels/workbook.xml.rels',
               '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s</Relationships>'
               % ''.join(rel_entries))
    z.writestr('xl/styles.xml', styles)
    z.close()


# =============================================================================
# 4. Main
# =============================================================================
def main():
    ap = argparse.ArgumentParser(description='List every hyperlink in Word (.docx) files in an Excel workbook.')
    ap.add_argument('inputs', nargs='+', help='.docx files and/or folders of .docx files')
    ap.add_argument('-o', '--out', help='output .xlsx (default: "<folder> - links.xlsx")')
    ap.add_argument('-c', '--config', help='same JSON config as docx2web.py (header block settings)')
    ap.add_argument('--check', action='store_true', help='try to open web links and network paths (slow)')
    ap.add_argument('--include-toc', action='store_true', help="also list links in Word's table of contents")
    a = ap.parse_args()
    cfg = dict(dw.DEFAULTS)
    cfg.update(dw.wv.load_config(a.config))
    files = dw.collect_inputs(a.inputs)
    all_files = {os.path.basename(f).lower(): os.path.basename(f) for f in files}

    rows = []
    log('Reading %d Word file(s)...' % len(files))
    for f in sorted(files, key=lambda p: os.path.basename(p).lower()):
        try:
            r, info = links_in_file(f, cfg, a.include_toc)
        except Exception as e:
            log('  could not read %s: %s' % (os.path.basename(f), e))
            continue
        rows.extend(r)
    for row in rows:
        row['notes'] = classify(row, all_files)

    if a.check:
        web = sorted(set(r['url'] for r in rows if r['status'] == 'NOT CHECKED'))
        log('Checking %d web address(es)... (Ctrl+C to stop)' % len(web))
        results = {}
        try:
            for i, u in enumerate(web, 1):
                results[u] = check_web(u)
                if i % 20 == 0:
                    log('  %d / %d' % (i, len(web)))
        except KeyboardInterrupt:
            log('  stopped - the rest stay "NOT CHECKED"')
        paths = {}
        for row in rows:
            if row['url'] in results:
                row['status'], msg = results[row['url']]
                row['notes'].insert(0, msg)
            elif row.get('path_to_check'):
                p = row['path_to_check']
                if p not in paths:
                    paths[p] = check_path(p)
                row['status'], msg = paths[p]
                row['notes'].insert(0, msg)

    # ---- usage counts per target
    uses = collections.Counter(r['target'] for r in rows if r['url'])
    # ---- sheet 1: every link
    headers = ['File', 'Chapter', 'Chapter title', 'Section', 'Where', 'Link text', 'Type', 'Status',
               'Target', 'SharePoint site', 'Times used (all files)', 'Notes']
    link_rows = []
    for r in rows:
        link_rows.append([r['file'], r['chapter'], r['chapter_title'], r['section'], r['where'], r['text'], r['type'],
                          r['status'], r['target'], r['site'], uses[r['target']] if r['url'] else '', '; '.join(r['notes'])])
    # ---- sheet 2: unique targets
    by_target = collections.OrderedDict()
    for r in rows:
        if not r['url']:
            continue
        t = by_target.setdefault(r['target'], {'type': r['type'], 'status': r['status'], 'site': r['site'], 'files': [],
                                               'texts': collections.Counter(), 'notes': r['notes']})
        if r['file'] not in t['files']:
            t['files'].append(r['file'])
        t['texts'][r['text']] += 1
    uniq_rows = []
    for tgt, t in sorted(by_target.items(), key=lambda kv: (-uses[kv[0]], kv[0].lower())):
        uniq_rows.append([tgt, t['type'], t['status'], uses[tgt], len(t['files']), ', '.join(t['files'][:30]) +
                          (' ...' if len(t['files']) > 30 else ''), ' | '.join(x for x, _ in t['texts'].most_common(3)),
                          t['site'], '; '.join(t['notes'])])
    # ---- sheet 3: summary
    def table(title, counter, label):
        out = [['', ''], [title, ''], [label, 'Links']]
        for k, v in counter.most_common():
            out.append([k or '(none)', v])
        return out
    st = collections.Counter(r['status'] for r in rows)
    ty = collections.Counter(r['type'] for r in rows)
    sites = collections.Counter(r['site'] for r in rows if r['site'])
    per_file = collections.Counter(r['file'] for r in rows)
    broken_file = collections.Counter(r['file'] for r in rows if r['status'] == 'BROKEN')
    summary = []
    for block in (table('By status', st, 'Status'), table('By type', ty, 'Type'),
                  table('SharePoint sites linked to', sites, 'Site') if sites else [],
                  table('Files with the most broken links', broken_file, 'File') if broken_file else []):
        summary.extend(block)
    files_without = sorted(set(all_files.values()) - set(per_file))
    title_rows = [('Hyperlink report', 8), ('Created %s from %d Word file(s): %d links, %d different targets, %d broken'
                                            % (datetime.date.today().isoformat(), len(files), len(rows), len(by_target),
                                               st.get('BROKEN', 0)), 0)]
    if files_without:
        summary.extend([['', ''], ['Files with no links', len(files_without)]] + [[f, ''] for f in files_without[:200]])

    out = a.out
    if not out:
        base = a.inputs[0].rstrip('/\\')
        base = base if os.path.isdir(base) else os.path.splitext(base)[0]
        out = base + ' - links.xlsx'
    write_xlsx(out, [
        {'name': 'Links', 'headers': headers, 'rows': link_rows, 'link_col': 8, 'status_col': 7,
         'widths': [16, 9, 30, 30, 11, 38, 26, 15, 60, 30, 10, 60]},
        {'name': 'Unique targets', 'headers': ['Target', 'Type', 'Status', 'Times used', 'Files', 'Used in', 'Link text (most common)',
                                               'SharePoint site', 'Notes'],
         'rows': uniq_rows, 'link_col': 0, 'status_col': 2, 'widths': [70, 26, 15, 10, 8, 40, 40, 30, 60]},
        {'name': 'Summary', 'headers': [], 'rows': summary, 'widths': [55, 12], 'title_rows': title_rows},
    ])
    log('  %d links (%d different targets) in %d files' % (len(rows), len(by_target), len(files)))
    for k, v in st.most_common():
        log('    %-15s %d' % (k, v))
    log('Wrote %s' % out)


if __name__ == '__main__':
    main()
