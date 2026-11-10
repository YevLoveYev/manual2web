"""
Shared helpers for pdf2web.py and docx2web.py.

Standard library only. Works on Python 3.6+.

Turns converted content into ONE self-contained HTML file:
sidebar table of contents, full-text search, deep links, optional embedded original.
"""
import base64
import html
import json
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, 'viewer_template.html')


def esc(s):
    return html.escape(s, quote=True)


def slug(text, used):
    """Make a unique, readable HTML id from heading text."""
    s = re.sub(r'[^A-Za-z0-9]+', '-', text).strip('-').lower()[:60] or 'section'
    if s[0].isdigit():
        s = 's' + s
    base, n = s, 2
    while s in used:
        s = '%s-%d' % (base, n)
        n += 1
    used.add(s)
    return s


def num_id(num, used, prefix='s'):
    """Id for a numbered heading: 7302.12 -> s7302-12"""
    s = prefix + re.sub(r'[^A-Za-z0-9]+', '-', num).strip('-')
    base, n = s, 2
    while s in used:
        s = '%s-%d' % (base, n)
        n += 1
    used.add(s)
    return s


def smart_title(s):
    """'MANAGING DOCUMENTS' -> 'Managing Documents' (leaves mixed-case text alone)."""
    if s != s.upper() or not re.search(r'[A-Z]', s):
        return s
    small = {'a', 'an', 'and', 'as', 'at', 'but', 'by', 'for', 'from', 'in', 'into', 'of', 'on',
             'or', 'the', 'to', 'with', 'within', 'via', 'per'}
    words = s.lower().split(' ')
    out = []
    for i, w in enumerate(words):
        if i > 0 and w in small:
            out.append(w)
        else:
            out.append(re.sub(r'(^|[/-])([a-z])', lambda m: m.group(1) + m.group(2).upper(), w))
    return ' '.join(out)


SMALL_WORDS = {'a', 'an', 'and', 'as', 'at', 'but', 'by', 'for', 'from', 'in', 'into', 'nor', 'of', 'on',
               'or', 'per', 'the', 'to', 'via', 'vs', 'with'}

# Words always written in capitals. Add your own with "keep_upper": [...] in a config file.
# (Any 2-4 letter ALL-CAPS word inside normal text, like PPE or NCIC, is kept as an acronym anyway.)
ACRONYMS = {'DEA', 'US', 'USA', 'U.S.', 'DOJ', 'FBI', 'ATF', 'DHS', 'CBP', 'ICE', 'HQ', 'LIMS', 'SOP', 'FAQ', 'ID',
            'PPE', 'OSHA', 'GSA', 'OPR', 'SA', 'SAC', 'ASAC', 'TFO', 'DI', 'CS', 'SF', 'LD', 'NIDA', 'NCIC',
            'GS', 'IT', 'PII', 'FOIA', 'CFR', 'USC', 'DC', 'II', 'III', 'IV', 'VI', 'VII', 'VIII', 'IX', 'XI', 'XII',
            'HIPAA', 'NADDIS', 'OCDETF', 'HIDTA', 'SWGDRUG', 'FINCEN', 'INTERPOL', 'EPIC', 'SIU', 'OIG', 'GAO',
            'DOD', 'IRS', 'USMS', 'USCG', 'HSI', 'NFLIS', 'CSA', 'CSOS', 'ARCOS', 'DTO', 'TCO', 'AUSA', 'USAO',
            'OCONUS', 'CONUS', 'TDY', 'PCS', 'SES', 'POC', 'NCIS', 'DIR', 'ROI', 'MOU', 'MOA', 'SOD',
            'LSD', 'DMT', 'DET', 'PCP', 'THC', 'GHB', 'MDMA', 'MDA', 'STP', 'DOM', 'LAAM', 'NIK', 'UV', 'ML', 'CBD'}

# Words with capitals inside them that are correct. Add your own with "keep_case": [...] in a config file.
CAMEL_WORDS = {'SharePoint', 'PowerPoint', 'PowerShell', 'OneDrive', 'OneNote', 'YouTube', 'LinkedIn', 'iPhone',
               'iPad', 'iCloud', 'eBay', 'FedEx', 'WhatsApp', 'JavaScript', 'McDonald', 'MacBook', 'WebEx',
               'FaceTime', 'iOS', 'PayPal', 'GitHub', 'TikTok', 'FirstNet', 'eGuardian', 'LexisNexis', 'WestLaw'}


def configure_case(cfg):
    """Apply "keep_upper" / "keep_case" lists from a config file."""
    ACRONYMS.update(x.upper() for x in (cfg.get('keep_upper') or []))
    CAMEL_WORDS.update(cfg.get('keep_case') or [])


def _cap_first(word):
    """Upper-case the first letter, and the first letter after - / ( or a double quote.
    After an apostrophe only in names like O'Neil / D'Angelo - never in agent's."""
    word = re.sub(r'(^|[/(\-"“])([a-z])', lambda m: m.group(1) + m.group(2).upper(), word)
    return re.sub(r"^(\W*[OoDd]['’])([a-z])", lambda m: m.group(1)[:-2] + m.group(1)[-2].upper() + m.group(1)[-1]
                  + m.group(2).upper(), word)


def _case_core(core, all_caps):
    """Case one word (no spaces, no hyphens). Returns (text, is_special) - special words are never lowercased."""
    if not core:
        return core, True
    poss = ''
    m = re.match(r"^(.*?)(['’][sS])$", core)              # possessive: US's, agent's
    if m and m.group(1):
        core, poss = m.group(1), m.group(2).lower()
    up = core.upper()
    letters = re.sub(r'[^A-Za-z]', '', core)
    if re.search(r'\d', core):                                  # 7302.12, SFL2, AM6311, 12
        return core + poss, True
    if '.' in core.strip('.') or (core.endswith('.') and len(letters) == 1 and False):
        if re.match(r'^(?:[A-Za-z]\.){1,}[A-Za-z]?\.?$', core) and len(letters) <= 4:   # U.S., U.S, u.s., D.C.
            if core.lower() in ('e.g.', 'i.e.', 'e.g', 'i.e'):
                return core.lower() + poss, True
            return up + poss, True
    plain = core.rstrip('.')
    pup = plain.upper()
    if pup in ACRONYMS:                                         # DEA, US, LIMS (also 'us' after all-caps text)
        if all_caps or plain.isupper():
            return pup + core[len(plain):] + poss, True
    if len(pup) > 2 and pup.endswith('S') and pup[:-1] in ACRONYMS and plain[-1] == 's' and plain[:-1].isupper():
        return plain + core[len(plain):] + poss, True           # SOPs, IDs, TFOs
    if not all_caps and re.match(r'^[A-Z]{2,4}s?$', plain):     # other short acronyms in normal text: NCIC
        return core + poss, True
    hit = plain if plain in CAMEL_WORDS else next((c for c in CAMEL_WORDS if c.lower() == plain.lower()), None)
    if hit and not all_caps:                                    # SharePoint, iPhone
        return hit + core[len(plain):] + poss, True
    return core.lower() + poss, False


def toc_case(s, mode='title'):
    """Capitalization for table-of-contents entries and headings.

    title (default): standard Title Case.
        'reviewing approving and posting documents - lab' -> 'Reviewing Approving and Posting Documents - Lab'
        Kept as written: codes with digits (7302.12, SFL2, AM6311, DEA-12), acronyms (DEA, U.S., LIMS, PPE),
        known CamelCase words (SharePoint, iPhone) - also inside hyphenated words (US-Mexico, non-US).
        Fixed: stray capitals (naLOXone, privilEges -> Naloxone, Privileges), a shouted word inside normal text
        (Foreign Counterpart ASSISTANCE -> Assistance), u.s. -> U.S., agent's never becomes Agent'S.
    upper: ALL CAPS.   asis: unchanged."""
    if not s or mode == 'asis':
        return s
    if mode == 'upper':
        return s.upper()
    all_caps = s == s.upper() and bool(re.search(r'[A-Z]{2}', s))
    words = s.split(' ')
    out = []
    for i, wd in enumerate(words):
        lead = re.match(r'^[^\w]*', wd).group(0)
        rest = wd[len(lead):]
        trail = re.search(r'[),;:!?"”\]]*$', rest).group(0)
        body = rest[:len(rest) - len(trail)] if trail else rest
        if not body:
            out.append(wd)
            continue
        # hyphenated / slashed compounds: case each part (US-Mexico, non-US, follow-up, US/Canada)
        parts = re.split(r'([-/–—])', body)
        first_or_last = i == 0 or i == len(words) - 1 or (out and out[-1].endswith((':', '-', '–', '—')))
        cased = []
        for k, p in enumerate(parts):
            if k % 2 == 1 or not p:
                cased.append(p)
                continue
            txt, special = _case_core(p, all_caps)
            if not special:
                is_small = txt.rstrip('.') in SMALL_WORDS and len(parts) == 1 and not first_or_last
                if not is_small:
                    txt = _cap_first(txt)
            cased.append(txt)
        out.append(lead + ''.join(cased) + trail)
    return ' '.join(out)


# ----------------------------------------------------------------------------
# Cross-reference links ("see PRO-7302.2C", "LOM 7101.2", "Exhibit 1/71")
# ----------------------------------------------------------------------------
# Converters wrap candidate references in these sentinels while rendering text,
# then resolve_refs() turns them into links once every target id is known.
REF_OPEN, REF_MID, REF_CLOSE = '\x02', '\x03', '\x04'


class RefLinker(object):
    """rules: list of {"regex": "...(group1 = key)...", "kind": "code"|"section"|"exhibit"}"""

    def __init__(self, rules):
        self.rules = [(re.compile(r['regex']), r.get('kind', 'code')) for r in (rules or [])]

    def mark(self, escaped_text):
        """Wrap references found in already-escaped text with sentinels."""
        if not self.rules:
            return escaped_text
        for rx, kind in self.rules:
            escaped_text = rx.sub(
                lambda m: '%s%s:%s%s%s%s' % (REF_OPEN, kind, m.group(1), REF_MID, m.group(0), REF_CLOSE),
                escaped_text)
        return escaped_text


def resolve_refs(body, targets):
    """targets: {'code': {key: id}, 'section': {num: id}, 'exhibit': {num: id}}
    Returns (html, number_of_links)."""
    count = [0]

    def section_target(num):
        n = re.sub(r'[A-Za-z]+$', '', num)
        while n:
            if n in targets.get('section', {}):
                return targets['section'][n]
            if '.' in n:
                n = n[:-1].rstrip('.')
            else:
                break
        return None

    def rep(m):
        kind, key, text = m.group(1), m.group(2), m.group(3)
        if kind == 'section':
            tgt = section_target(key)
        else:
            tgt = targets.get(kind, {}).get(key)
        if not tgt:
            return text
        count[0] += 1
        return '<a class="xref" href="#%s">%s</a>' % (tgt, text)

    # nested sentinels (two rules matching the same text) -> innermost wins
    pat = re.compile(re.escape(REF_OPEN) + r'(\w+):([^' + REF_MID + ']*)' + re.escape(REF_MID) +
                     r'([^' + REF_OPEN + REF_CLOSE + ']*)' + re.escape(REF_CLOSE))
    prev = None
    while prev != body:
        prev = body
        body = pat.sub(rep, body)
    body = body.replace(REF_OPEN, '').replace(REF_MID, '').replace(REF_CLOSE, '')
    return body, count[0]


# ----------------------------------------------------------------------------
# Final assembly
# ----------------------------------------------------------------------------
def assemble(out_path, nav, content_html, title, subtitle='', org='', meta='',
             original_bytes=None, original_name='', original_kind='none', pages=0, page_offset=0, footer='',
             link_base='', cfg=None):
    """Write the single self-contained HTML file.

    nav: list of chapters:
      {"id","num","t", "secs":[{"id","num","t","l"}], "exs":[...], "procs":[{"id","num","t","u"}]}
    """
    with open(TEMPLATE, encoding='utf-8') as f:
        t = f.read()
    cfg = cfg or {}
    # optional title block at the top of the page ("cover": false removes it)
    if cfg.get('cover', True) is False:
        t = re.sub(r'    <section class="cover" id="top">.*?</section>\n', '    <span class="anchor" id="top"></span>\n', t, count=1, flags=re.S)
    # optional notice that pops up when the page opens ("popup": true = the footer text, or its own text)
    popup = cfg.get('popup')
    popup_text = footer if popup is True else (popup if isinstance(popup, str) else '')
    popup_cfg = {'text': popup_text or '', 'title': cfg.get('popup_title') or 'Notice',
                 'button': cfg.get('popup_button') or 'OK', 'every': cfg.get('popup_frequency') or 'always'}
    mime = {'pdf': 'application/pdf',
            'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}.get(original_kind, '')
    if not original_bytes:
        original_kind = 'none'
    orig = {'kind': original_kind, 'name': original_name, 'mime': mime, 'pages': pages, 'offset': page_offset}
    b64 = base64.b64encode(original_bytes).decode('ascii') if original_bytes else ''
    b64 = '\n'.join(b64[i:i + 1000] for i in range(0, len(b64), 1000))
    js = lambda o: json.dumps(o, separators=(',', ':'), ensure_ascii=False).replace('</', '<\\/')
    doc_title = title + (' · ' + subtitle if subtitle else '')
    reps = {
        '__DOCTITLE__': esc(doc_title),
        '__TITLE__': esc(title),
        '__SUBTITLE__': esc(subtitle),
        '__ORG__': esc(org),
        '__META__': esc(meta) + (' · ' if meta else ''),
        '__FOOTER__': esc(footer or ''),
        '__LINKBASE__': js((link_base or '').split('#')[0]),
        '__BODYCLASS__': 'has-foot' if footer else '',
        '__BODYSTYLE__': footer_size_css(cfg.get('footer_size')),
        '__POPUP__': js(popup_cfg),
        '__ORIG__': js(orig),
        '__NAV__': js(nav),
    }
    for k, v in reps.items():
        t = t.replace(k, v)
    # content and the (large) original go in last so placeholders inside them are never touched
    t = t.replace('__CONTENT__', content_html, 1).replace('__PDF__', b64, 1)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(t)
    return os.path.getsize(out_path)


def footer_size_css(v):
    """14 -> '--foot-size:14px'; also accepts '14px', '11pt', '1.1em'. Anything else is ignored."""
    if v in (None, '', False):
        return ''
    v = str(v).strip()
    if re.match(r'^\d+(\.\d+)?$', v):
        v += 'px'
    return '--foot-size:%s' % v if re.match(r'^\d+(\.\d+)?(px|pt|em|rem)$', v) else ''


def load_config(path):
    if not path:
        return {}
    with open(path, encoding='utf-8') as f:
        return json.load(f)
