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
            'OCONUS', 'CONUS', 'TDY', 'PCS', 'SES', 'POC', 'NCIS', 'DIR', 'ROI', 'MOU', 'MOA', 'SOD'}

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


def toc_case(s, mode='title'):
    """Capitalization for table-of-contents entries and headings.

    title (default): standard Title Case.
        'reviewing approving and posting documents - lab' -> 'Reviewing Approving and Posting Documents - Lab'
        Kept as written: codes with digits (7302.12, SFL2, AM6311, DEA-12), acronyms (DEA, U.S., LIMS, PPE),
        known CamelCase words (SharePoint, iPhone).
        Fixed: stray capitals (naLOXone, privilEges -> Naloxone, Privileges), a shouted word inside normal text
        (Foreign Counterpart ASSISTANCE -> Assistance), U.S. stays U.S., agent's never becomes Agent'S.
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
        trail = re.search(r'[^\w.]*$', wd).group(0)
        core = wd[len(lead):len(wd) - len(trail)] if trail else wd[len(lead):]
        if not re.match(r'^(?:[A-Za-z]\.){2,}', core):
            core = core.rstrip('.')
        if not core:
            out.append(wd)
            continue
        up = core.upper()
        if re.search(r'\d', core):                                             # 7302.12, SFL2, AM6311, DEA-12
            out.append(wd)
            continue
        if re.match(r'^(?:[A-Za-z]\.){2,}[A-Za-z]?\.?$', core):                # U.S., D.C., e.g.
            out.append(wd.replace(core, up) if (all_caps or core[0].isupper()) else wd)
            continue
        if up in ACRONYMS or (up.endswith('S') and up[:-1] in ACRONYMS and len(up) > 2 and core[-1] == 's'):
            if all_caps or core[:-1].isupper() or core.isupper():                # DEA, LIMS, SOPs, IDs
                out.append(wd.replace(core, up if core.isupper() or all_caps else core))
                continue
        if not all_caps and re.match(r'^[A-Z]{2,4}s?$', core):                  # other short acronyms: NCIC, TFOs
            out.append(wd)
            continue
        hit = core if core in CAMEL_WORDS else next((c for c in CAMEL_WORDS if c.lower() == core.lower()), None)
        if hit and not all_caps:                                               # SharePoint, iPhone (also fixes 'sharepoint')
            out.append(wd.replace(core, hit))
            continue
        lw = wd.lower()                                                        # everything else: normal word
        first_or_last = i == 0 or i == len(words) - 1 or (out and out[-1].endswith((':', '-', '–', '—')))
        if core.lower() in SMALL_WORDS and not first_or_last:
            out.append(lw)
            continue
        out.append(_cap_first(lw))
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
             original_bytes=None, original_name='', original_kind='none', pages=0, page_offset=0):
    """Write the single self-contained HTML file.

    nav: list of chapters:
      {"id","num","t", "secs":[{"id","num","t","l"}], "exs":[...], "procs":[{"id","num","t","u"}]}
    """
    with open(TEMPLATE, encoding='utf-8') as f:
        t = f.read()
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


def load_config(path):
    if not path:
        return {}
    with open(path, encoding='utf-8') as f:
        return json.load(f)
