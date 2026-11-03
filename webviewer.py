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
