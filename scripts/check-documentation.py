#!/usr/bin/env python3
"""Check local links and fragments in the actual static documentation output."""
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import urlsplit, unquote
import sys
ROOT = Path(__file__).resolve().parents[1]
DIST = ROOT / 'docs/site/dist'
BASE = '/ofdrw.net/'
class Page(HTMLParser):
    def __init__(self, text):
        super().__init__(); self.links=[]; self.ids=set(); self.feed(text)
    def handle_starttag(self, tag, attrs):
        a=dict(attrs)
        if 'id' in a:self.ids.add(a['id'])
        if tag=='a' and a.get('href'):self.links.append(a['href'])
if not (DIST / 'pagefind/pagefind.js').is_file():
    sys.exit('Missing Pagefind search output')
pages={p:Page(p.read_text()) for p in DIST.rglob('*.html')}; errors=[]
for p,page in pages.items():
    for href in page.links:
        u=urlsplit(href)
        if (u.scheme or u.netloc) and u.netloc != 'whynpc9.github.io':continue
        path=unquote(u.path)
        if path.startswith('/'):
            if not path.startswith(BASE): errors.append(f'{p.relative_to(DIST)}: missing base: {href}');continue
            target=DIST/path[len(BASE):]
        else:target=p.parent/path if path else p
        if target.is_dir():target=target/'index.html'
        target=target.resolve()
        if not target.exists():errors.append(f'{p.relative_to(DIST)}: missing {href}')
        elif u.fragment and target in pages and unquote(u.fragment) not in pages[target].ids:errors.append(f'{p.relative_to(DIST)}: missing fragment {href}')
for e in sorted(set(errors)):print(e)
if errors:sys.exit(1)
print(f'Documentation links passed: {len(pages)} HTML pages')
