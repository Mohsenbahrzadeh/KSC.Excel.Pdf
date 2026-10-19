#!/usr/bin/env python3
"""Converts a guide .md into a print-ready HTML file (used for the PDF build).

Usage: python build-guide-html.py <input.md> <output.html> <fa|en>
Requires: pip install markdown
"""
import re
import sys
import markdown
from pathlib import Path

src, out, lang = Path(sys.argv[1]), Path(sys.argv[2]), sys.argv[3]
rtl = lang == 'fa'
text = src.read_text(encoding='utf-8')
text = re.sub(r'^\s*<div dir="rtl">\s*$', '', text, flags=re.M)
text = re.sub(r'^\s*</div>\s*$', '', text, flags=re.M)

FA = str.maketrans('۰۱۲۳۴۵۶۷۸۹', '0123456789')

html = markdown.markdown(text, extensions=['tables', 'fenced_code', 'sane_lists'], output_format='html5')

# number the h2 sections so the in-document table of contents works in the PDF
def h2(m):
    inner = m.group(1)
    n = re.match(r'\s*([0-9۰-۹]+)[).]', inner)
    sid = 's' + n.group(1).translate(FA) if n else 'x' + str(abs(hash(inner)) % 10**6)
    return f'<h2 id="{sid}">{inner}</h2>'
html = re.sub(r'<h2>(.*?)</h2>', h2, html, flags=re.S)

def link(m):
    # internal links start with the section number ("#15-security" / "#۱۵-امنیت"): point them at the numbered h2 ids
    n = re.match(r'#([0-9۰-۹]+)-', m.group(1))
    return f'<a href="#s{n.group(1).translate(FA)}">{m.group(2)}</a>' if n else m.group(0)
html = re.sub(r'<a href="(#[^"]*)">(.*?)</a>', link, html, flags=re.S)

# keep ".NET (8/10)" and similar Latin runs together inside right-to-left text
def isolate(m):
    return '<bdi dir="ltr">' + m.group(0) + '</bdi>'
parts = re.split(r'(<pre.*?</pre>|<code.*?</code>|<[^>]+>)', html, flags=re.S)
html = ''.join(p if (i % 2 == 1) else re.sub(r'\.NET(?: Framework)?(?: \([^)]*\))?', isolate, p) for i, p in enumerate(parts))

css = """
@page { size: A4; margin: 16mm 14mm 18mm 14mm; }
* { box-sizing: border-box; }
body { font-family: Tahoma, 'Segoe UI', Arial, sans-serif; font-size: 10.5pt; line-height: 1.75; color: #1f2933; }
h1 { font-size: 24pt; color: #1a5c38; margin: 0 0 6pt; border-bottom: 3px solid #217346; padding-bottom: 6pt; }
h2 { font-size: 16pt; color: #1a5c38; margin: 22pt 0 8pt; padding-bottom: 3pt; border-bottom: 1px solid #cfd8dc; page-break-after: avoid; }
h3 { font-size: 12.5pt; color: #264653; margin: 16pt 0 5pt; page-break-after: avoid; }
h4 { font-size: 11pt; margin: 12pt 0 4pt; page-break-after: avoid; }
p { margin: 5pt 0; }
a { color: #0b5ed7; text-decoration: none; }
hr { border: 0; border-top: 1px solid #d9dee4; margin: 14pt 0; }
code { font-family: Consolas, 'Cascadia Mono', monospace; font-size: 9pt; background: #eef1f4; padding: 0 3px; border-radius: 3px; direction: ltr; unicode-bidi: embed; }
pre { line-height: 1.4; direction: ltr; text-align: left; background: #f6f8fa; border: 1px solid #e1e4e8; border-radius: 6px; padding: 8pt 10pt; overflow: hidden; white-space: pre-wrap; word-break: break-word; page-break-inside: avoid; margin: 8pt 0; }
pre code { background: none; padding: 0; font-size: 8.6pt; line-height: 1.4; }
table { border-collapse: collapse; width: 100%; margin: 8pt 0; font-size: 9.4pt; page-break-inside: auto; }
tr { page-break-inside: avoid; }
th, td { border: 1px solid #d0d7de; padding: 4pt 7pt; vertical-align: top; }
th { background: #eaf3ee; color: #1a5c38; }
blockquote { margin: 8pt 0; padding: 4pt 12pt; background: #fff8e1; border-inline-start: 4px solid #f0b400; color: #5d4a00; }
blockquote p { margin: 3pt 0; }
img { max-width: 100%; border: 1px solid #d0d7de; border-radius: 6px; margin: 8pt auto; display: block; page-break-inside: avoid; }
ul, ol { padding-inline-start: 22pt; margin: 5pt 0; }
li { margin: 2pt 0; }
"""
doc = f"""<!doctype html><html lang="{lang}" dir="{'rtl' if rtl else 'ltr'}"><head><meta charset="utf-8"><title>KSC.Excel Guide</title><style>{css}</style></head><body>{html}</body></html>"""
out.write_text(doc, encoding='utf-8')
print('wrote', out)
