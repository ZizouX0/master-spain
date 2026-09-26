"""Render a visa-dossier markdown note to print-ready HTML (then html_to_pdf.js makes the PDF).

- [placeholders] are highlighted yellow, like in the main pack.
- A line "---" becomes a page break; each "## Email" section starts on a new page.
Usage: python3 md_to_pdf_html.py in.md out.html "Title"
"""
import re
import sys

import markdown

CSS = """
@page { size: A4; margin: 18mm 18mm 20mm 18mm; }
body { font-family: Arial, "Liberation Sans", Helvetica, sans-serif; font-size: 10pt; line-height: 1.4; color: #111; }
h1 { font-size: 18pt; color: #1F4E79; border-bottom: 1.5px solid #1F4E79; padding-bottom: 4pt; margin: 0 0 10pt; }
h2 { font-size: 13pt; color: #1F4E79; margin: 14pt 0 6pt; break-after: avoid; }
h2.email { break-before: page; }
p { margin: 0 0 6pt; }
ol, ul { margin: 0 0 6pt 18pt; padding: 0; }
li { margin: 0 0 2pt; }
table { width: 100%; border-collapse: collapse; margin: 6pt 0 10pt; font-size: 9.2pt; }
th, td { border: 1px solid #BFBFBF; padding: 3pt 5pt; vertical-align: top; text-align: left; }
th { background: #1F4E79; color: #fff; }
tr { break-inside: avoid; }
mark { background: #FFF200; color: inherit; padding: 0 1px; }
em { color: #555; }
.pb { break-after: page; }
.mail { border: 1px solid #C9D6E3; background: #F7FAFC; padding: 8pt 10pt; border-radius: 3px; margin: 6pt 0; }
code { font-family: inherit; background: #FFF200; }
"""


def main(src, out, title):
    md = open(src, encoding="utf-8").read()
    # horizontal rules -> explicit page-break marker
    md = re.sub(r"(?m)^---\s*$", '\n<div class="pb"></div>\n', md)
    html = markdown.markdown(md, extensions=["tables", "sane_lists"])
    # highlight [placeholders] (not inside tags)
    html = re.sub(r"\[([^\]<>]{1,60})\]", r"<mark>[\1]</mark>", html)
    # each Email section on its own page
    html = re.sub(r"<h2>(Email \d)", r'<h2 class="email">\1', html)
    doc = (f'<!doctype html><html lang="es"><head><meta charset="utf-8"><title>{title}</title>'
           f"<style>{CSS}</style></head><body>{html}</body></html>")
    open(out, "w", encoding="utf-8").write(doc)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "Visa dossier")
