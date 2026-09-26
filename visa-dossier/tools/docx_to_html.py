"""Convert the pack's .docx into print-ready HTML (used to produce the PDF with Chromium).

Handles only the features build_pack.js emits: headings, runs (bold, italic,
underline, colour, size, highlight), numbered/bulleted paragraphs, shaded
note paragraphs, tables and page breaks.
"""
import html
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def attr(el, name):
    return None if el is None else el.get(W + name)


def num_formats(z):
    root = ET.fromstring(z.read("word/numbering.xml"))
    abs_fmt = {}
    for a in root.findall(W + "abstractNum"):
        lv = {}
        for l in a.findall(W + "lvl"):
            lv[int(attr(l, "ilvl"))] = (attr(l.find(W + "numFmt"), "val"), attr(l.find(W + "lvlText"), "val"))
        abs_fmt[attr(a, "abstractNumId")] = lv
    return {attr(n, "numId"): abs_fmt[attr(n.find(W + "abstractNumId"), "val")] for n in root.findall(W + "num")}


def run_html(r):
    rpr = r.find(W + "rPr")
    text = "".join(t.text or "" for t in r.findall(W + "t"))
    if not text:
        return ""
    s = html.escape(text)
    style = []
    if rpr is not None:
        if (c := rpr.find(W + "color")) is not None:
            style.append(f"color:#{attr(c, 'val')}")
        if (sz := rpr.find(W + "sz")) is not None:
            style.append(f"font-size:{int(attr(sz, 'val')) / 2}pt")
        if rpr.find(W + "b") is not None:
            s = f"<b>{s}</b>"
        if rpr.find(W + "i") is not None:
            s = f"<i>{s}</i>"
        if rpr.find(W + "u") is not None:
            s = f"<u>{s}</u>"
        if rpr.find(W + "highlight") is not None:
            s = f"<mark>{s}</mark>"
    return f'<span style="{";".join(style)}">{s}</span>' if style else s


def para_html(p, fmts, counters):
    ppr = p.find(W + "pPr")
    if p.find(f".//{W}br[@{W}type='page']") is not None and not "".join(t.text or "" for t in p.iter(W + "t")).strip():
        return '<div class="pb"></div>'
    content = "".join(run_html(r) for r in p.findall(W + "r"))
    tag, cls, style = "p", [], []
    if ppr is not None:
        if (ps := ppr.find(W + "pStyle")) is not None:
            m = re.match(r"Heading(\d)", attr(ps, "val") or "")
            if m:
                tag = f"h{m.group(1)}"
        if (jc := ppr.find(W + "jc")) is not None:
            style.append(f"text-align:{ {'both': 'justify', 'center': 'center', 'right': 'right'}.get(attr(jc, 'val'), 'left')}")
        if (sp := ppr.find(W + "spacing")) is not None:
            if attr(sp, "before"):
                style.append(f"margin-top:{int(attr(sp, 'before')) / 20}pt")
            if attr(sp, "after") and tag == "p":
                style.append(f"margin-bottom:{int(attr(sp, 'after')) / 20}pt")
        if ppr.find(W + "shd") is not None and ppr.find(W + "pBdr") is not None:
            cls.append("note")
        if (npr := ppr.find(W + "numPr")) is not None:
            nid, lvl = attr(npr.find(W + "numId"), "val"), int(attr(npr.find(W + "ilvl"), "val") or 0)
            fmt, txt = fmts[nid][lvl]
            if fmt == "decimal":
                counters[nid] = counters.get(nid, 0) + 1
                marker = f"{counters[nid]}."
            else:
                marker = html.escape(txt)
            cls.append(f"li lvl{lvl}")
            content = f'<span class="mk">{marker}</span><span class="lt">{content}</span>'
    if not content:
        content = "&nbsp;"
    c = f' class="{" ".join(cls)}"' if cls else ""
    s = f' style="{";".join(style)}"' if style else ""
    return f"<{tag}{c}{s}>{content}</{tag}>"


def table_html(tbl, fmts, counters):
    widths = [int(attr(g, "w")) for g in tbl.find(W + "tblGrid").findall(W + "gridCol")]
    total = sum(widths)
    cols = "".join(f'<col style="width:{w / total * 100:.2f}%">' for w in widths)
    rows = []
    for tr in tbl.findall(W + "tr"):
        head = tr.find(f"{W}trPr/{W}tblHeader") is not None
        cells = []
        for tc in tr.findall(W + "tc"):
            inner = "".join(para_html(p, fmts, counters) for p in tc.findall(W + "p"))
            cells.append(f"<{'th' if head else 'td'}>{inner}</{'th' if head else 'td'}>")
        rows.append(f"<tr>{''.join(cells)}</tr>")
    head_rows = [r for r in rows if "<th>" in r]
    body_rows = [r for r in rows if "<th>" not in r]
    return f'<table><colgroup>{cols}</colgroup><thead>{"".join(head_rows)}</thead><tbody>{"".join(body_rows)}</tbody></table>'


CSS = """
@page { size: A4; margin: 20mm 20mm 22mm 20mm; }
* { box-sizing: border-box; }
body { font-family: Arial, "Liberation Sans", Helvetica, sans-serif; font-size: 10pt; line-height: 1.3; color: #111; margin: 0; }
p { margin: 0 0 6pt 0; }
h1 { font-size: 16pt; color: #1F4E79; margin: 6pt 0 10pt; padding-bottom: 4pt; border-bottom: 1.5px solid #1F4E79; break-after: avoid; }
h2 { font-size: 12.5pt; color: #222; margin: 12pt 0 6pt; break-after: avoid; }
h3 { font-size: 11pt; color: #1F4E79; margin: 10pt 0 5pt; break-after: avoid; }
mark { background: #FFF200; color: inherit; padding: 0 1px; }
.pb { break-after: page; }
.li { display: flex; margin: 0 0 3pt 0; }
.li .mk { flex: 0 0 18pt; text-align: left; }
.li.lvl0 { padding-left: 9pt; }
.li.lvl1 { padding-left: 36pt; }
.li .lt { flex: 1; }
.note { background: #EEF3F8; border-left: 3px solid #1F4E79; padding: 5pt 8pt; margin: 3pt 0 8pt; color: #444; font-size: 10pt; }
table { width: 100%; border-collapse: collapse; margin: 4pt 0 8pt; table-layout: fixed; font-size: 9.5pt; }
thead { display: table-header-group; }
tr { break-inside: avoid; }
th, td { border: 1px solid #BFBFBF; padding: 3pt 5pt; vertical-align: top; text-align: left; }
th { background: #1F4E79; color: #fff; }
th p, td p { margin: 0 !important; }
"""


def main(src, out):
    z = zipfile.ZipFile(src)
    fmts = num_formats(z)
    body = ET.fromstring(z.read("word/document.xml")).find(W + "body")
    counters, parts = {}, []
    for el in body:
        if el.tag == W + "p":
            parts.append(para_html(el, fmts, counters))
        elif el.tag == W + "tbl":
            parts.append(table_html(el, fmts, counters))
    doc = f'<!doctype html><html lang="fr"><head><meta charset="utf-8"><title>Spain Study Visa - Application Pack</title><style>{CSS}</style></head><body>{"".join(parts)}</body></html>'
    open(out, "w", encoding="utf-8").write(doc)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
