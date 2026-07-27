#!/usr/bin/env python3
"""Convert the main thesis Word file into LaTeX for the muthesis2026 template.

Produces, from a single .docx:

    chapters.tex          the numbered chapters (CHAPTER I .. REFERENCES)
    references.tex        a thebibliography built from the IEEE reference list
    abstract.tex          the abstract body (for \\abstract{...})
    acknowledgements.tex  the acknowledgements body (for \\acknowledgements{...})
    figures/figX-Y.png    every embedded figure, named by its figure number

and prints the front-matter metadata (title, author, advisors, keywords, ...)
so it can be copied into preamble.tex.

Front matter, the embedded table of contents, the appendices, and the list of
abbreviations are skipped: the class regenerates the first two, the appendices
are converted separately by docx2appendix.py, and citations ([N]) are turned
into \\cite{refN} by the shared render_text() so the numbers match references.tex.

Usage:
    python3 tools/docx2thesis.py thesis.docx

Requires: python-docx
"""
import os
import re
import shutil
import sys
import zipfile

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx2appendix import escape, render_text, render_table, table_to_rows  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
FIGDIR = os.path.join(ROOT, "figures")

SMALL = {"a", "an", "and", "or", "the", "of", "for", "to", "in", "on", "with",
         "at", "by", "from", "as", "but", "nor"}


def titlecase(s):
    words = s.split()
    out = []
    for i, w in enumerate(words):
        lw = w.lower()
        if i != 0 and lw in SMALL:
            out.append(lw)
        else:
            out.append(w[:1].upper() + w[1:].lower() if w.isupper() else w)
    return " ".join(out)


def blocks(doc):
    for c in doc.element.body.iterchildren():
        if c.tag == qn("w:p"):
            yield Paragraph(c, doc)
        elif c.tag == qn("w:tbl"):
            yield Table(c, doc)


def img_rid(p):
    m = re.search(r'r:embed="([^"]+)"', p._p.xml)
    return m.group(1) if m else None


FIG_CAP_RE = re.compile(r"^Figure\s+([0-9A-Z]+\.\d+)\s+([A-Z].*)$", re.S)
TAB_CAP_RE = re.compile(r"^Table\s+([0-9A-Z]+\.\d+)\s+(.*\S)\s*$", re.S)
H2_RE = re.compile(r"^\d+\.\d+\s*(.*)$", re.S)
H3_RE = re.compile(r"^\d+\.\d+\.\d+\s*(.*)$", re.S)
# Split a numbered heading into its number ("4.5.1") and its title.
NUM_TITLE_RE = re.compile(r"^(\d+(?:\.\d+)+)\s+(.*\S)\s*$", re.S)


def extract_figures(docx_path):
    """Copy embedded images to figures/figX-Y.<ext>; return rid -> filename."""
    z = zipfile.ZipFile(docx_path)
    rels = z.read("word/_rels/document.xml.rels").decode()
    rid2tgt = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    doc = Document(docx_path)
    allb = list(blocks(doc))
    os.makedirs(FIGDIR, exist_ok=True)
    rid2file = {}
    for i, b in enumerate(allb):
        if not isinstance(b, Paragraph):
            continue
        rid = img_rid(b)
        if not rid:
            continue
        num = _nearest_fig_number(allb, i)
        tgt = rid2tgt.get(rid, "")
        ext = os.path.splitext(tgt)[1] or ".png"
        name = ("fig" + num.replace(".", "-")) if num else ("fig-" + rid)
        fn = name + ext
        src = "word/" + tgt.replace("../", "")
        with z.open(src) as f, open(os.path.join(FIGDIR, fn), "wb") as o:
            shutil.copyfileobj(f, o)
        rid2file[rid] = fn
    return rid2file


def _nearest_fig_number(allb, i):
    """Figure number for the image at block i, from the nearest caption.

    The caption is often in the *same* paragraph as the image, so block i
    itself is checked first, then the following and preceding blocks.
    """
    for b in [allb[i]] + allb[i + 1:i + 9] + list(reversed(allb[max(0, i - 5):i])):
        if isinstance(b, Paragraph):
            m = FIG_CAP_RE.match(b.text.strip())
            if m:
                return m.group(1)
    return None


def is_caption_para(b):
    if not isinstance(b, Paragraph):
        return None
    t = b.text.strip()
    if FIG_CAP_RE.match(t):
        return ("fig", FIG_CAP_RE.match(t))
    if TAB_CAP_RE.match(t):
        return ("tab", TAB_CAP_RE.match(t))
    return None


def convert(docx_path):
    doc = Document(docx_path)
    allb = list(blocks(doc))
    rid2file = extract_figures(docx_path)

    # Region boundaries.
    def find(pred):
        return next((i for i, b in enumerate(allb) if pred(b)), None)

    def h1_is(b, name):
        return (isinstance(b, Paragraph) and b.style.name == "Heading 1"
                and b.text.strip().upper() == name)

    i_ack = find(lambda b: h1_is(b, "ACKNOWLEDGMENTS"))
    i_abs = find(lambda b: h1_is(b, "ABSTRACT"))
    i_ch1 = find(lambda b: isinstance(b, Paragraph)
                 and b.style.name == "Heading 1"
                 and re.match(r"^CHAPTER\s+[IVX]+$", b.text.strip()))
    i_ref = find(lambda b: h1_is(b, "REFERENCES"))

    meta = _metadata(allb, i_ack)
    ack = _collect_bodytext(allb, i_ack + 1, i_abs, meta.get("title", ""))
    abstract, keywords = _collect_abstract(allb, i_abs + 1, i_ch1)
    chapters = _emit_chapters(allb, i_ch1, i_ref, rid2file)
    refs = _emit_references(allb, i_ref)
    abbrev = _collect_abbreviations(allb)

    _write("acknowledgements.tex", ack)
    _write("abstract.tex", abstract)
    _write("chapters.tex", chapters)
    _write("references.tex", refs)
    if abbrev:
        _write("abbreviations.tex", abbrev)

    print("=== METADATA (copy into preamble.tex / main-*.tex) ===")
    for k, v in meta.items():
        print(f"  {k}: {v}")
    print(f"  keywords: {keywords}")
    print("Wrote chapters.tex, references.tex, abstract.tex, acknowledgements.tex")
    print(f"Extracted {len(rid2file)} figures into figures/")


def _write(name, text):
    with open(os.path.join(ROOT, name), "w", encoding="utf-8") as f:
        f.write(text.rstrip() + "\n")


# Lines that open the Word abstract page's header block (title, candidate,
# degree, committee). The class regenerates all of that from preamble.tex, so
# collection must stop there rather than copy it into the acknowledgements.
_ABS_HEADER_RE = re.compile(
    r"^(MR|MRS|MISS|MS)\.\s|ADVISORY COMMITTEE|^M\.(Sc|Eng|A)\.|^Ph\.D\.",
    re.I)

# Headings of the Word document's own front-matter lists. Everything from here
# on is a table of contents / list of tables / list of figures that LaTeX
# regenerates, so it must never be copied into the abstract.
_FRONT_LIST_RE = re.compile(
    r"^(CONTENTS|TABLE OF CONTENTS|LIST OF (TABLES|FIGURES|ABBREVIATIONS))$",
    re.I)


def _is_toc_style(b):
    return isinstance(b, Paragraph) and b.style.name.lower().startswith("toc")


def _letters(s):
    return re.sub(r"[^A-Za-z0-9]", "", s).upper()


def _collect_bodytext(allb, start, stop, title=""):
    """Acknowledgements body, stopping before the abstract-page header block."""
    # Compare on letters only and by prefix: the cover title and the
    # abstract-page title often differ in punctuation or a trailing clause
    # (e.g. the cover omits "& Prevention").
    stem = _letters(title)[:25]
    out = []
    for b in allb[start:stop]:
        if isinstance(b, Table) or _is_toc_style(b):
            continue
        t = b.text.strip()
        if not t or b.style.name.startswith("Heading"):
            continue
        if (stem and _letters(t).startswith(stem)) or _ABS_HEADER_RE.search(t):
            break
        out.append(render_text(t))
    return "\n\n".join(out)


def _collect_abstract(allb, start, stop):
    """Abstract body, stopping at the Word document's own contents listing."""
    paras, keywords = [], ""
    for b in allb[start:stop]:
        if isinstance(b, Table) or _is_toc_style(b):
            continue
        t = b.text.strip()
        if not t or b.style.name.startswith("Heading"):
            continue
        if _FRONT_LIST_RE.match(t):
            break
        if t.lower().startswith("keywords"):
            keywords = t.split(":", 1)[1].strip() if ":" in t else ""
            continue
        if re.match(r"^\d+\s+pages$", t, re.I):
            continue
        paras.append(render_text(t))
    return "\n\n".join(paras), keywords


def _collect_abbreviations(allb):
    """Build the \\listofabbreviations tabbing body from the Word table."""
    for i, b in enumerate(allb):
        if (isinstance(b, Paragraph)
                and b.text.strip().upper() == "LIST OF ABBREVIATIONS"):
            for b2 in allb[i + 1:i + 6]:
                if isinstance(b2, Table):
                    lines = [r"\begin{tabbing}", r"\hspace{3cm}\=\kill"]
                    for r in b2.rows:
                        cells = [c.text.strip() for c in r.cells]
                        if len(cells) < 2 or not cells[0]:
                            continue
                        lines.append("%s \\> %s \\\\"
                                     % (escape(cells[0]), escape(cells[1])))
                    lines.append(r"\end{tabbing}")
                    return "\n".join(lines)
    return ""


def _metadata(allb, i_ack):
    """Best-effort front-matter metadata from the blocks before ACKNOWLEDGMENTS."""
    texts = [b.text.strip() for b in allb[:i_ack + 40]
             if isinstance(b, Paragraph) and b.text.strip()]
    joined = "\n".join(texts)
    m = {}
    m["title"] = re.sub(r"\s+", " ", texts[0]) if texts else ""
    for t in texts:
        if re.search(r"\b\d{7}\b", t):
            mm = re.search(r"([A-Z][A-Za-z .]+?)\s+(\d{7})\s*([A-Z/]+)?", t)
            if mm:
                m.setdefault("candidate_line", t)
        if "M.Sc" in t or "M.Eng" in t or "Ph.D" in t.split(":")[0]:
            m.setdefault("degree_line", t)
        if "ADVISORY COMMITTEE" in t.upper():
            m["committee_line"] = t
        if re.match(r"^(MR\.|MISS|MS\.|MRS\.)", t.upper()):
            m.setdefault("author_line", t)
    return m


def _emit_chapters(allb, start, stop, rid2file):
    out = []
    pending_tab_cap = None            # (num, text) captured from a Table caption
    skip_caps = set()                 # block ids consumed as fig/tab captions
    list_buffer = []
    chap_no = [0]                     # running chapter number for \label{chap:N}

    def flush_list():
        if list_buffer:
            out.append(r"\begin{itemize}")
            out.extend(r"\item " + x for x in list_buffer)
            out.append(r"\end{itemize}")
            out.append("")
            list_buffer.clear()

    # Pre-scan: mark table-caption paragraphs and figure-caption paragraphs.
    tab_cap_for = {}   # table-block-index -> (num, text)
    fig_cap_text = {}  # fig-number -> caption text (without prefix)
    for i, b in enumerate(allb[start:stop], start):
        kind = is_caption_para(b)
        if not kind:
            continue
        typ, mm = kind
        if typ == "tab":
            # attach to the next table block
            for j in range(i + 1, min(i + 5, stop)):
                if isinstance(allb[j], Table):
                    tab_cap_for[j] = (mm.group(1), mm.group(2))
                    skip_caps.add(i)
                    break
        else:
            fig_cap_text[mm.group(1)] = mm.group(2)
            # Skip a stand-alone caption paragraph, but NOT one that also holds
            # the image (that block must still be emitted as the figure).
            if not img_rid(b):
                skip_caps.add(i)

    i = start
    while i < stop:
        b = allb[i]
        if i in skip_caps:
            i += 1
            continue

        if isinstance(b, Table):
            flush_list()
            num, cap = tab_cap_for.get(i, (None, f"Table"))
            label = "tab:" + (num.replace(".", "-") if num else str(i))
            out.append(render_table(table_to_rows(b), cap, label))
            out.append("")
            i += 1
            continue

        style = b.style.name
        t = b.text.strip()
        rid = img_rid(b)

        if rid:  # figure
            flush_list()
            num = _nearest_fig_number(allb, i)
            fn = rid2file.get(rid)
            cap = fig_cap_text.get(num, "")
            if fn:
                out.append(r"\begin{figure}[htbp]")
                out.append(r"\centering")
                out.append(r"\includegraphics[width=\linewidth]{%s}" % fn)
                if cap:
                    lab = "fig:" + (num.replace(".", "-") if num else str(i))
                    out.append(r"\caption{%s}\label{%s}" % (render_text(cap), lab))
                out.append(r"\end{figure}")
                out.append("")
            i += 1
            continue

        if not t:
            i += 1
            continue

        if style == "Heading 1":
            flush_list()
            if re.match(r"^CHAPTER\s+[IVX]+$", t):
                chap_no[0] += 1
                # the next Heading 1 is the chapter title
                for j in range(i + 1, stop):
                    tj = allb[j].text.strip() if isinstance(allb[j], Paragraph) else ""
                    if tj:
                        out.append(r"\chapter{%s}\label{chap:%d}"
                                   % (escape(titlecase(tj)), chap_no[0]))
                        out.append("")
                        skip_caps.add(j)
                        break
            i += 1
            continue

        if style in ("Heading 2", "Heading 3"):
            flush_list()
            m = NUM_TITLE_RE.match(t)
            cmd = "section" if style == "Heading 2" else "subsection"
            if m:
                out.append(r"\%s{%s}\label{sec:%s}"
                           % (cmd, escape(m.group(2)), m.group(1)))
            else:
                out.append(r"\%s{%s}" % (cmd, escape(t)))
            out.append("")
            i += 1
            continue

        if style == "List Paragraph":
            list_buffer.append(render_text(t))
            i += 1
            continue

        # ordinary paragraph
        flush_list()
        out.append(render_text(t))
        out.append("")
        i += 1

    flush_list()
    return "\n".join(out)


def _emit_references(allb, i_ref):
    entries = {}
    for b in allb[i_ref + 1:]:
        if isinstance(b, Table):
            continue
        t = re.sub(r"\s+", " ", b.text.strip())
        if not t:
            continue
        if isinstance(b, Paragraph) and b.style.name.startswith("Heading"):
            break  # reached Appendix A / next section
        m = re.match(r"^\[(\d+)\]\s*(.*\S)\s*$", t)
        if m:
            entries[int(m.group(1))] = m.group(2)
    if not entries:
        return ""
    lines = [r"% Generated by tools/docx2thesis.py from the Word reference list.",
             r"\begin{thebibliography}{999}"]
    for n in sorted(entries):
        lines.append(r"\bibitem{ref%d} %s" % (n, escape(entries[n])))
    lines.append(r"\end{thebibliography}")
    return "\n".join(lines)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("usage: docx2thesis.py thesis.docx")
    convert(sys.argv[1])
