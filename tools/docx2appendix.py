#!/usr/bin/env python3
"""Convert Word appendix files into a LaTeX ``appendices.tex`` for muthesis2026.

The script reads one or more ``appendix_<Letter>_*.docx`` files, keeps only the
content (headings, paragraphs, tables) and emits LaTeX that fits the Mahidol
thesis class: ``\\appendices`` followed by one ``\\chapter`` per appendix, with
``A.1`` style headings mapped onto ``\\section``/``\\subsection`` and every
"Table X.Y ..." paragraph consumed as the caption of the table that follows it.

Usage:
    python3 tools/docx2appendix.py FILE.docx [FILE.docx ...] -o appendices.tex

Requires: python-docx  (pip install python-docx)
"""
import argparse
import re
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

# ---------------------------------------------------------------------------
# LaTeX escaping
# ---------------------------------------------------------------------------
# Order matters: backslash first, then the rest.
_SIMPLE = [
    ("\\", r"\textbackslash{}"),
    ("&", r"\&"),
    ("%", r"\%"),
    ("$", r"\$"),
    ("#", r"\#"),
    ("_", r"\_\allowbreak{}"),   # allow long snake_case identifiers to wrap
    ("{", r"\{"),
    ("}", r"\}"),
    ("~", r"\textasciitilde{}"),
    ("^", r"\textasciicircum{}"),
    ("<", r"\textless{}"),
    (">", r"\textgreater{}"),
    ("|", r"\textbar{}"),
]

_UNICODE = {
    "µ": r"\textmu{}",   # µ micro sign
    "μ": r"\textmu{}",   # μ greek mu
    "×": r"$\times$",    # ×
    "−": r"$-$",         # − U+2212 true minus sign
    "±": r"$\pm$",       # ±
    "≥": r"$\ge$",       # ≥
    "≤": r"$\le$",       # ≤
    "→": r"$\rightarrow$",  # →
    "–": "--",           # – en dash
    "—": "---",          # — em dash
    "‘": "`",            # ‘
    "’": "'",            # ’
    "“": "``",           # “
    "”": "''",           # ”
    "…": r"\ldots{}",    # …
    "°": r"\textdegree{}",  # °
    " ": "~",            # non-breaking space
    "‑": "-",            # non-breaking hyphen
}


# Combined char -> replacement map. Slash gets an \allowbreak breakpoint so
# long paths can wrap. Built into ONE regex: re.sub scans left to right and
# never re-examines the text it inserts, so replacement values may safely
# contain LaTeX metacharacters (braces, backslashes) without double-escaping.
_ESC = {"/": r"/\allowbreak{}"}
_ESC.update(_SIMPLE)
_ESC.update(_UNICODE)
_ESC_RE = re.compile("|".join(re.escape(k) for k in _ESC))


def escape(text):
    """Escape a run of plain text for LaTeX in a single, non-recursive pass."""
    return _ESC_RE.sub(lambda m: _ESC[m.group(0)], text)


# ---------------------------------------------------------------------------
# Document walking
# ---------------------------------------------------------------------------
def iter_block_items(doc):
    """Yield paragraphs and tables in document order."""
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


HEADING_RE = re.compile(r"^([A-Z])\.(\d+)(?:\.(\d+))?(?:\.(\d+))?\s+(.*\S)\s*$")
CAPTION_RE = re.compile(r"^Table\s+[A-Z]\.\d+\s+(.*\S)\s*$")


def table_columns_are_numeric(rows):
    """Return a list of booleans: True where every data cell is numeric."""
    ncols = len(rows[0])
    flags = []
    num_re = re.compile(r"^[\d,]+$")
    for c in range(ncols):
        vals = [rows[r][c].strip() for r in range(1, len(rows))]
        vals = [v for v in vals if v and v != "--" and v != "—"]
        flags.append(bool(vals) and all(num_re.match(v) for v in vals))
    return flags


def max_cell_len(rows, col):
    return max(len(rows[r][col]) for r in range(len(rows)))


def render_table(rows, caption, label):
    """Render a table. Chooses longtable / tabularx / plain tabular sensibly."""
    ncols = len(rows[0])
    numeric = table_columns_are_numeric(rows)
    header = rows[0]
    data = rows[1:]
    # "Text-heavy" is judged from the DATA cells only. A long header over short
    # numeric data (common in results tables) should be scaled, not wrapped.
    data_maxlen = [max((len(r[c]) for r in data), default=0) for c in range(ncols)]
    long_cols = [i for i, m in enumerate(data_maxlen) if m > 24]

    def cell(text, bold=False):
        out = escape(text)
        return (r"\textbf{" + out + "}") if bold and text.strip() else out

    def row_tex(cells, bold=False):
        return " & ".join(cell(c, bold) for c in cells) + r" \\"

    # Long tables (many rows, short cells) -> longtable that breaks across pages.
    if len(data) > 22 and not long_cols:
        colspec = "".join("r" if numeric[c] else "l" for c in range(ncols))
        lines = [r"\begin{center}", r"\begin{small}",
                 r"\begin{longtable}{" + colspec + "}",
                 r"\caption{" + escape(caption) + r"}\label{" + label + r"}\\",
                 r"\hline",
                 row_tex(header, bold=True),
                 r"\hline", r"\endfirsthead",
                 r"\multicolumn{" + str(ncols) + r"}{l}{\textit{(continued)}}\\",
                 r"\hline",
                 row_tex(header, bold=True),
                 r"\hline", r"\endhead",
                 r"\hline", r"\endfoot"]
        for r in data:
            lines.append(row_tex(r))
        lines += [r"\hline", r"\end{longtable}", r"\end{small}", r"\end{center}"]
        return "\n".join(lines)

    # Text-heavy tables -> tabularx so long columns wrap to the text width.
    if long_cols:
        parts = []
        for c in range(ncols):
            if c in long_cols:
                parts.append("X")
            elif numeric[c]:
                parts.append("r")
            else:
                parts.append("l")
        colspec = "".join(parts)
        lines = [r"\begin{table}[htbp]", r"\centering", r"\small",
                 r"\caption{" + escape(caption) + r"}\label{" + label + r"}",
                 r"\begin{tabularx}{\textwidth}{" + colspec + "}",
                 r"\hline",
                 row_tex(header, bold=True),
                 r"\hline"]
        for r in data:
            lines.append(row_tex(r))
        lines += [r"\hline", r"\end{tabularx}", r"\end{table}"]
        return "\n".join(lines)

    # Remaining tables have short cells. Render a plain tabular, always guarded
    # by an adjustbox max-width so a wide row can only shrink, never overflow.
    # Wide tables (>= 6 columns) additionally start at \small.
    colspec = "".join("r" if numeric[c] else "l" for c in range(ncols))
    size = [r"\small"] if ncols >= 6 else []
    lines = [r"\begin{table}[htbp]", r"\centering",
             r"\caption{" + escape(caption) + r"}\label{" + label + r"}",
             r"\begin{adjustbox}{max width=\textwidth}"]
    lines += size
    lines += [r"\begin{tabular}{" + colspec + "}",
              r"\hline",
              row_tex(header, bold=True),
              r"\hline"]
    for r in data:
        lines.append(row_tex(r))
    lines += [r"\hline", r"\end{tabular}", r"\end{adjustbox}", r"\end{table}"]
    return "\n".join(lines)


def table_to_rows(tbl):
    rows = []
    for r in tbl.rows:
        rows.append([c.text.replace("\n", " ").strip() for c in r.cells])
    return rows


def convert_file(path, letter):
    """Convert one appendix docx into a list of LaTeX lines (one chapter)."""
    doc = Document(str(path))
    blocks = list(iter_block_items(doc))

    out = []
    pending_caption = None
    table_counter = 0
    chapter_set = False
    seen_first_para = False  # first non-empty paragraph is "Appendix X"

    for blk in blocks:
        if isinstance(blk, Table):
            table_counter += 1
            caption = pending_caption or f"Table {letter}.{table_counter}"
            label = f"tab:app{letter.lower()}-{table_counter}"
            out.append(render_table(table_to_rows(blk), caption, label))
            out.append("")
            pending_caption = None
            continue

        text = blk.text.strip()
        if not text:
            continue

        # First paragraph "Appendix X" -> skip; second -> chapter title.
        if not seen_first_para and re.match(r"^Appendix\s+[A-Z]\s*$", text):
            seen_first_para = True
            continue
        if not chapter_set:
            out.append(r"\chapter{" + escape(text) + "}")
            out.append("")
            chapter_set = True
            continue

        # Table caption line -> hold for the next table.
        m = CAPTION_RE.match(text)
        if m:
            pending_caption = m.group(1)
            continue

        # Numbered heading -> section / subsection / subsubsection.
        m = HEADING_RE.match(text)
        if m and m.group(1) == letter:
            _, n1, n2, n3, title = m.groups()
            if n3:
                out.append(r"\subsubsection{" + escape(title) + "}")
            elif n2:
                out.append(r"\subsection{" + escape(title) + "}")
            else:
                out.append(r"\section{" + escape(title) + "}")
            out.append("")
            continue

        # Ordinary paragraph. Italicise a leading "Note." style lead-in.
        lead = re.match(r"^(Note|Implementation limitation|Dashboard status "
                        r"semantics)\.\s+(.*)$", text, re.S)
        if lead:
            out.append(r"\noindent\textit{" + escape(lead.group(1))
                       + ".} " + escape(lead.group(2)))
        else:
            out.append(escape(text))
        out.append("")

    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="+", help="appendix .docx files")
    ap.add_argument("-o", "--output", default="appendices.tex",
                    help="output .tex file (default: appendices.tex)")
    args = ap.parse_args()

    # Sort by the appendix letter found in the filename, not by filename.
    def letter_of(p):
        m = re.search(r"appendix_([A-Z])", Path(p).name)
        return m.group(1) if m else Path(p).name.upper()

    files = sorted(args.files, key=letter_of)

    lines = [
        "% Generated by tools/docx2appendix.py -- do not edit by hand.",
        "% Regenerate with: python3 tools/docx2appendix.py <docx files> -o appendices.tex",
        r"\appendices",
        "",
    ]
    for p in files:
        letter = letter_of(p)
        lines.append(f"% ---- Appendix {letter}: {Path(p).name} ----")
        lines.extend(convert_file(p, letter))
        lines.append("")

    Path(args.output).write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Wrote {args.output} ({len(files)} appendices)")


if __name__ == "__main__":
    main()
