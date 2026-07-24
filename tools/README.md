# Word → LaTeX converters

Two scripts turn the Word sources into LaTeX for the `muthesis2026` template:

- **`docx2thesis.py`** — the main thesis file (front matter, Chapters I–V,
  figures, tables, and the IEEE reference list).
- **`docx2appendix.py`** — the appendix files (Appendix A–F). Its escaping,
  table-rendering, and citation helpers are imported and reused by
  `docx2thesis.py`, so table styling and `\cite{refN}` numbering stay identical
  across the body and the appendices.

## docx2thesis.py

```bash
pip install python-docx
python3 tools/docx2thesis.py thesis.docx
```

Outputs, written next to the template:

| File | Purpose |
|------|---------|
| `chapters.tex` | Chapters I–V as `\chapter`/`\section`/`\subsection`, with figures and tables |
| `references.tex` | a `thebibliography` built verbatim from the numbered IEEE list |
| `abstract.tex` | the abstract body, pulled in by `\abstract{\input{abstract}}` |
| `acknowledgements.tex` | the acknowledgements body |
| `figures/figX-Y.png` | every embedded figure, extracted and named by its number |

It skips the title page, the embedded table of contents, the appendices (done
separately), and the list of abbreviations — the class regenerates the first
two. `CHAPTER I` + its title line become one `\chapter{}`; `1.2 Heading`
becomes `\section{Heading}` (the class re-numbers); every `[N]` citation
becomes `\cite{refN}`, matched by `\bibitem{refN}` in `references.tex`. It also
prints the front-matter metadata (author, advisors, keywords, …) to copy into
`preamble.tex`.

### Manual follow-ups after conversion

- **Title.** The title page says "…Ransomware Detection"; the abstract-page
  header adds "& Prevention". Set `\title{}` to the official wording.
- **Committee & ranks.** `preamble.tex` is filled with the advisor
  (Ittipon Rassameeroj) and co-advisor (Vasaka Visoottiviseth); the third
  advisory-committee member (Assadarat Khurat, Dr.-Ing.), the exam-committee
  chair, the faculty dean, and the academic ranks (`Dr.` vs `Asst.~Prof.` …)
  still need confirming. Thai-language fields are left as placeholders.
- **Cross-references.** "Table 3.3", "Figure 4.2", "Chapter 4" are plain text;
  convert to `\ref{}`/`\autoref{}` if you want live links (labels already exist
  on the generated tables and figures, e.g. `tab:4-4`, `fig:2-1`).
- **Figures.** Extracted at their embedded resolution; replace any you have a
  higher-resolution original for.

---

# Word → LaTeX appendix converter

`docx2appendix.py` converts Word appendix files into a single `appendices.tex`
that plugs into the `muthesis2026` class. It keeps only the *content*
(headings, paragraphs, tables) — the thesis formatting is supplied by the
template — and maps it onto the class structure:

| Word                                  | LaTeX                                   |
|---------------------------------------|-----------------------------------------|
| First paragraph `Appendix A`          | dropped (the class prints `APPENDIX A`) |
| Second paragraph (the title)          | `\chapter{...}` → numbered A, B, C, …   |
| `A.1  Heading`                        | `\section{Heading}` → auto-numbered A.1 |
| `A.1.1  Heading`                      | `\subsection{Heading}`                  |
| `Table A.1  Caption.` + table         | `table`/`longtable` with that caption   |
| `Note. ...`                           | `\textit{Note.} ...`                    |

## Requirements

```bash
pip install python-docx
```

## Usage

```bash
python3 tools/docx2appendix.py path/to/appendix_*.docx -o appendices.tex
```

Files are ordered by the letter in their name (`appendix_A`, `appendix_B`, …),
not by filename, so upload/download order does not matter.

The generated `appendices.tex` starts with `\appendices`; it is pulled into the
document with a single `\input{appendices}` placed after `\bibliography` and
before `\biography` (already wired into `main-thematic-traditional.tex`). The
template needs three extra packages, also already added there:

```latex
\usepackage{tabularx}   % width-constrained tables
\usepackage{longtable}  % tables that break across pages
\usepackage{adjustbox}  % scale wide tables to \textwidth
```

## How tables are rendered

The script picks a layout per table so nothing overflows the A4 text block:

- **Text-heavy** (a data cell longer than ~24 chars) → `tabularx` with `X`
  columns that wrap to the text width.
- **Long lists** (more than 22 rows, short cells) → `longtable` that breaks
  across pages and repeats the header.
- **Wide** (6+ columns) or any other table → plain `tabular` wrapped in
  `adjustbox{max width=\textwidth}`, so a wide row shrinks to fit instead of
  running into the margin. Wide tables also start at `\small`.

Character handling: LaTeX metacharacters are escaped, common Unicode is mapped
(µ→`\textmu`, ×→`$\times$`, −→`$-$`, en/em dashes, curly quotes, …), and
`\allowbreak` is inserted after `_` and `/` so long identifiers
(`RUN_IceFire_Set1_182048`) and paths wrap.

## Manual steps after generating

The converter handles structure and formatting, but a few things still need a
human pass:

1. **Citations.** Numeric reference markers (`[9]`, `[28]`, …) are turned into
   real `\cite{}` commands via the `CITE_MAP` at the top of the script; any
   number not in that map is left as literal `[N]`. The two currently mapped
   (`bringoltz2025clear` for CLEAR and `davies2022napierone` for NapierOne)
   have entries in `references.bib`. Reconcile both keys with your main-thesis
   bibliography so the numbers renumber correctly in the full document.
2. **Cross-references.** Mentions of "Table 3.3", "Section 3.2", "Chapter 4"
   are plain text. Convert to `\ref{}`/`\autoref{}` if you want live links.
3. **Front matter.** Author, title, committee, and abstract are *not* part of
   the appendices — fill those in `preamble.tex` and the `main-*.tex` file.
4. **Proofread wide tables.** The scaled tables (e.g. Table F.1) are legible
   but small; consider splitting any you would rather keep at full size.

## Regenerating

Re-run the command whenever the Word appendices change, then recompile:

```bash
python3 tools/docx2appendix.py appendix_*.docx -o appendices.tex
latexmk -pdf main-thematic-traditional.tex
```
