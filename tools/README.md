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

1. **Citations.** Cross-references like `[9]` and `[28]` are copied as literal
   text. Replace them with real `\cite{key}` commands and add the entries to
   `references.bib` if you want them linked and renumbered.
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
