#!/usr/bin/env python3
"""Assemble the single-file main document from the generated parts.

`docx2thesis.py` and `docx2appendix.py` each write standalone .tex fragments.
This script inlines them into one self-contained
`main-thematic-traditional.tex`, following the front-matter order the
muthesis2026 template defines:

    \\maketitle
    \\acknowledgements{...}
    \\abstract{...}
    \\tableofcontents
    \\listoftables
    \\listoffigures
    \\listofabbreviations{...}
    <chapters>  <references>  <appendices>
    \\biography

Everything above \\begin{document} (document class, packages, \\title,
\\author, \\input{preamble}) is preserved from the existing main file, so
hand-edited settings survive a rebuild.

Usage:
    python3 tools/build_main.py [--keep-parts]

By default the inlined fragment files are deleted afterwards, leaving one file
to edit. Pass --keep-parts to keep them.
"""
import argparse
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
MAIN = os.path.join(ROOT, "main-thematic-traditional.tex")

PARTS = ["acknowledgements", "abstract", "chapters", "references",
         "appendices", "abbreviations"]


def read(name):
    path = os.path.join(ROOT, name + ".tex")
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return f.read().strip("\n")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--keep-parts", action="store_true",
                    help="do not delete the fragment files after inlining")
    args = ap.parse_args()

    with open(MAIN, encoding="utf-8") as f:
        current = f.read()
    head, sep, _ = current.partition(r"\begin{document}")
    if not sep:
        sys.exit("main file has no \\begin{document}")

    part = {name: read(name) for name in PARTS}
    missing = [n for n in ("chapters", "references", "appendices")
               if not part.get(n)]
    if missing:
        sys.exit("missing generated parts: " + ", ".join(missing)
                 + " (run docx2thesis.py / docx2appendix.py first)")

    body = [head.rstrip() + "\n", r"\begin{document}", "", r"\maketitle", "",
            r"\linespacing{1.5}"]

    if part.get("acknowledgements"):
        body += [r"\acknowledgements{" + part["acknowledgements"] + "}", ""]

    if part.get("abstract"):
        body += [r"\abstract{%", r"\linespacing{1.2}", part["abstract"], "}", ""]

    body += [r"\tableofcontents", r"\listoftables", r"\listoffigures"]
    if part.get("abbreviations"):
        body += [r"\listofabbreviations{", part["abbreviations"], "}"]
    body.append("")

    body += [part["chapters"], ""]

    # The REFERENCES table-of-contents entry is restored in the preamble via
    # natbib's \bibsection hook (see the main file), not here: doing it around
    # this \input would either record the wrong page or add a blank page.
    body += [part["references"], ""]

    body += [part["appendices"], "", r"\biography", "", r"\end{document}", ""]

    with open(MAIN, "w", encoding="utf-8") as f:
        f.write("\n".join(body))

    removed = []
    if not args.keep_parts:
        for name in PARTS:
            path = os.path.join(ROOT, name + ".tex")
            if os.path.exists(path):
                os.remove(path)
                removed.append(name + ".tex")

    print("Rebuilt main-thematic-traditional.tex")
    if removed:
        print("Inlined and removed: " + ", ".join(removed))


if __name__ == "__main__":
    main()
