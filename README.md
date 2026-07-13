# Claude Code Excel2Latex Table

A [Claude Code](https://claude.com/claude-code) skill that converts an Excel
table into a publication-quality LaTeX table sized to fit on one page in
Overleaf — proper `booktabs` rules, `siunitx` decimal alignment, correct
handling of merged headers and spacer columns — and verifies the result by
actually compiling it, iterating until it fits.

## The problem this solves

If you design tables in Excel (styled with narrow spacer columns for visual
breathing room, merged header cells, bold summary rows) and need them in a
LaTeX paper, a straight conversion is usually a mess: spacer columns become
real ragged columns, numbers don't align on the decimal point, borders become
`\hline` grids instead of `booktabs` rules, and the table overflows the page
width once pasted into Overleaf. Fixing this by hand — nudging column
widths, guessing at font sizes, re-pasting into Overleaf to see if it fits
this time — is the tedious part this skill removes.

## How it works

1. **Extract** — a script reads the `.xlsx` file's structure (values, merges,
   bold/border formatting, column widths) into compact JSON, flagging narrow
   empty columns as spacers rather than data.
2. **Lay out** — Claude decides column types, header structure, and rule
   placement using a small, editable layout spec — the judgment-heavy step.
3. **Render** — a script deterministically turns the layout spec + extracted
   JSON into the LaTeX tabular body, so numbers are never retyped by hand
   between iterations.
4. **Compile & check** — a script compiles the table inside your paper's
   actual preamble/class and reports whether it fits, numerically (exact
   point measurements), not by eyeballing a screenshot. Claude iterates on
   the layout spec using this numeric feedback until it fits, then does one
   final visual check on a cropped preview image before handing it back.

See [`xlsx2latex/SKILL.md`](xlsx2latex/SKILL.md) for the full workflow this
skill follows.

## Installation

Requires [Claude Code](https://claude.com/claude-code).

```bash
git clone https://github.com/<your-username>/xlsx2latex.git
mkdir -p ~/.claude/skills
cp -r xlsx2latex/xlsx2latex ~/.claude/skills/xlsx2latex
```

(Or copy `xlsx2latex/` into a project's `.claude/skills/` directory instead
of `~/.claude/skills/` to scope it to one project.)

Install the Python dependencies the scripts use:

```bash
pip install -r xlsx2latex/requirements.txt
```

You'll also need a LaTeX engine to compile previews — pick one:

- **pdflatex + latexmk** (recommended — matches Overleaf's own engine, so
  local previews are the most faithful):
  ```bash
  brew install --cask basictex   # macOS; ~100MB
  tlmgr install latexmk booktabs siunitx multirow threeparttable makecell
  ```
  On Linux, install any TeX Live distribution (`texlive-latex-extra` covers
  the packages above) via your package manager.
- **tectonic** (lighter, zero-config, no system installer needed):
  ```bash
  brew install tectonic   # or: cargo install tectonic
  ```
  Tectonic is a different engine than Overleaf's pdflatex, so treat it as a
  good local-iteration fallback rather than a byte-for-byte match.

## Usage

In Claude Code, just describe what you want:

> "Convert the table in `results.xlsx` to a LaTeX table for my paper, it
> needs to fit in a single column"

The skill triggers automatically on requests like this. The first time you
use it for a given paper, it'll ask for your LaTeX preamble (and `.bib` file,
if the table has citations) so it can compile against your actual document
class and margins — after that, converting more tables for the same paper
needs no further setup.

## Repository layout

```
xlsx2latex/
├── SKILL.md                    # the skill Claude Code reads and follows
├── scripts/
│   ├── extract_xlsx.py         # .xlsx -> structure JSON
│   ├── render_body.py          # structure JSON + layout spec -> LaTeX tabular
│   └── compile_preview.py      # compile, numeric fit-check, cropped preview PNG
├── references/
│   └── table-patterns.md       # booktabs/siunitx recipes (multi-level headers,
│                                # significance stars, threeparttable notes, ...)
└── examples/
    └── sample_results.xlsx     # a demo table exercising spacers/merges/bold/borders
```

## Scope

Tables that need to span multiple pages (`longtable`) are intentionally out
of scope for now — the working assumption is "always fits on one page,"
using font/spacing/rotation adjustments to get there. If a table is
genuinely too large even in landscape, the skill will say so rather than
force a bad fit.

## License

MIT — see [LICENSE](LICENSE).
