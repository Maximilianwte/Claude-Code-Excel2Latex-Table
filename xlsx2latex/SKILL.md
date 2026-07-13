---
name: xlsx2latex
description: Converts an Excel (.xlsx) table into a publication-quality LaTeX table sized to fit on one page in Overleaf, using booktabs/siunitx conventions instead of a naive grid conversion, then iteratively checks and fixes the fit by actually compiling it. Use this whenever the user wants to turn an Excel/spreadsheet table into LaTeX for a paper, mentions pasting a table into Overleaf, complains about a LaTeX table not fitting on the page or looking misaligned, or asks to convert xlsx/csv data into a booktabs/academic-style table — even if they don't say "xlsx2latex" or name this skill explicitly. Also use it proactively if the user pastes raw Excel-derived data and asks for a LaTeX table.
---

# xlsx2latex

Turn an Excel table into a LaTeX table that actually looks right in Overleaf
— correct decimal alignment, no vertical rules, fits on one page — without
the usual back-and-forth of copy-pasting into Overleaf, seeing it overflow or
misalign, and manually fiddling with column widths.

## Why this needs more than a straight conversion

A naive xlsx-to-LaTeX conversion turns every Excel column into a real LaTeX
column, every border into `\hline`, and every float into its raw
`0.812000000001` value. That produces exactly the mess this skill exists to
avoid: narrow spacer columns the user added purely for visual breathing room
become extra ragged columns, numbers don't align on the decimal point, and
the whole thing overflows the page width in Overleaf.

The fix is to treat Excel styling as *intent* to interpret, not data to copy
literally, and to verify the result by actually compiling it rather than
guessing whether it'll fit.

## Prerequisites

A LaTeX engine must be installed to run the compile step:

- **pdflatex + latexmk** (recommended) — this is what Overleaf itself uses,
  so a local compile is the most faithful preview of what the user will see.
  Install via `brew install --cask basictex` (macOS, ~100MB) then
  `tlmgr install latexmk booktabs siunitx multirow threeparttable makecell`,
  or any TeX Live distribution on other platforms.
- **tectonic** — a good zero-config fallback (`brew install tectonic` /
  `cargo install tectonic`, no system installer needed). It fetches packages
  on demand and compiles fine, but it's a different engine than Overleaf's
  pdflatex, so treat it as "good enough to iterate locally," not a
  byte-for-byte guarantee. Prefer pdflatex/latexmk when both are available.

If neither is installed, tell the user and offer the install command rather
than silently failing partway through a table.

## Workflow

### Step 0 — Find or set up the project

Tables belong to a **project** (one per paper) so the user's preamble and
bibliography are configured once, not re-pasted for every table. Projects
live in `.xlsx2latex/projects/<name>/` relative to wherever the user is
working.

Look for an existing project first (check the current directory for
`.xlsx2latex/projects/*`, or ask the user which paper this table belongs to).
If none exists, this is a new project — see **Project setup** below before
doing anything else. If one exists, skip straight to Step 1; setup should
happen at most once per paper.

### Step 1 — Extract the table's structure

```
python scripts/extract_xlsx.py path/to/table.xlsx --out structure.json
```

Add `--sheet "Name"` or `--range A1:F12` if the table isn't the whole active
sheet. This script does the mechanical, error-prone-if-done-by-hand part:
reading merges, bold/italic, column widths, border positions, and rendering
numbers the way Excel displays them (respecting decimal places, %, thousands
separators) rather than as raw floats. Read the resulting JSON's
`style_notes` block — it flags spacer columns, bold rows, and border
positions as *hints* for the decisions below, not as instructions to follow
blindly. A column the script flags as a spacer because it's narrow and empty
almost always is one, but use judgment if something looks off.

### Step 2 — Decide the layout (this is the judgment-heavy step)

Write a small **layout spec** JSON (see schema below) that captures the
decisions a human designing this table for print would make. This file is
the *only* thing you write or edit by hand each iteration — never retype the
data values themselves; `render_body.py` pulls those straight from
`structure.json` so numbers can't get transposed or mistyped between rounds,
and editing a 10-20 line layout spec is far cheaper than rewriting a whole
tabular block every time something needs to shrink.

Decisions to make here, and why:

- **Column types.** Any column of comparable decimal numbers should be an
  `S` column (siunitx) with `table-format` set to the widest value in that
  column — this is what makes numbers align on the decimal point instead of
  looking ragged. Text/label columns are `l` (rarely `c`/`r`).
- **Spacer columns are not real columns.** Never map a `structure.json`
  spacer column into the layout's `columns` array. The visual gap it gave in
  Excel becomes `tabcolsep` spacing or a slightly wider stub column instead.
- **Booktabs, never grid lines.** `\toprule`/`\midrule`/`\bottomrule` only.
  Use the `style_notes.border_top_rows` / `border_bottom_rows` /
  `double_border_bottom_rows` hints from `structure.json` to decide where
  `\midrule` belongs (a double-bottom border usually signals "final/best
  row," which reads naturally as a midrule-then-bold-row rather than a
  literal double rule).
- **Multi-level headers.** If `structure.json` has merges in the header rows,
  build `header_rows` with `{"text": ..., "span": N}` entries and matching
  `cmidrules`. Read `references/table-patterns.md` for the exact pattern —
  in particular, any header cell sitting over an `S` column must be wrapped
  in a multicolumn even at span 1, or siunitx will try to parse the header
  text as a number and error out.
- **Header text can be abbreviated or wrapped**, but the underlying data
  values are never touched — that asymmetry is the point of keeping this
  step separate from `render_body.py`.

Also write the **table wrapper** (`table.tex`): a `\begin{table}...\end{table}`
(or `\begin{threeparttable}` if there are notes — see
`references/table-patterns.md`) containing `\caption`, `\label`, and
`\input{body.tex}` where the tabular goes. This is a normal small piece of
LaTeX you write directly; it doesn't need a script since it's mostly static
per table.

**Layout spec schema:**

```json
{
  "columns": [
    {"type": "l"},
    {"type": "S", "table_format": "1.3"}
  ],
  "header_rows": [
    ["Method", {"text": "Accuracy", "span": 2, "align": "c"}]
  ],
  "cmidrules": [{"row": 0, "span": "2-3"}],
  "header_row_count": 2,
  "midrule_after_rows": [4],
  "tabcolsep": "6pt",
  "booktabs": {"toprule": true, "midrule_after_header": true, "bottomrule": true}
}
```

- `columns` must have exactly one entry per **non-spacer** column, in order.
- `header_rows` items are either a plain string or `{"text","span","align"}`
  for a `\multicolumn` cell.
- `cmidrules[].row` is the 0-indexed position within `header_rows` the rule
  goes under (not the row index in the original spreadsheet).
- `header_row_count` tells `render_body.py` how many leading spreadsheet rows
  to skip when emitting data rows, since header text comes from `header_rows`
  instead.
- `midrule_after_rows` uses the *spreadsheet* row index (0-indexed within the
  extracted range) from `structure.json`.

### Step 3 — Render the tabular body

```
python scripts/render_body.py structure.json layout.json --out body.tex
```

Deterministic — no model judgment involved. Bold rows (from
`structure.json`'s `style_notes.bold_rows`), significance-star stripping,
character escaping, and multirow/multicolumn from merges all happen here
automatically.

### Step 4 — Compile and check the fit, numerically first

```
python scripts/compile_preview.py --project .xlsx2latex/projects/<name> \
    --body body.tex --table table.tex
```

This wraps the table in the project's *actual* preamble and document class
(not a generic one) and compiles it, then reports a compact numeric summary:

```
COMPILE: success
WIDTH: table=482.1pt target=455.2pt (OVERFLOWS, 26.9pt over)
OVERFULL: 1 warning(s), worst 26.9pt too wide
```

Iterate on `layout.json` using **only this numeric output** — don't render an
image yet. A width comparison in points is exact and costs a few lines of
text; a rendered page image costs far more and can't tell you *how much*
overflow there is, only that something looks off. Numbers converge faster.

If the summary reports undefined citation keys, treat that as a distinct,
higher-priority problem before worrying about width at all: an undefined
`\cite{}` renders as a narrow `[?]` placeholder, which can make an
over-width table numerically *appear* to fit. Fix the citation key (check
spelling against the project's `references.bib`) and recompile before
trusting any width number that ran alongside undefined citations.

**When it doesn't fit, apply exactly one change and recompile — don't stack
several changes at once, or you won't know which one mattered.** Try changes
in this order; each is progressively more disruptive to the table's
appearance, so stop at the first one that gets it under budget:

1. Shrink `tabcolsep` (try `4pt`), abbreviate a header, or wrap a long header
   across two lines with `\makecell`.
2. Drop the font size one step: `\small`, then `\footnotesize` if still
   needed. Wrap the tabular in `{\footnotesize ... }` inside `table.tex`.
3. Rotate to a landscape table (`sidewaystable`, needs `\usepackage{rotating}`
   in the project preamble) — usually a better outcome than an unreadably
   small font when a table just has too many columns.
4. `\resizebox` as the last resort. It's the easiest fix mechanically but it
   scales the whole table including its font size relative to the body
   text and every other table in the paper, which is the kind of
   inconsistency a reviewer or copyeditor will flag. Tell the user this is
   what you're doing and why, rather than silently reaching for it first.

Cap this loop at **5 rounds**. If it still doesn't fit after trying the
relevant options above, stop and tell the user honestly — e.g. "this table
has too many columns to fit in a single-column layout even in footnotesize;
I'd recommend landscape orientation or splitting it into two tables" — rather
than forcing a `\resizebox` that leaves the font comically small.

### Step 5 — One visual check, once the numbers fit

```
python scripts/compile_preview.py --project .xlsx2latex/projects/<name> \
    --body body.tex --table table.tex --final
```

This adds a cropped preview PNG (whitespace trimmed, so you're not paying to
look at page margins) to the same numeric summary. Read the image now — this
is the one point in the loop where visual inspection earns its cost, since it
catches things a width number can't: an awkward two-line header wrap, a
spacer proportion that looks off, decimals that are technically aligned but
visually cramped. If something looks wrong, go back to Step 2 for that one
issue, then re-run Steps 3-5 rather than re-checking numerically-fine things
by eye again.

### Deliverables

Once the preview looks right, give the user:

- The final `table.tex` (or a combined snippet) — ready to drop into Overleaf
  with `\input{}`, or to paste directly.
- Any `\usepackage{}` lines not already covered by the project's stored
  preamble.
- The final cropped preview PNG, so they can sanity-check before pasting into
  Overleaf without needing to compile it themselves first.

## Project setup

Run this once per paper, not once per table — that's the entire point of
keeping it separate from the per-table workflow above.

1. Ask the user to paste their LaTeX preamble (the packages/custom commands
   from the top of their Overleaf `main.tex`) and, if the table needs
   citations, their `.bib` file. Before saving, **strip the `\documentclass{...}`
   line and any `\usepackage[...]{geometry}` line** — `compile_preview.py`
   generates those itself from `config.json` (step 2 below), and leaving them
   in `preamble.tex` too causes a fatal "Two `\documentclass`" error since the
   wrapper document ends up with both. Save what's left verbatim to
   `.xlsx2latex/projects/<name>/preamble.tex`, and the `.bib` file (unchanged)
   to `references.bib`.
2. Ask (or infer from the preamble) the document class, class options, and
   margin geometry — these came out of the preamble in step 1, so they need
   to go *somewhere*, and `config.json` is that somewhere. Also decide
   `target_width`: for a single-column paper this is `"textwidth"`; for a
   two-column paper (like NeurIPS/ACL style) where the table should fit in
   one column, use `"columnwidth"` instead — comparing a table against the
   wrong one of these is the most common way a table "fits" in your local
   preview but overflows once actually placed in the real two-column layout.
   Save to `config.json`:

   ```json
   {
     "document_class": "article",
     "class_options": "11pt",
     "geometry": "margin=1in",
     "engine": "auto",
     "bib_backend": "natbib",
     "target_width": "textwidth"
   }
   ```

   `bib_backend` is `"none"`, `"natbib"` (uses `\bibliographystyle`/
   `\bibliography`, compiled via `bibtex`), or `"biblatex"` (uses
   `\printbibliography`, compiled via `biber`) — detect this from whether the
   preamble loads `natbib` or `biblatex`. `engine` is normally left as
   `"auto"`, which prefers `latexmk`, then `pdflatex`, then `tectonic`, in
   that order. If the actual paper uses a two-column class the local
   toolchain doesn't have installed (e.g. a conference's custom `.cls`), it's
   fine to compile the preview against a plain two-column `article` with a
   matching `\columnwidth` — the point is measuring against the right target
   width, not reproducing the exact class.

3. **Trial-compile immediately** with a table-free empty document
   (`compile_preview.py` does this automatically the first time it's run
   against a fresh project, since `table.tex` can just be a placeholder
   sentence) to catch a missing journal `.cls` file (e.g. `elsarticle.cls`)
   right away rather than mid-table later. If it fails on a missing class or
   style file, ask the user to also provide that file and place it in the
   project directory.

Updating later is just overwriting `preamble.tex` or `references.bib` with a
new paste — no other setup step needs to be repeated.

**Token cost note:** the preamble and `.bib` are read once at setup and then
live on disk. Don't re-read them into context on later table runs — only the
handful of fields in `config.json` are needed per table, and
`compile_preview.py` handles stitching the actual preamble file in at compile
time without the model needing to see its contents again.

## Reference material

`references/table-patterns.md` has concrete LaTeX recipes: S-column syntax,
multi-level headers, threeparttable notes, significance stars, multirow stub
columns, landscape tables. Read it when a table needs one of these patterns
rather than re-deriving the syntax from scratch.

## Out of scope

Tables that need to break across multiple pages (`longtable`) are explicitly
not handled — the hard constraint here is "always fits on one page," using
the shrink/rotate strategies above. If a table is fundamentally too large for
one page even in landscape, say so and suggest splitting it into two tables
rather than reaching for `longtable`.
