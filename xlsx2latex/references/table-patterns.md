# LaTeX table patterns for publication tables

Recipes for the situations that come up repeatedly when turning an Excel table
into a LaTeX table for a paper. Read the section you need — you don't need to
read this whole file for a simple table.

## Required packages

```latex
\usepackage{booktabs}       % \toprule \midrule \bottomrule
\usepackage{siunitx}        % S columns, decimal alignment, \num{}
\usepackage{multirow}       % \multirow for row-spanning stub cells
\usepackage{threeparttable} % table notes that don't overflow the table width
\usepackage{makecell}       % \makecell for wrapped/multi-line headers
\newcommand{\sym}[1]{\textsuperscript{#1}}  % significance stars, footnote marks
```

If the project already has a preamble, check whether these are present before
adding duplicates.

## Basic booktabs table

Never use `\hline` or vertical `|` rules in a table meant for a paper — three
horizontal rules is the convention nearly every journal and style guide
expects.

```latex
\begin{tabular}{lS[table-format=1.3]S[table-format=1.3]}
\toprule
Method & {Accuracy} & {F1} \\
\midrule
Baseline & 0.812 & 0.795 \\
Ours     & 0.891 & 0.876 \\
\bottomrule
\end{tabular}
```

Note `{Accuracy}` — a plain-text header over an `S` column must be wrapped in
braces or `\multicolumn{1}{c}{...}`, otherwise siunitx tries to parse the
header text itself as a number and errors out. `render_body.py` in this skill
handles this automatically; if you're writing headers by hand, remember it.

## S columns: decimal alignment

`S` columns from siunitx right-align on the decimal point, which is what
makes a column of numbers actually easy to scan — this is the single biggest
visual upgrade over a naive Excel-to-LaTeX conversion.

`table-format=1.3` means "1 digit before the decimal, 3 after." Set it to
match the widest number in the column so all rows align, e.g. a column with
values from 0.8 to 12.45 needs `table-format=2.2`.

```latex
S[table-format=1.3]     % 0.891
S[table-format=2.1]     % 12.4
S[table-format=-1.2]    % allow a leading minus sign: -0.45
```

Negative numbers: `S` columns handle a literal minus sign in the data
correctly (`-0.45` renders as a proper math minus, not a hyphen) as long as
`table-format` reserves space for the sign. Outside `S` columns (plain `l`
column with numeric-looking text), wrap negative numbers in `\num{-0.45}` so
siunitx still renders a true minus rather than a typographic hyphen.

## Multi-level headers

Grouped headers ("Accuracy" spanning "Train"/"Test" sub-columns) use
`\multicolumn` for the group label and `\cmidrule(lr){...}` under it — never
a full-width `\hline`, which would visually connect columns that aren't
actually related.

```latex
\begin{tabular}{lS[table-format=1.3]S[table-format=1.3]S[table-format=1.3]S[table-format=1.3]}
\toprule
Method & \multicolumn{2}{c}{Accuracy} & \multicolumn{2}{c}{F1 Score} \\
\cmidrule(lr){2-3}
\cmidrule(lr){4-5}
 & \multicolumn{1}{c}{Train} & \multicolumn{1}{c}{Test} & \multicolumn{1}{c}{Train} & \multicolumn{1}{c}{Test} \\
\midrule
Baseline & 0.812 & 0.795 & 0.803 & 0.781 \\
\bottomrule
\end{tabular}
```

The `(lr)` in `\cmidrule(lr){2-3}` trims the rule slightly at both ends so
adjacent group rules don't visually merge into one line — small detail,
noticeably cleaner in the compiled output.

## Wrapping long header text

Before shrinking the font (see the fitting order in SKILL.md), try wrapping
a long header across two lines — it's usually the cheaper fix and reads
better than an abbreviation the reader has to decode:

```latex
\makecell{Validation\\Accuracy}
```

## Table notes (threeparttable)

Footnote-style notes under a table (significance levels, data source,
abbreviation key) should use `threeparttable` rather than a `\\` + text line
directly under `\bottomrule` — the plain-text approach doesn't respect the
table's width and can run past the margin.

```latex
\begin{table}
\centering
\begin{threeparttable}
\caption{Model comparison.}
\begin{tabular}{lS[table-format=1.3]}
\toprule
Method & {Accuracy} \\
\midrule
Ours & 0.891\sym{**} \\
\bottomrule
\end{tabular}
\begin{tablenotes}
\small
\item[] \sym{**} $p < 0.01$, paired $t$-test against baseline.
\end{tablenotes}
\end{threeparttable}
\end{table}
```

## Significance stars in numeric columns

`0.452***` can't sit directly in an `S` column — siunitx's strict number
parser doesn't recognize trailing stars as part of the number. Strip the
stars from the numeric value, format it normally, then append the stars as a
superscript outside the number:

```latex
0.452\sym{***}
```

`render_body.py` does this automatically for any cell whose display value
ends in `*`, `**`, or `***`.

## Row-spanning stub columns (multirow)

When Excel has a merged cell spanning several rows in the leftmost column
(e.g. a variable name grouping several sub-rows):

```latex
\multirow{3}{*}{Group A} & row 1 values \\
                          & row 2 values \\
                          & row 3 values \\
```

`render_body.py` emits this automatically when `structure.json` has a merge
with `rowspan > 1` in a non-header row.

## Landscape / rotated tables

When a table is inherently wide (many columns) rather than just slightly
over budget, rotating the whole table is often more readable than shrinking
text to the point of being hard to read:

```latex
\begin{sidewaystable}
\centering
\caption{...}
\begin{tabular}{...}
...
\end{tabular}
\end{sidewaystable}
```

Requires `\usepackage{rotating}`.

## Resizebox — last resort only

```latex
\resizebox{\textwidth}{!}{%
\begin{tabular}{...}
...
\end{tabular}
}
```

This scales the *entire* table including fonts, so the table's numbers end
up a different size than the surrounding body text and any other tables in
the paper — reviewers notice this inconsistency. Only reach for it after
`\tabcolsep` shrinking, `\footnotesize`, header wrapping, and rotation have
all been tried and the table still doesn't fit (see the fitting order in
SKILL.md).

## Escaping characters from Excel text

Excel cell text often contains characters that are special in LaTeX:

| Character | Escape as |
|---|---|
| `%` | `\%` |
| `&` | `\&` |
| `_` | `\_` |
| `#` | `\#` |
| `$` | `\$` |
| `~` | `\textasciitilde{}` |
| `^` | `\textasciicircum{}` |

`render_body.py` escapes all of these automatically for text cells. Numeric
cells are never escaped (their `display` value from `extract_xlsx.py` is
already a clean formatted number).
