#!/usr/bin/env python3
"""
Turn extracted table structure + a small layout spec into the LaTeX tabular
*body* (column format string, header rows, data rows, rules) deterministically.

This exists so the model never retypes data values by hand between iterations.
Only `layout.json` (column types, spacing, fitting strategy, header text) is
something the model edits; the numbers always flow straight from the
extracted xlsx JSON.

Usage:
    python render_body.py structure.json layout.json > body.tex
    python render_body.py structure.json layout.json --out body.tex
"""
import argparse
import json
import re
import sys

LATEX_ESCAPE = {
    "&": r"\&", "%": r"\%", "$": r"\$", "#": r"\#",
    "_": r"\_", "{": r"\{", "}": r"\}", "~": r"\textasciitilde{}",
    "^": r"\textasciicircum{}",
}
STAR_RE = re.compile(r"^(.*?)(\*{1,3})$")


def escape_text(s):
    s = str(s)
    out = []
    for ch in s:
        out.append(LATEX_ESCAPE.get(ch, ch))
    return "".join(out)


def strip_stars(display):
    """Split trailing significance stars ('0.452***') from the numeric part."""
    m = STAR_RE.match(display)
    if m:
        return m.group(1), m.group(2)
    return display, None


def cell_lookup(row, col):
    for entry in row["cells"]:
        if entry["col"] == col:
            return entry
    return None


def merge_at(merges, row0, col0):
    for m in merges:
        if m["row0"] == row0 and m["col0"] == col0:
            return m
    return None


def covered_by_merge(merges, row0, col0):
    for m in merges:
        if m["row0"] <= row0 <= m["row1"] and m["col0"] <= col0 <= m["col1"]:
            if (row0, col0) != (m["row0"], m["col0"]):
                return True
    return False


def render_cell_content(entry, col_spec, use_num_wrapper):
    """Render one cell's LaTeX content given its data and the column's type."""
    if entry is None:
        return ""
    col_type = col_spec.get("type", "l") if col_spec else "l"
    if entry.get("numeric"):
        display = entry["display"]
        base, stars = strip_stars(display)
        if col_type == "S":
            content = base  # siunitx S columns take the plain formatted number as-is
        elif use_num_wrapper:
            content = r"\num{%s}" % base
        else:
            content = base
        if stars:
            content += r"\sym{%s}" % stars
    else:
        content = escape_text(entry["value"])

    if entry.get("italic"):
        content = r"\textit{%s}" % content
    return content


def build_column_format(layout, structure):
    non_spacer = [c["index"] for c in structure["columns"] if not c["is_spacer"]]
    cols = layout["columns"]
    if len(cols) != len(non_spacer):
        print(f"WARNING: layout has {len(cols)} columns but structure has "
              f"{len(non_spacer)} non-spacer columns. Alignment may be wrong.",
              file=sys.stderr)
    parts = []
    for c in cols:
        t = c.get("type", "l")
        if t == "S" and c.get("table_format"):
            parts.append(f"S[table-format={c['table_format']}]")
        else:
            parts.append(t)
    sep = layout.get("tabcolsep")
    colformat = " ".join(parts)
    return colformat, non_spacer


def render_header_rows(layout):
    """Header text/spans come from the layout spec (abbreviation is a judgment
    call), rendered here so escaping/multicolumn syntax stays consistent.
    cmidrules are interleaved right after the header row they annotate
    (each entry's "row" is the 0-indexed position in header_rows), matching
    the standard \\multicolumn header / \\cmidrule / sub-header pattern.

    Every header cell that sits over an S column is wrapped in \\multicolumn{1}{c}{...}
    even when it doesn't span multiple columns: siunitx's S column tries to parse
    *everything* in that column as a number, so plain text like "Train" left bare
    would make siunitx choke trying to parse it as one."""
    lines = []
    header_rows = layout.get("header_rows", [])
    cmidrules_by_row = {}
    for cm in layout.get("cmidrules", []):
        cmidrules_by_row.setdefault(cm["row"], []).append(cm["span"])
    col_types = [c.get("type", "l") for c in layout["columns"]]

    for i, hrow in enumerate(header_rows):
        cells = []
        col_pos = 0  # tracks which physical column the next cell starts at
        for item in hrow:
            if isinstance(item, dict):
                span = item.get("span", 1)
                align = item.get("align", "c")
                text = item.get("text", "")
            else:
                span = 1
                align = "c"
                text = item if item else ""

            spans_S = any(col_types[c] == "S" for c in range(col_pos, min(col_pos + span, len(col_types))))
            if span > 1 or (spans_S and text):
                cells.append(r"\multicolumn{%d}{%s}{%s}" % (span, align, text))
            else:
                cells.append(text)
            col_pos += span
        lines.append(" & ".join(cells) + r" \\")
        for span in cmidrules_by_row.get(i, []):
            lines.append(r"\cmidrule(lr){%s}" % span)
    return lines


def render_data_rows(structure, layout, non_spacer_cols):
    col_by_index = {ci: layout["columns"][i] for i, ci in enumerate(non_spacer_cols)}
    use_num_wrapper = layout.get("num_wrapper_outside_S", True)
    header_row_count = layout.get("header_row_count", 0)
    bold_rows = set(structure["style_notes"].get("bold_rows", []))
    midrule_after = set(layout.get("midrule_after_rows", []))
    merges = structure["merges"]

    lines = []
    for row in structure["rows"]:
        r = row["index"]
        if r < header_row_count:
            continue  # header rows come from layout, not re-derived here
        cells_out = []
        for ci in non_spacer_cols:
            if covered_by_merge(merges, r, ci):
                continue
            m = merge_at(merges, r, ci)
            entry = cell_lookup(row, ci)
            content = render_cell_content(entry, col_by_index.get(ci), use_num_wrapper)
            if m and m["colspan"] > 1:
                align = col_by_index.get(ci, {}).get("type", "l")
                content = r"\multicolumn{%d}{%s}{%s}" % (m["colspan"], align, content)
            cells_out.append(content)
        prefix = r"\bfseries " if r in bold_rows else ""
        if prefix and cells_out:
            cells_out[0] = prefix + cells_out[0]
        lines.append(" & ".join(cells_out) + r" \\")
        if r in midrule_after:
            lines.append(r"\midrule")
    return lines


def render(structure, layout):
    colformat, non_spacer_cols = build_column_format(layout, structure)
    out = []
    booktabs = layout.get("booktabs", {"toprule": True, "midrule_after_header": True, "bottomrule": True})

    out.append(r"\begin{tabular}{%s}" % colformat)
    if booktabs.get("toprule", True):
        out.append(r"\toprule")
    out.extend(render_header_rows(layout))
    if booktabs.get("midrule_after_header", True) and layout.get("header_rows"):
        out.append(r"\midrule")
    out.extend(render_data_rows(structure, layout, non_spacer_cols))
    if booktabs.get("bottomrule", True):
        out.append(r"\bottomrule")
    out.append(r"\end{tabular}")
    return "\n".join(out)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("structure_json", help="Output of extract_xlsx.py")
    ap.add_argument("layout_json", help="Small layout spec (column types, header text, fitting strategy)")
    ap.add_argument("--out", default=None, help="Write .tex here instead of stdout")
    args = ap.parse_args()

    with open(args.structure_json) as f:
        structure = json.load(f)
    with open(args.layout_json) as f:
        layout = json.load(f)

    body = render(structure, layout)
    if args.out:
        with open(args.out, "w") as f:
            f.write(body + "\n")
        print(f"Wrote {args.out}", file=sys.stderr)
    else:
        print(body)


if __name__ == "__main__":
    main()
