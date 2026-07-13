#!/usr/bin/env python3
"""
Compile a LaTeX table snippet inside the project's real preamble/class and
report whether it fits — numerically first (cheap, exact), and only render a
cropped preview image when asked (--final), since images cost far more
tokens than a width comparison.

Usage:
    python compile_preview.py --project .xlsx2latex/projects/mypaper \
        --body build/body.tex --table build/table.tex

    # after the numeric loop converges, get a preview image too:
    python compile_preview.py --project .xlsx2latex/projects/mypaper \
        --body build/body.tex --table build/table.tex --final
"""
import argparse
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

TYPEOUT_RE = re.compile(r"XLSX2LATEX::(\w+)::([\d.]+)pt")
OVERFULL_RE = re.compile(r"Overfull \\hbox \(([\d.]+)pt too wide\)")
UNDEF_CITE_RE = re.compile(r"Citation [`']([^']+)' .*undefined")
LATEX_ERROR_RE = re.compile(r"^! (.+)$", re.MULTILINE)


def load_config(project_dir: Path):
    cfg_path = project_dir / "config.json"
    if not cfg_path.exists():
        print(f"ERROR: no config.json in {project_dir}. Run project setup first "
              f"(see SKILL.md 'Project setup').", file=sys.stderr)
        sys.exit(1)
    return json.loads(cfg_path.read_text())


def find_engine(requested):
    if requested in ("latexmk", "pdflatex", "auto", None):
        if shutil.which("latexmk"):
            return "latexmk"
        if shutil.which("pdflatex"):
            return "pdflatex"
        if requested in ("latexmk", "pdflatex"):
            print(f"WARNING: requested engine '{requested}' not found, "
                  f"falling back to tectonic if available.", file=sys.stderr)
    if shutil.which("tectonic"):
        return "tectonic"
    return None


def build_wrapper(config, project_dir: Path, build_dir: Path, table_file: str, has_bib: bool):
    doc_class = config.get("document_class", "article")
    class_options = config.get("class_options", "")
    geometry = config.get("geometry")
    bib_backend = config.get("bib_backend", "none")

    lines = []
    opts = f"[{class_options}]" if class_options else ""
    lines.append(rf"\documentclass{opts}{{{doc_class}}}")
    if geometry:
        lines.append(rf"\usepackage[{geometry}]{{geometry}}")
    preamble_path = project_dir / "preamble.tex"
    preamble_text = preamble_path.read_text() if preamble_path.exists() else ""
    if preamble_text:
        lines.append(r"\input{preamble.tex}")
    # fall back so a bare-bones project (no preamble supplied yet, or one that's
    # missing table packages) still compiles
    if "booktabs" not in preamble_text:
        lines.append(r"\usepackage{booktabs,siunitx,multirow,threeparttable,makecell}")
    if r"\sym" not in preamble_text:
        lines.append(r"\providecommand{\sym}[1]{\textsuperscript{#1}}")
    lines.append(r"\begin{document}")
    lines.append(r"\pagestyle{empty}")
    lines.append(r"\newsavebox\xlsxtablebox")
    lines.append(r"\sbox{\xlsxtablebox}{\input{body.tex}}")
    lines.append(r"\typeout{XLSX2LATEX::TABLEWIDTH::\the\wd\xlsxtablebox}")
    lines.append(r"\typeout{XLSX2LATEX::TEXTWIDTH::\the\textwidth}")
    lines.append(r"\typeout{XLSX2LATEX::COLUMNWIDTH::\the\columnwidth}")
    lines.append(rf"\input{{{table_file}}}")
    if has_bib and bib_backend == "natbib":
        lines.append(r"\bibliographystyle{plainnat}")
        lines.append(r"\bibliography{references}")
    elif has_bib and bib_backend == "biblatex":
        lines.append(r"\printbibliography")
    lines.append(r"\end{document}")
    return "\n".join(lines)


def run(cmd, cwd):
    return subprocess.run(cmd, cwd=cwd, capture_output=True, text=True, timeout=120)


def compile_wrapper(engine, build_dir: Path, has_bib: bool, bib_backend: str):
    """Returns (success, combined_log_text)."""
    if engine == "latexmk":
        cmd = ["latexmk", "-pdf", "-interaction=nonstopmode", "-halt-on-error", "-f", "wrapper.tex"]
        r = run(cmd, build_dir)
        log = (build_dir / "wrapper.log")
        return (build_dir / "wrapper.pdf").exists(), (log.read_text(errors="replace") if log.exists() else r.stdout + r.stderr)

    if engine == "pdflatex":
        base_cmd = ["pdflatex", "-interaction=nonstopmode", "-halt-on-error", "wrapper.tex"]
        run(base_cmd, build_dir)
        if has_bib:
            if bib_backend == "natbib" and shutil.which("bibtex"):
                run(["bibtex", "wrapper"], build_dir)
            elif bib_backend == "biblatex" and shutil.which("biber"):
                run(["biber", "wrapper"], build_dir)
            run(base_cmd, build_dir)
        run(base_cmd, build_dir)  # second pass for cross-refs/width
        log = (build_dir / "wrapper.log")
        return (build_dir / "wrapper.pdf").exists(), (log.read_text(errors="replace") if log.exists() else "")

    if engine == "tectonic":
        cmd = ["tectonic", "--outdir", ".", "--keep-logs", "--print", "wrapper.tex"]
        r = run(cmd, build_dir)
        log = (build_dir / "wrapper.log")
        log_text = log.read_text(errors="replace") if log.exists() else (r.stdout + r.stderr)
        return (build_dir / "wrapper.pdf").exists(), log_text

    return False, "ERROR: no usable LaTeX engine found (checked latexmk, pdflatex, tectonic)."


def parse_log(log_text):
    findings = {"table_width": None, "text_width": None, "column_width": None,
                "overfull": [], "undefined_cites": [], "fatal_errors": []}
    for m in TYPEOUT_RE.finditer(log_text):
        key, val = m.group(1), float(m.group(2))
        if key == "TABLEWIDTH":
            findings["table_width"] = val
        elif key == "TEXTWIDTH":
            findings["text_width"] = val
        elif key == "COLUMNWIDTH":
            findings["column_width"] = val
    findings["overfull"] = [float(m.group(1)) for m in OVERFULL_RE.finditer(log_text)]
    findings["undefined_cites"] = sorted(set(UNDEF_CITE_RE.findall(log_text)))
    findings["fatal_errors"] = LATEX_ERROR_RE.findall(log_text)
    return findings


def summarize(success, findings, target_width="textwidth"):
    lines = []
    if not success:
        lines.append("COMPILE: FAILED")
        if findings["fatal_errors"]:
            lines.append("ERROR: " + findings["fatal_errors"][0])
        else:
            lines.append("ERROR: compile failed, no PDF produced (see full log for detail)")
        return "\n".join(lines)

    lines.append("COMPILE: success")
    tw = findings["table_width"]
    xw = findings["column_width"] if target_width == "columnwidth" else findings["text_width"]
    if tw is not None and xw is not None:
        delta = xw - tw
        status = "fits" if delta >= 0 else "OVERFLOWS"
        lines.append(f"WIDTH: table={tw:.1f}pt target={xw:.1f}pt [{target_width}] "
                      f"({status}, {abs(delta):.1f}pt {'to spare' if delta >= 0 else 'over'})")
    else:
        lines.append("WIDTH: could not measure (typeout markers not found in log)")

    if findings["overfull"]:
        worst = max(findings["overfull"])
        lines.append(f"OVERFULL: {len(findings['overfull'])} warning(s), worst {worst:.1f}pt too wide")
    else:
        lines.append("OVERFULL: none")

    if findings["undefined_cites"]:
        lines.append(f"CITATIONS: {len(findings['undefined_cites'])} undefined key(s): "
                      + ", ".join(findings["undefined_cites"])
                      + " -- check spelling against references.bib (undefined cites render "
                        "narrower than real ones and can produce a false 'fits')")
    return "\n".join(lines)


def make_preview_png(build_dir: Path, dpi=200):
    pdf = build_dir / "wrapper.pdf"
    if not pdf.exists() or not shutil.which("pdftoppm"):
        return None
    out_prefix = build_dir / "preview"
    run(["pdftoppm", "-png", "-r", str(dpi), str(pdf), str(out_prefix)], build_dir)
    pages = sorted(build_dir.glob("preview*.png"))
    if not pages:
        return None
    png_path = pages[0]
    try:
        from PIL import Image, ImageOps
        img = Image.open(png_path).convert("RGB")
        gray = ImageOps.grayscale(img)
        bbox = ImageOps.invert(gray).getbbox()
        if bbox:
            pad = 20
            l, t, r, b = bbox
            l, t = max(0, l - pad), max(0, t - pad)
            r, b = min(img.width, r + pad), min(img.height, b + pad)
            img = img.crop((l, t, r, b))
        img.save(png_path)
    except ImportError:
        print("NOTE: Pillow not installed, skipping autocrop (full page PNG saved instead).", file=sys.stderr)
    return png_path


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--project", required=True, help="Path to .xlsx2latex/projects/<name>")
    ap.add_argument("--body", required=True, help="Path to tabular body .tex (from render_body.py)")
    ap.add_argument("--table", required=True, help="Path to full table float .tex (\\input{}s the body)")
    ap.add_argument("--final", action="store_true", help="Also render a cropped preview PNG")
    ap.add_argument("--dpi", type=int, default=200)
    args = ap.parse_args()

    # resolve to absolute immediately: build_dir is used both as a subprocess cwd
    # and as a path argument to commands run in that cwd (e.g. pdftoppm), and a
    # relative build_dir makes those two uses double up into a bad path.
    project_dir = Path(args.project).resolve()
    config = load_config(project_dir)
    build_dir = project_dir / "build"
    build_dir.mkdir(parents=True, exist_ok=True)

    # stage inputs into the build dir so relative \input{} paths resolve
    shutil.copy(args.body, build_dir / "body.tex")
    shutil.copy(args.table, build_dir / "table.tex")
    if (project_dir / "preamble.tex").exists():
        shutil.copy(project_dir / "preamble.tex", build_dir / "preamble.tex")
    has_bib = (project_dir / "references.bib").exists()
    if has_bib:
        shutil.copy(project_dir / "references.bib", build_dir / "references.bib")

    wrapper = build_wrapper(config, project_dir, build_dir, "table.tex", has_bib)
    (build_dir / "wrapper.tex").write_text(wrapper)

    engine = find_engine(config.get("engine", "auto"))
    if engine is None:
        print("COMPILE: FAILED\nERROR: no LaTeX engine found. Install one: "
              "brew install --cask basictex (pdflatex, matches Overleaf) or brew install tectonic.",
              file=sys.stderr)
        sys.exit(1)

    success, log_text = compile_wrapper(engine, build_dir, has_bib, config.get("bib_backend", "none"))
    findings = parse_log(log_text)
    print(summarize(success, findings, config.get("target_width", "textwidth")))
    print(f"ENGINE: {engine}", file=sys.stderr)

    if args.final and success:
        png_path = make_preview_png(build_dir, args.dpi)
        if png_path:
            print(f"PREVIEW: {png_path}")
        else:
            print("PREVIEW: failed to render (pdftoppm not found or no output page)", file=sys.stderr)

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
