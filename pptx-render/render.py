#!/usr/bin/env python3
"""Render a .pptx deck to per-slide PNGs (and optionally a PDF) for visual QA.

Engine order: installed PowerPoint (COM, best fidelity -- real fonts, real
layout), then LibreOffice (+ PyMuPDF for rasterising). Nothing is faked: if no
engine is available the script exits non-zero and says so.

    python render.py deck.pptx                 # -> render/slide-01.png ...
    python render.py deck.pptx -o qa --pdf     # also qa/deck.pdf
    python render.py deck.pptx --slides 1,4-6  # subset only
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

POWERPOINT_CANDIDATES = [
    r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
    r"C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE",
    r"C:\Program Files (x86)\Microsoft Office\Office16\POWERPNT.EXE",
    r"C:\Program Files\Microsoft Office\Office15\POWERPNT.EXE",
]
SOFFICE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]
TRAILING_NUMBER_RE = re.compile(r"(\d+)(?!.*\d)")


def die(msg: str, code: int = 2):
    print(f"render.py: {msg}", file=sys.stderr)
    raise SystemExit(code)


def parse_slides(spec: str | None, total: int) -> list[int]:
    """'1,4-6' -> [1,4,5,6]; None -> every slide."""
    if not spec:
        return list(range(1, total + 1))
    out: list[int] = []
    for chunk in spec.replace(" ", "").split(","):
        if not chunk:
            continue
        if "-" in chunk:
            a, b = chunk.split("-", 1)
            out.extend(range(int(a), int(b) + 1))
        else:
            out.append(int(chunk))
    bad = [n for n in out if n < 1 or n > total]
    if bad:
        die(f"slides out of range 1..{total}: {bad}")
    return sorted(dict.fromkeys(out))


def find_powerpoint() -> str | None:
    if sys.platform != "win32":
        return None
    for p in POWERPOINT_CANDIDATES:
        if Path(p).is_file():
            return p
    try:
        import winreg

        for key in (
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE",
            r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE",
        ):
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key) as k:
                    path = winreg.QueryValue(k, None)
                if path and Path(path).is_file():
                    return path
            except OSError:
                continue
    except ImportError:
        pass
    return None


def find_soffice() -> str | None:
    which = shutil.which("soffice") or shutil.which("libreoffice")
    if which:
        return which
    for p in SOFFICE_CANDIDATES:
        if Path(p).is_file():
            return p
    return None


def collect_exports(raw_dir: Path, outdir: Path, wanted: list[int]) -> list[Path]:
    """PowerPoint names exports per locale ('Slide1.PNG', 'Слайд1.PNG'). Rename to
    slide-NN.png in slide order and keep only the wanted slides."""
    found: list[tuple[int, Path]] = []
    for f in raw_dir.iterdir():
        if not f.is_file() or f.suffix.lower() not in {".png", ".jpg", ".jpeg"}:
            continue
        m = TRAILING_NUMBER_RE.search(f.stem)
        if m:
            found.append((int(m.group(1)), f))
    found.sort()
    written: list[Path] = []
    for num, src in found:
        if num not in wanted:
            continue
        dst = outdir / f"slide-{num:02d}.png"
        shutil.move(str(src), dst)
        written.append(dst)
    return sorted(written)


def render_powerpoint(deck: Path, outdir: Path, width: int | None, spec: str | None, want_pdf: bool):
    try:
        import win32com.client as win32  # type: ignore
    except ImportError:
        die("pywin32 missing: python -m pip install pywin32", 3)

    pdf_path = outdir / (deck.stem + ".pdf") if want_pdf else None
    app = pres = None
    raw = Path(tempfile.mkdtemp(prefix="pptx-render-"))
    try:
        # DispatchEx avoids attaching to a PowerPoint the user already has open.
        app = win32.DispatchEx("PowerPoint.Application")
        pres = app.Presentations.Open(str(deck.resolve()), True, False, False)
        total = int(pres.Slides.Count)
        wanted = parse_slides(spec, total)
        setup = pres.PageSetup
        w = int(width or 1920)
        h = int(round(w * float(setup.SlideHeight) / float(setup.SlideWidth)))

        if pdf_path is not None:
            if pdf_path.exists():
                pdf_path.unlink()
            pres.SaveAs(str(pdf_path.resolve()), 32)  # ppSaveAsPDF

        if len(wanted) == total:
            pres.Export(str(raw), "PNG", w, h)
        else:
            for num in wanted:
                pres.Slides(num).Export(str(raw / f"slide{num}.png"), "PNG", w, h)
    finally:
        for closer in (lambda: pres.Close() if pres is not None else None,
                       lambda: app.Quit() if app is not None else None):
            try:
                closer()
            except Exception:
                pass

    written = collect_exports(raw, outdir, wanted)
    shutil.rmtree(raw, ignore_errors=True)
    return written, pdf_path, total, wanted


def render_libreoffice(exe: str, deck: Path, outdir: Path, spec: str | None, want_pdf: bool):
    with tempfile.TemporaryDirectory(prefix="pptx-render-") as tmp:
        subprocess.run(
            [exe, "--headless", "--norestore", "--convert-to", "pdf", "--outdir", tmp, str(deck.resolve())],
            check=True, capture_output=True, timeout=300,
        )
        produced = list(Path(tmp).glob("*.pdf"))
        if not produced:
            die("LibreOffice produced no PDF", 4)
        final_pdf = outdir / (deck.stem + ".pdf")
        if want_pdf:
            shutil.move(str(produced[0]), final_pdf)
        raster_src = final_pdf if want_pdf else produced[0]
        try:
            import fitz  # PyMuPDF
        except ImportError:
            if final_pdf:
                print(f"render.py: wrote {final_pdf}; PyMuPDF missing for PNGs "
                      f"(python -m pip install pymupdf)", file=sys.stderr)
                return [], final_pdf, 0, []
            die("PyMuPDF missing for PNG rasterising: python -m pip install pymupdf", 3)
        doc = fitz.open(raster_src)
        total = doc.page_count
        wanted = parse_slides(spec, total)
        written = []
        for idx in wanted:
            png = outdir / f"slide-{idx:02d}.png"
            doc[idx - 1].get_pixmap(dpi=150).save(png)
            written.append(png)
        doc.close()
    return written, (final_pdf if want_pdf else None), total, wanted


def main() -> int:
    ap = argparse.ArgumentParser(description="Render .pptx slides to PNG (and PDF) for visual review.")
    ap.add_argument("deck", help="path to .pptx/.potx file")
    ap.add_argument("-o", "--outdir", default="render", help="output directory (default: render)")
    ap.add_argument("--width", type=int, default=None, help="PNG width in px (default 1920)")
    ap.add_argument("--slides", default=None, help="subset, e.g. 1,3-5")
    ap.add_argument("--pdf", action="store_true", help="also save <deck>.pdf")
    ap.add_argument("--engine", choices=["auto", "powerpoint", "libreoffice"], default="auto")
    ap.add_argument("--keep-old", action="store_true", help="keep previous slide-*.png in outdir")
    args = ap.parse_args()

    deck = Path(args.deck)
    if not deck.is_file():
        die(f"no such file: {deck}")
    if deck.suffix.lower() not in {".pptx", ".potx", ".ppt"}:
        die(f"not a PowerPoint file: {deck.suffix}")
    if deck.suffix.lower() == ".ppt":
        die("legacy .ppt: convert to .pptx first")

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    if not args.keep_old:
        for old in outdir.glob("slide-*.png"):
            old.unlink()

    ppt = find_powerpoint() if args.engine in {"auto", "powerpoint"} else None
    soffice = find_soffice() if args.engine in {"auto", "libreoffice"} else None
    if args.engine == "powerpoint" and not ppt:
        die("PowerPoint not found", 3)
    if args.engine == "libreoffice" and not soffice:
        die("LibreOffice not found", 3)
    if not ppt and not soffice:
        die("no renderer available: install PowerPoint or LibreOffice", 3)

    written: list[Path] = []
    pdf = None
    total = 0
    wanted: list[int] = []
    if ppt:
        try:
            written, pdf, total, wanted = render_powerpoint(deck, outdir, args.width, args.slides, args.pdf)
        except SystemExit:
            raise
        except Exception as exc:  # COM errors are opaque; fall back once
            if soffice and args.engine == "auto":
                print(f"render.py: PowerPoint failed ({exc}); trying LibreOffice", file=sys.stderr)
                written, pdf, total, wanted = render_libreoffice(soffice, deck, outdir, args.slides, args.pdf)
            else:
                die(f"PowerPoint export failed: {exc}", 4)
    elif soffice:
        written, pdf, total, wanted = render_libreoffice(soffice, deck, outdir, args.slides, args.pdf)

    if total and len(written) != len(wanted):
        print(f"render.py: WARNING expected {len(wanted)} PNGs, got {len(written)}", file=sys.stderr)
    for p in written:
        print(p.resolve())
    if pdf:
        print(pdf.resolve())
    if not written:
        die("no images produced", 4)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
