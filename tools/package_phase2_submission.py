#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


def _sanitize_filename(text: str) -> str:
    text = text.strip()
    # Remove common prefixes like "شکل ۴-۱:" and normalize separators for filenames.
    text = re.sub(r"^\s*شکل\s*[۰-۹0-9\-]+\s*:\s*", "", text)
    text = text.replace("—", "-")
    text = text.replace("–", "-")
    text = text.replace("(", " ")
    text = text.replace(")", " ")
    text = text.replace("+", " ")
    text = re.sub(r"[\\/:*?\"<>|]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    text = text.replace(" ", "_")
    text = re.sub(r"_+", "_", text)
    text = text.strip("._- \u200c\u200f")
    return text or "بدون_عنوان"


def _extract_fig_captions(py_path: Path) -> dict[str, str]:
    """
    Best-effort: map figure id like '2-1' to its caption title from generate_sad_final_docx.py.
    """
    if not py_path.exists():
        return {}
    text = py_path.read_text(encoding="utf-8", errors="ignore")
    caps: dict[str, str] = {}

    # Scan in order: whenever we see make_fig_marker("X"), the next fig_caption("...") is for it.
    marker_re = re.compile(r'make_fig_marker\("(?P<id>\d+-\d+)"\)')
    caption_re = re.compile(r'fig_caption\("(?P<cap>[^"]+)"\)')

    marker_iter = list(marker_re.finditer(text))
    if not marker_iter:
        return {}

    for m in marker_iter:
        fig_id = m.group("id")
        after = text[m.end() : m.end() + 800]  # small window
        cm = caption_re.search(after)
        if not cm:
            continue
        cap = cm.group("cap").strip()
        # Keep only the "title" part (after colon) when present.
        title = cap.split(":", 1)[1].strip() if ":" in cap else cap
        caps[fig_id] = title

    return caps


def _run_soffice_convert_to_pdf(docx_path: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    cmd = [
        "soffice",
        "--headless",
        "--nologo",
        "--nolockcheck",
        "--nodefault",
        "--norestore",
        "--invisible",
        "--convert-to",
        "pdf",
        str(docx_path),
        "--outdir",
        str(out_dir),
    ]
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    pdf_path = out_dir / (docx_path.stem + ".pdf")
    if not pdf_path.exists():
        raise RuntimeError("PDF conversion finished but output PDF was not found.")
    return pdf_path


def _iter_vp_diagrams(diagrams_dir: Path) -> list[Path]:
    if not diagrams_dir.exists():
        return []
    out: list[Path] = []
    for ext in ("png", "jpg", "jpeg", "pdf"):
        out.extend(sorted(diagrams_dir.glob(f"*-vp.{ext}")))
    return out


def _fig_id_from_filename(path: Path) -> str | None:
    # e.g. fig-2-1-context-vp.png -> 2-1
    m = re.match(r"^fig-(\d+-\d+)-", path.name)
    return m.group(1) if m else None


def _zip_dir(src_dir: Path, zip_path: Path) -> None:
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for p in sorted(src_dir.rglob("*")):
            if p.is_dir():
                continue
            z.write(p, arcname=str(p.relative_to(src_dir)))


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Package Phase-2 deliverables: convert SAD-Final.docx to PDF, rename VP diagrams, and build a zip."
        )
    )
    parser.add_argument("--docx", type=Path, default=Path("SAD-Final.docx"), help="Input .docx")
    parser.add_argument("--diagrams-dir", type=Path, default=Path("diagrams"), help="Diagrams directory")
    parser.add_argument("--py-source", type=Path, default=Path("generate_sad_final_docx.py"), help="Caption source")
    parser.add_argument("--student1", required=True, help="شماره دانشجویی نفر اول")
    parser.add_argument("--student2", required=True, help="شماره دانشجویی نفر دوم")
    parser.add_argument(
        "--doc-title",
        default="سند_معماری_نرم‌افزار",
        help="عنوان فایل PDF سند (بدون پسوند). پیش‌فرض: سند_معماری_نرم‌افزار",
    )
    parser.add_argument("--out-dir", type=Path, default=Path("dist/phase2"), help="Staging output directory")
    parser.add_argument("--zip", type=Path, default=Path("dist/phase2.zip"), help="Zip path")
    args = parser.parse_args()

    docx_path: Path = args.docx
    if not docx_path.exists():
        print(f"Missing docx: {docx_path}", file=sys.stderr)
        return 2

    captions = _extract_fig_captions(args.py_source)
    vp_diagrams = _iter_vp_diagrams(args.diagrams_dir)
    if not vp_diagrams:
        print(f"No VP diagrams found under: {args.diagrams_dir} (expected *-vp.png/jpg/pdf)", file=sys.stderr)
        return 3

    out_dir: Path = args.out_dir
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="sad-phase2-") as tmp:
        tmp_dir = Path(tmp)
        pdf_path = _run_soffice_convert_to_pdf(docx_path, tmp_dir)
        pdf_out_name = f"{args.student1}_{args.student2}_{_sanitize_filename(args.doc_title)}.pdf"
        shutil.copy2(pdf_path, out_dir / pdf_out_name)

    # Copy and rename VP diagrams.
    for p in vp_diagrams:
        fig_id = _fig_id_from_filename(p)
        title = captions.get(fig_id, p.stem)
        title = _sanitize_filename(title)
        out_name = f"{args.student1}_{args.student2}_{title}{p.suffix.lower()}"
        shutil.copy2(p, out_dir / out_name)

    _zip_dir(out_dir, args.zip)
    print(f"Wrote folder: {out_dir}")
    print(f"Wrote zip: {args.zip}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
