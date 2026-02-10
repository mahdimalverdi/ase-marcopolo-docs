#!/usr/bin/env python3
from __future__ import annotations

import argparse
import html as html_lib
import shutil
import subprocess
import tempfile
from pathlib import Path

from PIL import Image


def find_mermaid_js() -> Path:
    candidates = [
        Path.home()
        / ".vscode/extensions/shd101wyy.markdown-preview-enhanced-0.8.20/crossnote/dependencies/mermaid/mermaid.min.js",
        Path.home() / ".vscode/extensions/hediet.vscode-drawio-1.9.0/drawio/src/main/webapp/js/mermaid/mermaid.min.js",
    ]

    for p in candidates:
        if p.exists():
            return p

    roots = [Path.home() / ".vscode/extensions", Path.home() / ".vscode-server/extensions"]
    for root in roots:
        if not root.exists():
            continue
        for p in root.rglob("mermaid.min.js"):
            return p

    raise FileNotFoundError(
        "mermaid.min.js پیدا نشد. یک افزونه مثل Markdown Preview Enhanced یا Draw.io را در VS Code نصب کنید."
    )


def crop_whitespace(img: Image.Image, padding: int = 24) -> Image.Image:
    rgb = img.convert("RGB")
    bg = (255, 255, 255)

    # Find bounding box of non-white pixels.
    bbox = None
    pix = rgb.load()
    w, h = rgb.size
    x_min, y_min, x_max, y_max = w, h, 0, 0
    found = False
    for y in range(h):
        for x in range(w):
            if pix[x, y] != bg:
                x_min = min(x_min, x)
                y_min = min(y_min, y)
                x_max = max(x_max, x)
                y_max = max(y_max, y)
                found = True
    if not found:
        return img

    x_min = max(0, x_min - padding)
    y_min = max(0, y_min - padding)
    x_max = min(w - 1, x_max + padding)
    y_max = min(h - 1, y_max + padding)
    bbox = (x_min, y_min, x_max + 1, y_max + 1)
    return img.crop(bbox)


def render_one(
    *,
    chrome: Path,
    mermaid_js: Path,
    src: Path,
    out_png: Path,
    width: int,
    height: int,
    time_budget_ms: int,
) -> None:
    code = src.read_text(encoding="utf-8")
    # Mermaid code is placed into HTML; escape it so tokens like "<<interface>>" are not treated as tags.
    code_html = html_lib.escape(code)

    html_doc = f"""<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <style>
      html, body {{
        margin: 0;
        padding: 0;
        background: #fff;
      }}
      .wrap {{
        padding: 40px;
      }}
      .mermaid {{
        font-family: DejaVu Sans, Arial, sans-serif;
      }}
    </style>
  </head>
  <body>
    <div class="wrap">
      <div class="mermaid">
{code_html}
      </div>
    </div>
    <script src="{mermaid_js.as_uri()}"></script>
    <script>
      mermaid.initialize({{
        startOnLoad: true,
        securityLevel: "strict",
        theme: "default"
      }});
    </script>
  </body>
</html>
"""

    out_png.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="mermaid-render-") as td:
        td_path = Path(td)
        html_path = td_path / "diagram.html"
        raw_png = td_path / "raw.png"
        html_path.write_text(html_doc, encoding="utf-8")

        cmd = [
            str(chrome),
            "--headless=new",
            "--no-sandbox",
            "--disable-gpu",
            "--disable-dev-shm-usage",
            "--no-first-run",
            "--no-default-browser-check",
            "--hide-scrollbars",
            f"--window-size={width},{height}",
            f"--virtual-time-budget={time_budget_ms}",
            f"--screenshot={raw_png}",
            str(html_path.as_uri()),
        ]
        # Retry with older headless mode if Chrome crashes (SIGTRAP happens intermittently in some sandboxes).
        for attempt in range(2):
            try:
                subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                break
            except subprocess.CalledProcessError:
                if attempt == 0:
                    cmd[cmd.index("--headless=new")] = "--headless"
                    continue
                raise

        with Image.open(raw_png) as im:
            im = im.convert("RGB")
            im = crop_whitespace(im, padding=28)

            # Ensure minimum width for readability.
            if im.size[0] < 1600:
                scale = 1600 / max(im.size[0], 1)
                new_size = (1600, int(im.size[1] * scale))
                im = im.resize(new_size, Image.Resampling.LANCZOS)

            im.save(out_png, format="PNG", optimize=True)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--src-dir", default="diagrams/mermaid", help="Folder containing .mmd files")
    parser.add_argument("--out-dir", default="diagrams", help="Output folder for .png files")
    parser.add_argument("--width", type=int, default=2200, help="Chrome viewport width")
    parser.add_argument("--height", type=int, default=2000, help="Chrome viewport height")
    parser.add_argument("--time-budget-ms", type=int, default=5000, help="Time budget to let Mermaid render")
    args = parser.parse_args()

    chrome = shutil.which("google-chrome") or shutil.which("chromium") or shutil.which("chromium-browser")
    if not chrome:
        raise SystemExit("مرورگر Chrome/Chromium پیدا نشد.")

    mermaid_js = find_mermaid_js()

    src_dir = Path(args.src_dir)
    out_dir = Path(args.out_dir)
    if not src_dir.exists():
        raise SystemExit(f"مسیر ورودی پیدا نشد: {src_dir}")

    mmd_files = sorted(src_dir.glob("*.mmd"))
    if not mmd_files:
        raise SystemExit(f"هیچ فایل .mmd در این مسیر نیست: {src_dir}")

    for src in mmd_files:
        out_png = out_dir / (src.stem + ".png")
        render_one(
            chrome=Path(chrome),
            mermaid_js=mermaid_js,
            src=src,
            out_png=out_png,
            width=args.width,
            height=args.height,
            time_budget_ms=args.time_budget_ms,
        )
        print(f"OK: {src.name} -> {out_png}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
