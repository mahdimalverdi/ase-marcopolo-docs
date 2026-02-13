#!/usr/bin/env python3
# Generates a filled SAD .docx based on SAD-Template.docx and SAD.md content.

from __future__ import annotations

import datetime as _dt
import copy
import re
import zipfile
import argparse
from pathlib import Path
from typing import Final
import xml.etree.ElementTree as ET

from PIL import Image, ImageDraw, ImageFont


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W_NS}

REL_NS: Final = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS: Final = "http://schemas.openxmlformats.org/package/2006/content-types"
R_NS: Final = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS: Final = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS: Final = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS: Final = "http://schemas.openxmlformats.org/drawingml/2006/picture"

EMU_PER_INCH: Final = 914400
DEFAULT_IMAGE_DPI: Final = 96

FIGURE_MARKER_PREFIX: Final = "[FIG:"

# Make XML output use stable, conventional prefixes (LibreOffice is sometimes picky).
ET.register_namespace("w", W_NS)
ET.register_namespace("wp", WP_NS)
ET.register_namespace("a", A_NS)
ET.register_namespace("pic", PIC_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("ct", CT_NS)
ET.register_namespace("rel", REL_NS)

# Header/footer metadata defaults (adjust if needed)
SYSTEM_NAME: Final = "مارکوپولو"
DOC_VERSION: Final = "1.0"
DOC_CLASSIFICATION: Final = "محرمانه"
GROUP_MEMBERS_FALLBACK: Final = "محمد صادقی، مهدی مالوردی"

_PERSIAN_DIGITS_TRANS = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")


def _to_persian_digits(text: str) -> str:
    return text.translate(_PERSIAN_DIGITS_TRANS)


def _read_group_members_from_markdown(path: Path) -> str | None:
    if not path.exists():
        return None
    try:
        text = path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return None
    m = re.search(r"^\s*-\s*اعضای گروه:\s*(.+?)\s*$", text, flags=re.M)
    if not m:
        return None
    val = m.group(1).strip()
    return val or None


def get_group_members() -> str:
    return (
        _read_group_members_from_markdown(Path("SAD.md"))
        or _read_group_members_from_markdown(Path("SAD-From-Template.md"))
        or GROUP_MEMBERS_FALLBACK
    )


def fill_cover_group_members(root: ET.Element, members: str) -> None:
    """
    Fill the (usually empty) paragraphs under 'نام اعضای گروه:' on the cover page.
    Tries to keep the template formatting by reusing run properties from the label line.
    """
    body = root.find("w:body", NS)
    if body is None:
        return
    paras = list(body.findall("w:p", NS))

    def _ensure_run_with_text(p: ET.Element, text: str, *, rpr_source: ET.Element | None) -> None:
        r = p.find("w:r", NS)
        if r is None:
            r = ET.SubElement(p, _qn("w:r"))
        if rpr_source is not None and r.find("w:rPr", NS) is None:
            r.append(copy.deepcopy(rpr_source))
        t = r.find("w:t", NS)
        if t is None:
            t = ET.SubElement(r, _qn("w:t"))
        t.text = text

    for i, p in enumerate(paras):
        if _p_text(p) != "نام اعضای گروه:":
            continue

        label_run = p.find("w:r", NS)
        label_rpr = label_run.find("w:rPr", NS) if label_run is not None else None

        for j in range(i + 1, min(i + 20, len(paras))):
            pj = paras[j]
            if _p_text(pj):
                break
            # Pick the first empty paragraph that has cover-like formatting (centered).
            jc = pj.find("w:pPr/w:jc", NS)
            if jc is not None and jc.attrib.get(_qns(W_NS, "val")) == "center":
                _ensure_run_with_text(pj, members, rpr_source=label_rpr)
                return
        return


def _gregorian_to_jalali(gy: int, gm: int, gd: int) -> tuple[int, int, int]:
    """
    Convert Gregorian date to Jalali (Solar Hijri).
    Implementation based on the common arithmetic conversion algorithm.
    """
    g_d_m = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334]
    gy2 = gy + 1 if gm > 2 else gy
    days = (
        355666
        + 365 * gy
        + (gy2 + 3) // 4
        - (gy2 + 99) // 100
        + (gy2 + 399) // 400
        + gd
        + g_d_m[gm - 1]
    )
    jy = -1595 + 33 * (days // 12053)
    days %= 12053
    jy += 4 * (days // 1461)
    days %= 1461
    if days > 365:
        jy += (days - 1) // 365
        days = (days - 1) % 365
    if days < 186:
        jm = 1 + days // 31
        jd = 1 + days % 31
    else:
        jm = 7 + (days - 186) // 30
        jd = 1 + (days - 186) % 30
    return jy, jm, jd


def _update_header_footer_xml(file_bytes: dict[str, bytes]) -> None:
    """
    Update header/footer placeholders so the output doesn't keep template '...'.
    Keeps the existing layout and edits only w:t text nodes.
    """
    today = _dt.date.today()
    jy, jm, jd = _gregorian_to_jalali(today.year, today.month, today.day)
    jalali_day = _to_persian_digits(f"{jd:02d}")
    jalali_month = _to_persian_digits(f"{jm:02d}")
    jy_str = str(jy)
    jalali_year = _to_persian_digits(jy_str)

    def _patch_part(part_path: str, edits: dict[int, str]) -> None:
        raw = file_bytes.get(part_path)
        if raw is None:
            return
        try:
            root = ET.fromstring(raw)
        except Exception:
            return
        nodes = root.findall(".//w:t", NS)
        for idx, val in edits.items():
            if 0 <= idx < len(nodes):
                nodes[idx].text = val
        file_bytes[part_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    # word/header1.xml tokens (by index):
    # 0 'سامانه ' | 1 '...' | 2 'نسخه 1.0' | ... | 7 'تاريخ: ' | 8 dd | 9 '/' | 10 mm | 11 '/140' | 12 '3'
    header_year_prefix = _to_persian_digits(f"/{jy_str[:-1]}" if len(jy_str) == 4 else f"/{jy_str}")
    header_year_last = _to_persian_digits(jy_str[-1] if len(jy_str) == 4 else "")
    _patch_part(
        "word/header1.xml",
        {
            1: SYSTEM_NAME,
            2: f"نسخه {DOC_VERSION}",
            8: jalali_day,
            10: jalali_month,
            11: header_year_prefix,
            12: header_year_last,
        },
    )

    # word/footer2.xml tokens (by index):
    # 0 'محرمانه' | 3 '2025' | 4 '، سامانه ...' | 5 'صفحه ' | 6 PAGE | 8 NUMPAGES
    _patch_part(
        "word/footer2.xml",
        {
            0: DOC_CLASSIFICATION,
            3: "\u200f" + jalali_year,
            4: f"، سامانه {SYSTEM_NAME}",
        },
    )


def _qns(uri: str, local: str) -> str:
    return f"{{{uri}}}{local}"


def _load_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    # Persian shaping is not guaranteed in PIL; keep labels short and Latin-friendly by default.
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSansCondensed.ttf",
    ]
    for p in candidates:
        try:
            return ImageFont.truetype(p, size=size)
        except Exception:
            continue
    return ImageFont.load_default()


def _ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def make_fig_marker(fig_id: str) -> ET.Element:
    # This paragraph is replaced with an image at build time (or removed if the image is missing).
    return make_p(f"{FIGURE_MARKER_PREFIX}{fig_id}]", jc="center", spacing_before=0, spacing_after=0)


def _qn(tag: str) -> str:
    # Qualified name for WordprocessingML tags (e.g., w:p)
    prefix, local = tag.split(":", 1)
    if prefix != "w":
        raise ValueError(f"Unsupported prefix: {prefix}")
    return f"{{{W_NS}}}{local}"


def _p_text(p: ET.Element) -> str:
    return "".join([t.text for t in p.findall(".//w:t", NS) if t.text]).strip()


def _set_run_text(r: ET.Element, text: str) -> None:
    text = sanitize_text(text)
    t = ET.SubElement(r, _qn("w:t"))
    if text.startswith(" ") or text.endswith(" "):
        t.set(f"{{{XML_NS}}}space", "preserve")
    t.text = text


_REMOVE_CHARS_RE = re.compile(
    "["  # keep ZWNJ (\u200c) but remove other invisible/control marks commonly introduced
    "\u200b"  # ZWSP
    "\u200d"  # ZWJ
    "\u200e"  # LRM
    "\u200f"  # RLM
    "\ufeff"  # BOM
    "\u2066\u2067\u2068\u2069"  # LRI/RLI/FSI/PDI
    "\u202a\u202b\u202c\u202d\u202e"  # bidi embeddings/overrides
    "]"
)


def sanitize_text(text: str) -> str:
    # Remove unwanted invisible/control chars (keep Persian half-space ZWNJ \u200c).
    text = _REMOVE_CHARS_RE.sub("", text)
    # Normalize spaces
    text = text.replace("\u00a0", " ")
    # Replace punctuation that often looks "auto-generated" to plain equivalents.
    text = text.replace("…", "...")
    text = text.replace("—", "-").replace("–", "-").replace("−", "-")
    # Avoid guillemets in final doc (prefer plain text).
    text = text.replace("«", "").replace("»", "")
    # Normalize Arabic variants to Persian forms.
    text = text.replace("ي", "ی").replace("ك", "ک")
    # Collapse excessive spaces (but keep ZWNJ intact).
    text = re.sub(r"[ \t]{2,}", " ", text)
    return text.strip()


def _add_rtl_props(parent: ET.Element, bold: bool = False, italic: bool = False, color: str | None = None) -> None:
    rPr = ET.SubElement(parent, _qn("w:rPr"))
    ET.SubElement(rPr, _qn("w:rFonts"), {_qn("w:cs"): "B Nazanin"})
    if bold:
        ET.SubElement(rPr, _qn("w:b"))
        ET.SubElement(rPr, _qn("w:bCs"))
    if italic:
        ET.SubElement(rPr, _qn("w:i"))
        ET.SubElement(rPr, _qn("w:iCs"))
    if color:
        ET.SubElement(rPr, _qn("w:color"), {_qn("w:val"): color})
    ET.SubElement(rPr, _qn("w:rtl"))
    ET.SubElement(rPr, _qn("w:lang"), {_qn("w:bidi"): "fa-IR"})


def make_p(
    text: str = "",
    *,
    style: str | None = None,
    bold: bool = False,
    italic: bool = False,
    jc: str = "lowKashida",
    color: str | None = None,
    spacing_before: int = 100,
    spacing_after: int = 100,
    keep_next: bool = False,
) -> ET.Element:
    p = ET.Element(_qn("w:p"))
    pPr = ET.SubElement(p, _qn("w:pPr"))
    if style:
        ET.SubElement(pPr, _qn("w:pStyle"), {_qn("w:val"): style})
    ET.SubElement(
        pPr,
        _qn("w:spacing"),
        {
            _qn("w:before"): str(spacing_before),
            _qn("w:beforeAutospacing"): "1",
            _qn("w:after"): str(spacing_after),
            _qn("w:afterAutospacing"): "1",
            _qn("w:line"): "360",
            _qn("w:lineRule"): "auto",
        },
    )
    if keep_next:
        ET.SubElement(pPr, _qn("w:keepNext"))
    ET.SubElement(pPr, _qn("w:jc"), {_qn("w:val"): jc})
    _add_rtl_props(pPr, bold=False, italic=False)

    if text:
        r = ET.SubElement(p, _qn("w:r"))
        _add_rtl_props(r, bold=bold, italic=italic, color=color)
        _set_run_text(r, text)
    return p


def make_label(text: str) -> ET.Element:
    """
    A bold label paragraph used inside use-case specs (e.g., کنشگرها/پیش‌شرط‌ها/پس‌شرط‌ها).
    Keeps it visually distinct from body text without turning it into a numbered heading/TOC item.
    """
    return make_p(text, bold=True, spacing_before=200, spacing_after=80, keep_next=True)


def make_tbl(headers: list[str], rows: list[list[str]], *, col_weights: list[int] | None = None) -> ET.Element:
    tbl = ET.Element(_qn("w:tbl"))

    tblPr = ET.SubElement(tbl, _qn("w:tblPr"))
    ET.SubElement(tblPr, _qn("w:tblStyle"), {_qn("w:val"): "TableGrid"})
    ET.SubElement(tblPr, _qn("w:bidiVisual"))
    total_w = 8530  # dxa; matches template's main tables
    ET.SubElement(tblPr, _qn("w:tblW"), {_qn("w:w"): str(total_w), _qn("w:type"): "dxa"})
    ET.SubElement(tblPr, _qn("w:jc"), {_qn("w:val"): "center"})
    cell_mar = ET.SubElement(tblPr, _qn("w:tblCellMar"))
    ET.SubElement(cell_mar, _qn("w:left"), {_qn("w:w"): "0", _qn("w:type"): "dxa"})
    ET.SubElement(cell_mar, _qn("w:right"), {_qn("w:w"): "0", _qn("w:type"): "dxa"})
    ET.SubElement(cell_mar, _qn("w:top"), {_qn("w:w"): "0", _qn("w:type"): "dxa"})
    ET.SubElement(cell_mar, _qn("w:bottom"), {_qn("w:w"): "0", _qn("w:type"): "dxa"})
    borders = ET.SubElement(tblPr, _qn("w:tblBorders"))
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        ET.SubElement(
            borders,
            _qn(f"w:{side}"),
            {
                _qn("w:val"): "single",
                _qn("w:sz"): "6",
                _qn("w:space"): "0",
                _qn("w:color"): "C9C9C9",
                _qn("w:themeColor"): "accent3",
                _qn("w:themeTint"): "99",
            },
        )
    ET.SubElement(
        tblPr,
        _qn("w:tblLook"),
        {
            _qn("w:val"): "04A0",
            _qn("w:firstRow"): "1",
            _qn("w:lastRow"): "0",
            _qn("w:firstColumn"): "1",
            _qn("w:lastColumn"): "0",
            _qn("w:noHBand"): "0",
            _qn("w:noVBand"): "1",
        },
    )

    ncols = len(headers)
    if col_weights is None:
        col_weights = [1] * ncols
    if len(col_weights) != ncols:
        col_weights = [1] * ncols
    weight_sum = sum(col_weights) or ncols
    widths = [max(300, int(total_w * w / weight_sum)) for w in col_weights]
    # Fix rounding drift
    drift = total_w - sum(widths)
    widths[-1] += drift

    def _cell_p(text: str, *, bold: bool) -> ET.Element:
        # Table rows are prone to spilling across pages in Word/LibreOffice when
        # paragraph spacing is large. Use compact spacing inside table cells.
        p = ET.Element(_qn("w:p"))
        pPr = ET.SubElement(p, _qn("w:pPr"))
        ET.SubElement(
            pPr,
            _qn("w:spacing"),
            {
                _qn("w:before"): "0",
                _qn("w:after"): "0",
                _qn("w:line"): "240",
                _qn("w:lineRule"): "auto",
            },
        )
        ET.SubElement(pPr, _qn("w:jc"), {_qn("w:val"): "center"})
        _add_rtl_props(pPr, bold=False, italic=False)

        r = ET.SubElement(p, _qn("w:r"))
        _add_rtl_props(r, bold=bold, italic=False, color=None)
        _set_run_text(r, text)
        return p

    def tc(text: str, *, header: bool = False) -> ET.Element:
        cell = ET.Element(_qn("w:tc"))
        tcPr = ET.SubElement(cell, _qn("w:tcPr"))
        # Width will be set by caller after cell creation; placeholder here
        ET.SubElement(tcPr, _qn("w:tcW"), {_qn("w:w"): "0", _qn("w:type"): "dxa"})
        ET.SubElement(tcPr, _qn("w:vAlign"), {_qn("w:val"): "center"})
        # Keep cell text readable and right-to-left
        cell.append(_cell_p(text, bold=header))
        return cell

    # Column grid
    tblGrid = ET.SubElement(tbl, _qn("w:tblGrid"))
    for w in widths:
        ET.SubElement(tblGrid, _qn("w:gridCol"), {_qn("w:w"): str(w)})

    # Header row
    tr = ET.SubElement(tbl, _qn("w:tr"))
    trPr = ET.SubElement(tr, _qn("w:trPr"))
    # Repeat header on each page when the table spans multiple pages.
    ET.SubElement(trPr, _qn("w:tblHeader"))
    # Prevent the header row itself from splitting across pages.
    ET.SubElement(trPr, _qn("w:cantSplit"))
    for col_w, h in zip(widths, headers, strict=False):
        cell = tc(h, header=True)
        tcW = cell.find("./w:tcPr/w:tcW", NS)
        if tcW is not None:
            tcW.set(_qn("w:w"), str(col_w))
        tr.append(cell)

    # Data rows
    for row in rows:
        tr = ET.SubElement(tbl, _qn("w:tr"))
        trPr = ET.SubElement(tr, _qn("w:trPr"))
        # Avoid splitting a single row between two pages (improves readability in Word/LibreOffice).
        ET.SubElement(trPr, _qn("w:cantSplit"))
        for col_w, c in zip(widths, row, strict=False):
            cell = tc(c)
            tcW = cell.find("./w:tcPr/w:tcW", NS)
            if tcW is not None:
                tcW.set(_qn("w:w"), str(col_w))
            tr.append(cell)

    return tbl


def _simple_box_diagram(
    path: Path,
    *,
    title: str,
    boxes: list[tuple[str, tuple[int, int, int, int]]],
    arrows: list[tuple[tuple[int, int], tuple[int, int]]],
) -> None:
    img = Image.new("RGB", (1600, 900), "white")
    draw = ImageDraw.Draw(img)
    font = _load_font(30)
    small = _load_font(24)

    draw.rounded_rectangle((30, 30, 1570, 870), radius=18, outline=(40, 40, 40), width=3)
    draw.text((60, 50), title, fill=(20, 20, 20), font=font)

    for label, (x1, y1, x2, y2) in boxes:
        draw.rounded_rectangle((x1, y1, x2, y2), radius=16, outline=(30, 30, 30), width=3, fill=(245, 245, 245))
        bbox = draw.textbbox((0, 0), label, font=small)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        draw.text(((x1 + x2 - tw) / 2, (y1 + y2 - th) / 2), label, fill=(15, 15, 15), font=small)

    for (x1, y1), (x2, y2) in arrows:
        draw.line((x1, y1, x2, y2), fill=(60, 60, 60), width=4)
        dx, dy = x2 - x1, y2 - y1
        mag = max((dx * dx + dy * dy) ** 0.5, 1.0)
        ux, uy = dx / mag, dy / mag
        px, py = -uy, ux
        hx, hy = x2 - ux * 18, y2 - uy * 18
        p1 = (x2, y2)
        p2 = (hx + px * 10, hy + py * 10)
        p3 = (hx - px * 10, hy - py * 10)
        draw.polygon([p1, p2, p3], fill=(60, 60, 60))

    img.save(path, format="PNG")


def ensure_default_diagrams(diagrams_dir: Path) -> dict[str, Path]:
    """
    Ensures a set of simple placeholder diagrams exist under diagrams/.
    You can later replace these PNGs with exported Visual Paradigm diagrams (same filenames),
    and the generator will embed them into the .docx.
    """
    _ensure_dir(diagrams_dir)

    fig_paths: dict[str, Path] = {
        "2-1": diagrams_dir / "fig-2-1-context.png",
        "2-2": diagrams_dir / "fig-2-2-container.png",
        "2-3": diagrams_dir / "fig-2-3-component.png",
        "4-1": diagrams_dir / "fig-4-1-usecase.png",
        "4-2": diagrams_dir / "fig-4-2-uc01-cache-hit.png",
        "4-3": diagrams_dir / "fig-4-3-uc01-cache-miss.png",
        "4-4": diagrams_dir / "fig-4-4-activity-uc01.png",
        "4-5": diagrams_dir / "fig-4-5-uc02-start-pay.png",
        "4-6": diagrams_dir / "fig-4-6-uc02-callback-verify.png",
        "4-7": diagrams_dir / "fig-4-7-uc02-issue-notify.png",
        "4-8": diagrams_dir / "fig-4-8-activity-uc02.png",
        "4-9": diagrams_dir / "fig-4-9-state-booking.png",
        "5-1": diagrams_dir / "fig-5-1-class-analytical.png",
        "5-2": diagrams_dir / "fig-5-2-class-design.png",
        "5-3": diagrams_dir / "fig-5-3-crc-common.png",
        "7-1": diagrams_dir / "fig-7-1-deploy.png",
        "9-1": diagrams_dir / "fig-9-1-erd.png",
    }

    if not fig_paths["2-1"].exists():
        _simple_box_diagram(
            fig_paths["2-1"],
            title="Context (MarcoPolo)",
            boxes=[
                ("User", (90, 260, 340, 370)),
                ("Web/Mobile", (420, 260, 720, 370)),
                ("API", (800, 260, 1060, 370)),
                ("Providers", (1180, 160, 1490, 270)),
                ("Payment", (1180, 310, 1490, 420)),
                ("Notify", (1180, 460, 1490, 570)),
                ("Support", (1180, 610, 1490, 720)),
            ],
            arrows=[
                ((340, 315), (420, 315)),
                ((720, 315), (800, 315)),
                ((1060, 315), (1180, 215)),
                ((1060, 315), (1180, 365)),
                ((1060, 315), (1180, 515)),
                ((1060, 315), (1180, 665)),
            ],
        )

    if not fig_paths["2-2"].exists():
        _simple_box_diagram(
            fig_paths["2-2"],
            title="Containers",
            boxes=[
                ("Web UI", (120, 180, 420, 280)),
                ("Mobile UI", (120, 320, 420, 420)),
                ("API (Backend)", (520, 240, 930, 410)),
                ("DB", (1030, 180, 1460, 280)),
                ("Cache", (1030, 320, 1460, 420)),
                ("Queue", (1030, 460, 1460, 560)),
                ("External Services", (520, 460, 930, 620)),
            ],
            arrows=[
                ((420, 230), (520, 300)),
                ((420, 370), (520, 330)),
                ((930, 290), (1030, 230)),
                ((930, 330), (1030, 370)),
                ((930, 370), (1030, 510)),
                ((930, 470), (930, 410)),
            ],
        )

    if not fig_paths["2-3"].exists():
        _simple_box_diagram(
            fig_paths["2-3"],
            title="Backend Components",
            boxes=[
                ("API Layer", (120, 200, 520, 310)),
                ("Domain Services", (120, 360, 520, 470)),
                ("Integrations", (640, 200, 1120, 310)),
                ("Data Access", (640, 360, 1120, 470)),
                ("Background Jobs", (640, 520, 1120, 630)),
            ],
            arrows=[
                ((520, 255), (640, 255)),
                ((520, 415), (640, 415)),
                ((520, 415), (640, 575)),
                ((320, 310), (320, 360)),
            ],
        )

    for fig_id in [
        "4-1",
        "4-2",
        "4-3",
        "4-4",
        "4-5",
        "4-6",
        "4-7",
        "4-8",
        "4-9",
        "5-1",
        "5-2",
        "5-3",
        "7-1",
        "9-1",
    ]:
        p = fig_paths[fig_id]
        if p.exists():
            continue
        _simple_box_diagram(
            p,
            title=f"Figure {fig_id}",
            boxes=[("Diagram (replace later)", (350, 350, 1250, 520))],
            arrows=[],
        )

    return fig_paths


def _ensure_png_content_type(file_bytes: dict[str, bytes]) -> None:
    ct_xml = file_bytes.get("[Content_Types].xml")
    if ct_xml is None:
        return
    root = ET.fromstring(ct_xml)
    existing = {d.attrib.get("Extension") for d in root.findall(_qns(CT_NS, "Default"))}
    if "png" in existing:
        return
    ET.SubElement(root, _qns(CT_NS, "Default"), {"Extension": "png", "ContentType": "image/png"})
    xml = ET.tostring(root, encoding="utf-8", xml_declaration=True).decode("utf-8", errors="replace")
    # LibreOffice is more compatible with default-namespace (no prefix) in [Content_Types].xml
    xml = re.sub(r"<(/?)ct:", r"<\1", xml)
    xml = xml.replace(f'xmlns:ct=\"{CT_NS}\"', f'xmlns=\"{CT_NS}\"')
    file_bytes["[Content_Types].xml"] = xml.encode("utf-8")


def _ensure_document_rels(file_bytes: dict[str, bytes]) -> ET.Element:
    rels_path = "word/_rels/document.xml.rels"
    rels_xml = file_bytes.get(rels_path)
    if rels_xml is None:
        root = ET.Element(_qns(REL_NS, "Relationships"))
        xml = ET.tostring(root, encoding="utf-8", xml_declaration=True).decode("utf-8", errors="replace")
        xml = re.sub(r"<(/?)rel:", r"<\1", xml)
        xml = xml.replace(f'xmlns:rel=\"{REL_NS}\"', f'xmlns=\"{REL_NS}\"')
        file_bytes[rels_path] = xml.encode("utf-8")
        return root
    return ET.fromstring(rels_xml)


def _next_rid(rels_root: ET.Element) -> str:
    max_n = 0
    for rel in rels_root.findall(_qns(REL_NS, "Relationship")):
        rid = rel.attrib.get("Id", "")
        if rid.startswith("rId"):
            try:
                max_n = max(max_n, int(rid[3:]))
            except Exception:
                continue
    return f"rId{max_n + 1}"


def _add_image_relationship(rels_root: ET.Element, target: str) -> str:
    rid = _next_rid(rels_root)
    ET.SubElement(
        rels_root,
        _qns(REL_NS, "Relationship"),
        {
            "Id": rid,
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            "Target": target,
        },
    )
    return rid


def _px_to_emu(px: int, dpi: int = DEFAULT_IMAGE_DPI) -> int:
    return int(px / dpi * EMU_PER_INCH)


def _make_image_paragraph(rid: str, *, cx: int, cy: int, docpr_id: int, name: str) -> ET.Element:
    # Centered paragraph with an inline picture.
    p = ET.Element(_qn("w:p"))
    pPr = ET.SubElement(p, _qn("w:pPr"))
    ET.SubElement(pPr, _qn("w:jc"), {_qn("w:val"): "center"})
    ET.SubElement(
        pPr,
        _qn("w:spacing"),
        {
            _qn("w:before"): "0",
            _qn("w:beforeAutospacing"): "1",
            _qn("w:after"): "0",
            _qn("w:afterAutospacing"): "1",
            _qn("w:line"): "360",
            _qn("w:lineRule"): "auto",
        },
    )
    ET.SubElement(pPr, _qn("w:keepNext"))
    _add_rtl_props(pPr, bold=False, italic=False)

    r = ET.SubElement(p, _qn("w:r"))
    _add_rtl_props(r, bold=False, italic=False)

    drawing = ET.SubElement(r, _qn("w:drawing"))
    # Use wp:anchor (as LibreOffice itself produces) instead of wp:inline for better compatibility.
    anchor = ET.SubElement(
        drawing,
        _qns(WP_NS, "anchor"),
        {
            "behindDoc": "0",
            "distT": "0",
            "distB": "0",
            "distL": "0",
            "distR": "0",
            "simplePos": "0",
            "locked": "0",
            "layoutInCell": "0",
            "allowOverlap": "1",
            "relativeHeight": "2",
        },
    )
    ET.SubElement(anchor, _qns(WP_NS, "simplePos"), {"x": "0", "y": "0"})
    posH = ET.SubElement(anchor, _qns(WP_NS, "positionH"), {"relativeFrom": "column"})
    ET.SubElement(posH, _qns(WP_NS, "align")).text = "center"
    posV = ET.SubElement(anchor, _qns(WP_NS, "positionV"), {"relativeFrom": "paragraph"})
    ET.SubElement(posV, _qns(WP_NS, "posOffset")).text = "0"

    ET.SubElement(anchor, _qns(WP_NS, "extent"), {"cx": str(cx), "cy": str(cy)})
    ET.SubElement(anchor, _qns(WP_NS, "effectExtent"), {"l": "0", "t": "0", "r": "0", "b": "0"})
    ET.SubElement(anchor, _qns(WP_NS, "wrapSquare"), {"wrapText": "largest"})
    ET.SubElement(anchor, _qns(WP_NS, "docPr"), {"id": str(docpr_id), "name": name, "descr": ""})
    cNv = ET.SubElement(anchor, _qns(WP_NS, "cNvGraphicFramePr"))
    ET.SubElement(cNv, _qns(A_NS, "graphicFrameLocks"), {"noChangeAspect": "1"})

    graphic = ET.SubElement(anchor, _qns(A_NS, "graphic"))
    graphicData = ET.SubElement(graphic, _qns(A_NS, "graphicData"), {"uri": "http://schemas.openxmlformats.org/drawingml/2006/picture"})

    pic = ET.SubElement(graphicData, _qns(PIC_NS, "pic"))
    nvPicPr = ET.SubElement(pic, _qns(PIC_NS, "nvPicPr"))
    ET.SubElement(nvPicPr, _qns(PIC_NS, "cNvPr"), {"id": str(docpr_id), "name": name, "descr": ""})
    cNvPicPr = ET.SubElement(nvPicPr, _qns(PIC_NS, "cNvPicPr"))
    ET.SubElement(cNvPicPr, _qns(A_NS, "picLocks"), {"noChangeAspect": "1", "noChangeArrowheads": "1"})

    blipFill = ET.SubElement(pic, _qns(PIC_NS, "blipFill"))
    blip = ET.SubElement(blipFill, _qns(A_NS, "blip"), {_qns(R_NS, "embed"): rid, "cstate": "print"})
    stretch = ET.SubElement(blipFill, _qns(A_NS, "stretch"))
    ET.SubElement(stretch, _qns(A_NS, "fillRect"))

    spPr = ET.SubElement(pic, _qns(PIC_NS, "spPr"), {"bwMode": "auto"})
    xfrm = ET.SubElement(spPr, _qns(A_NS, "xfrm"))
    ET.SubElement(xfrm, _qns(A_NS, "off"), {"x": "0", "y": "0"})
    ET.SubElement(xfrm, _qns(A_NS, "ext"), {"cx": str(cx), "cy": str(cy)})
    prstGeom = ET.SubElement(spPr, _qns(A_NS, "prstGeom"), {"prst": "rect"})
    ET.SubElement(prstGeom, _qns(A_NS, "avLst"))

    return p


def _max_docpr_id(doc_root: ET.Element) -> int:
    max_id = 0
    for el in doc_root.iter():
        if el.tag == _qns(WP_NS, "docPr"):
            try:
                max_id = max(max_id, int(el.attrib.get("id", "0")))
            except Exception:
                continue
    return max_id


def embed_figures(
    root: ET.Element,
    file_bytes: dict[str, bytes],
    *,
    diagrams_dir: Path,
    autogen: bool = True,
    embed_images: bool = True,
) -> None:
    """
    Replaces marker paragraphs like [FIG:2-1] with embedded images from diagrams/.
    If embed_images is False, marker paragraphs are removed (captions remain).
    """
    if autogen:
        ensure_default_diagrams(diagrams_dir)

    # Always map figure IDs to the expected filenames (so replacing later is stable).
    expected = ensure_default_diagrams(diagrams_dir) if autogen else {
        "2-1": diagrams_dir / "fig-2-1-context.png",
        "2-2": diagrams_dir / "fig-2-2-container.png",
        "2-3": diagrams_dir / "fig-2-3-component.png",
        "4-1": diagrams_dir / "fig-4-1-usecase.png",
        "4-2": diagrams_dir / "fig-4-2-uc01-cache-hit.png",
        "4-3": diagrams_dir / "fig-4-3-uc01-cache-miss.png",
        "4-4": diagrams_dir / "fig-4-4-activity-uc01.png",
        "4-5": diagrams_dir / "fig-4-5-uc02-start-pay.png",
        "4-6": diagrams_dir / "fig-4-6-uc02-callback-verify.png",
        "4-7": diagrams_dir / "fig-4-7-uc02-issue-notify.png",
        "4-8": diagrams_dir / "fig-4-8-activity-uc02.png",
        "4-9": diagrams_dir / "fig-4-9-state-booking.png",
        "5-1": diagrams_dir / "fig-5-1-class-analytical.png",
        "5-2": diagrams_dir / "fig-5-2-class-design.png",
        "5-3": diagrams_dir / "fig-5-3-crc-common.png",
        "7-1": diagrams_dir / "fig-7-1-deploy.png",
        "9-1": diagrams_dir / "fig-9-1-erd.png",
    }

    def _pick_diagram_path(base_path: Path | None) -> Path | None:
        if base_path is None:
            return None
        vp_path = base_path.with_name(f"{base_path.stem}-vp{base_path.suffix}")
        if vp_path.exists():
            return vp_path
        if base_path.exists():
            return base_path
        # If the base diagram is missing but the VP variant exists (e.g., autogen disabled),
        # prefer the VP export when available.
        return vp_path if vp_path.exists() else None

    rels_root: ET.Element | None = None
    docpr_id = 1000

    body = root.find("w:body", NS)
    if body is None:
        return

    # Scan paragraphs and replace markers.
    for p in list(body.findall("w:p", NS)):
        txt = _p_text(p)
        if not (txt.startswith(FIGURE_MARKER_PREFIX) and txt.endswith("]")):
            continue
        fig_id = txt[len(FIGURE_MARKER_PREFIX) : -1]
        img_path = _pick_diagram_path(expected.get(fig_id))
        if not embed_images:
            body.remove(p)
            continue
        if not img_path or not img_path.exists():
            body.remove(p)
            continue

        if rels_root is None:
            _ensure_png_content_type(file_bytes)
            rels_root = _ensure_document_rels(file_bytes)
            docpr_id = max(_max_docpr_id(root) + 1, 1000)

        img_bytes = img_path.read_bytes()
        media_name = f"image-fig-{fig_id}.png"
        media_target = f"media/{media_name}"
        file_bytes[f"word/{media_target}"] = img_bytes

        rid = _add_image_relationship(rels_root, media_target)

        with Image.open(img_path) as im:
            cx = _px_to_emu(im.size[0])
            cy = _px_to_emu(im.size[1])
            # Fit to page width (roughly): cap width to ~6.5 inches.
            max_cx = int(6.5 * EMU_PER_INCH)
            if cx > max_cx:
                scale = max_cx / max(cx, 1)
                cx = int(cx * scale)
                cy = int(cy * scale)

        img_p = _make_image_paragraph(rid, cx=cx, cy=cy, docpr_id=docpr_id, name=media_name)
        docpr_id += 1

        idx = list(body).index(p)
        body.remove(p)
        body.insert(idx, img_p)

    if rels_root is not None:
        xml = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True).decode("utf-8", errors="replace")
        xml = re.sub(r"<(/?)rel:", r"<\1", xml)
        xml = xml.replace(f'xmlns:rel=\"{REL_NS}\"', f'xmlns=\"{REL_NS}\"')
        file_bytes["word/_rels/document.xml.rels"] = xml.encode("utf-8")


def replace_first_paragraph_text(root: ET.Element, old: str, new: str) -> None:
    for p in root.findall(".//w:body/w:p", NS):
        if _p_text(p) == old:
            # Remove all runs and add a single run
            for child in list(p):
                if child.tag == _qn("w:r"):
                    p.remove(child)
            r = ET.SubElement(p, _qn("w:r"))
            _add_rtl_props(r)
            _set_run_text(r, new)
            return


def fill_history_table(root: ET.Element) -> None:
    # Find the first table after the paragraph "تاريخچه بازبيني"
    body = root.find("w:body", NS)
    if body is None:
        return
    children = list(body)
    idx = None
    for i, el in enumerate(children):
        if el.tag == _qn("w:p") and _p_text(el) == "تاريخچه بازبيني":
            idx = i
            break
    if idx is None:
        return
    tbl = None
    for el in children[idx + 1 :]:
        if el.tag == _qn("w:tbl"):
            tbl = el
            break
        if el.tag == _qn("w:p") and _p_text(el):
            # hit another paragraph; stop searching
            continue
    if tbl is None:
        return

    # Try to write into the first data row (second row). If not present, do nothing.
    trs = tbl.findall("./w:tr", NS)
    if len(trs) < 2:
        return
    data_tr = trs[1]
    tcs = data_tr.findall("./w:tc", NS)
    if len(tcs) < 4:
        return

    today = _dt.date.today()
    jy, jm, jd = _gregorian_to_jalali(today.year, today.month, today.day)
    jalali_date = _to_persian_digits(f"{jy:04d}/{jm:02d}/{jd:02d}")
    members = get_group_members()
    prepared_by = members.replace("،", " / ")
    values = [jalali_date, DOC_VERSION, "نسخه نهایی", prepared_by]
    for tc_el, val in zip(tcs, values, strict=False):
        # Replace first paragraph text inside cell
        p = tc_el.find("./w:p", NS)
        if p is None:
            continue
        for child in list(p):
            if child.tag == _qn("w:r"):
                p.remove(child)
        r = ET.SubElement(p, _qn("w:r"))
        _add_rtl_props(r)
        _set_run_text(r, val)


def ensure_toc_field(root: ET.Element, file_bytes: dict[str, bytes]) -> None:
    """
    Ensure the template's TOC updates on open.

    The provided `SAD-Template.docx` already contains a proper TOC field (complex field).
    We keep its structure intact (so the visible layout stays like the template) and only
    enable automatic field updates in document settings.
    """
    # Ensure fields update on open (Word/LibreOffice will regenerate the TOC result).
    settings = file_bytes.get("word/settings.xml")
    if settings is None:
        return
    try:
        sroot = ET.fromstring(settings)
    except Exception:
        return
    upd = sroot.find("w:updateFields", NS)
    if upd is None:
        upd = ET.SubElement(sroot, _qn("w:updateFields"))
    upd.attrib[_qn("w:val")] = "true"
    file_bytes["word/settings.xml"] = ET.tostring(sroot, encoding="utf-8", xml_declaration=True)


def _p_style_val(p: ET.Element) -> str | None:
    pPr = p.find("w:pPr", NS)
    if pPr is None:
        return None
    st = pPr.find("w:pStyle", NS)
    if st is None:
        return None
    return st.attrib.get(_qns(W_NS, "val"))


def _max_bookmark_id(root: ET.Element) -> int:
    max_id = 0
    for el in root.iter():
        if el.tag == _qn("w:bookmarkStart"):
            try:
                max_id = max(max_id, int(el.attrib.get(_qns(W_NS, "id"), "0")))
            except Exception:
                continue
    return max_id


def ensure_heading_bookmarks(root: ET.Element, *, content_start_idx: int) -> list[tuple[int, str, str]]:
    """
    Ensure each Heading1/2/3 paragraph in the generated content has a unique bookmark.
    Returns a list of (level, title, bookmark_name) in document order.
    """
    body = root.find("w:body", NS)
    if body is None:
        return []
    children = list(body)
    next_bm_id = max(_max_bookmark_id(root) + 1, 1000)
    toc_items: list[tuple[int, str, str]] = []
    seq = 1

    for el in children[content_start_idx:]:
        if el.tag != _qn("w:p"):
            continue
        st = _p_style_val(el)
        if st not in ("Heading1", "Heading2", "Heading3"):
            continue
        title = _p_text(el)
        if not title:
            continue
        level = 1 if st == "Heading1" else 2 if st == "Heading2" else 3

        # If there's already a bookmark on this paragraph, reuse the first one.
        existing = el.find("w:bookmarkStart", NS)
        if existing is not None:
            name = existing.attrib.get(_qns(W_NS, "name"))
            if name:
                toc_items.append((level, title, name))
                continue

        name = f"_TocCustom{seq}"
        seq += 1

        bm_start = ET.Element(_qn("w:bookmarkStart"), {_qn("w:id"): str(next_bm_id), _qn("w:name"): name})
        bm_end = ET.Element(_qn("w:bookmarkEnd"), {_qn("w:id"): str(next_bm_id)})
        next_bm_id += 1

        # Place bookmark around the whole paragraph content.
        el.insert(0, bm_start)
        el.append(bm_end)

        toc_items.append((level, title, name))

    return toc_items


def rebuild_toc_like_template(root: ET.Element, *, content_start_idx: int) -> None:
    """
    Replace the visible TOC entries area with paragraphs styled like the template (TOC1/2/3),
    but generated from the current headings. Page numbers are inserted as PAGEREF fields so
    they update correctly (including 2-digit pages) when fields are refreshed in Word/LibreOffice.
    """
    body = root.find("w:body", NS)
    if body is None:
        return
    children = list(body)

    toc_heading_idx = None
    for i, el in enumerate(children):
        if el.tag == _qn("w:p") and _p_text(el) == "فهرست مطالب":
            toc_heading_idx = i
            break
    if toc_heading_idx is None:
        return

    toc_end_idx = None
    for i in range(toc_heading_idx + 1, len(children)):
        el = children[i]
        if el.tag == _qn("w:p") and _p_text(el) == "سند معماری نرم‌افزار":
            toc_end_idx = i
            break
    if toc_end_idx is None:
        return

    toc_items = ensure_heading_bookmarks(root, content_start_idx=content_start_idx)
    if not toc_items:
        return

    # Remove existing visible TOC paragraphs between heading and end marker.
    for el in list(children[toc_heading_idx + 1 : toc_end_idx]):
        if el.tag != _qn("w:p"):
            continue
        st = _p_style_val(el) or ""
        if st.startswith("TOC"):
            body.remove(el)

    h1 = h2 = h3 = 0
    insert_at = toc_heading_idx + 1
    for level, title, bm_name in toc_items:
        if level == 1:
            h1 += 1
            h2 = 0
            h3 = 0
            prefix = f"{h1}-"
            style = "TOC1"
        elif level == 2:
            h2 += 1
            h3 = 0
            prefix = f"{h1}-{h2}-"
            style = "TOC2"
        else:
            h3 += 1
            prefix = f"{h1}-{h2}-{h3}-"
            style = "TOC3"

        p = ET.Element(_qn("w:p"))
        pPr = ET.SubElement(p, _qn("w:pPr"))
        ET.SubElement(pPr, _qn("w:pStyle"), {_qn("w:val"): style})

        # Hyperlinked entry text.
        link = ET.SubElement(p, _qn("w:hyperlink"), {_qn("w:anchor"): bm_name, _qn("w:history"): "1"})
        r_text = ET.SubElement(link, _qn("w:r"))
        _add_rtl_props(r_text)
        _set_run_text(r_text, sanitize_text(prefix + title))

        # Leader dots and right-aligned page number area is handled by the TOC style tab stops.
        r_tab = ET.SubElement(p, _qn("w:r"))
        _add_rtl_props(r_tab)
        ET.SubElement(r_tab, _qn("w:tab"))

        # Page number as a field (updates to 10, 11, ... correctly).
        fld = ET.SubElement(p, _qn("w:fldSimple"), {_qn("w:instr"): f"PAGEREF {bm_name} \\h"})
        r_page = ET.SubElement(fld, _qn("w:r"))
        _add_rtl_props(r_page)
        _set_run_text(r_page, "")

        body.insert(insert_at, p)
        insert_at += 1


def build_sad_content(*, fig_caption_red: bool = True) -> list[ET.Element]:
    el: list[ET.Element] = []

    def fig_caption(text: str) -> ET.Element:
        return make_p(
            text,
            italic=True,
            jc="center",
            color="FF0000" if fig_caption_red else None,
            spacing_before=0,
            spacing_after=100,
        )

    # 1) کلیات سند
    el.append(make_p("کليات سند", style="Heading1"))
    el.append(make_p("هدف", style="Heading2"))
    el.append(
        make_p(
            "سامانه مارکوپولو یک پلتفرم برخط برای جست‌وجو، خرید و مدیریت خدمات سفر و گردشگری است. "
            "این سند معماری سامانه را به‌صورت رسمی توصیف می‌کند و معیارهای لازم برای ارزیابی تصمیم‌های معماری را ارائه می‌دهد. "
            "در بخش‌های بعدی، مرزبندی اجزا، رابط‌ها و نحوه تعامل با سامانه‌های بیرونی توضیح داده می‌شود و تصمیم‌های کلیدی به‌همراه پیامدهای آن‌ها ثبت می‌گردد تا در توسعه، آزمون و نگهداشت به کار بیاید. "
            "معیارهای کلیدی ارزیابی معماری نیز مشخص شده است تا سنجش و پیگیری تصمیم‌ها امکان‌پذیر باشد. "
            "این سند جایگزین طراحی جزئیِ پیاده‌سازی یا مستندات رابط‌های بیرونیِ هر تأمین‌کننده نیست؛ تمرکز آن روی تصمیم‌های کلان، رفتار سناریوهای کلیدی و سنجه‌هایی است که معماری را قابل داوری می‌کند.",
            jc="both",
        )
    )

    el.append(make_p("مخاطبان و نحوه استفاده", style="Heading2"))
    el.append(
        make_p(
            "این سند برای هماهنگ‌کردن نگاه تیم‌ها به معماری سامانه نوشته شده است. "
            "هدف آن این است که تصمیم‌های کلیدی، مرزبندی مسئولیت‌ها و معیارهای کیفیتی به‌صورت شفاف ثبت شوند تا در زمان توسعه، "
            "آزمون و بهره‌برداری به اختلاف‌نظرهای سلیقه‌ای تبدیل نشوند. در عمل، هر بخش از سند یک کار مشخص را پشتیبانی می‌کند: "
            "نمودارهای کلان برای درک ساختار، سناریوها برای بررسی رفتار و خطاها، و بخش‌های کیفیت برای داوری بر اساس سنجه.",
            jc="both",
        )
    )
    el.append(
        make_tbl(
            ["مخاطب", "چیزی که از سند انتظار دارد", "بخش‌های کلیدی"],
            [
                ["تیم فنی", "تصمیم‌های معماری، مرزبندی دامنه‌ها، قراردادهای بیرونی", "نمایش معماری، دید منطقی، دید توسعه"],
                ["تیم آزمون", "سناریوهای قابل آزمون، جریان‌های خطا و معیارهای پذیرش", "دید سناریوها، کیفیت، کارایی"],
                ["تیم عملیات", "الگوی استقرار، پایش، بازیابی و مدیریت رخداد", "دید فیزیکی، بهره‌برداری و عملیات"],
                ["کسب‌وکار", "حدود قابلیت‌ها، ریسک‌ها و پیامدهای تصمیم‌ها", "کلیات سند، اهداف و محدودیت‌ها"],
            ],
            col_weights=[1, 2, 2],
        )
    )

    el.append(make_p("فرض‌ها و پیش‌فرض‌ها", style="Heading2"))
    el.append(
        make_p(
            "برای اینکه تصمیم‌های معماری قابل داوری باشد، چند فرض پایه در نظر گرفته شده است. "
            "اگر هر کدام از این فرض‌ها تغییر کند، لازم است بخش‌های مربوط بازنگری شود تا تناقض ایجاد نشود.",
            jc="both",
        )
    )
    el.append(
        make_tbl(
            ["فرض", "اثر روی معماری", "نشانه پایش", "اگر فرض نقض شود"],
            [
                [
                    "پاسخ تأمین‌کنندگان ناهمگون و گاهی ناپایدار است.",
                    "کنترل مهلت زمانی، کاهش سطح خدمت و تلاش‌مجدد کنترل‌شده ضروری است.",
                    "افزایش نرخ خطا یا زمان پاسخ به تفکیک تأمین‌کننده",
                    "ترکیب نتایج و تجربه کاربر باید بازطراحی شود (مثلاً نمایش تدریجی یا محدودسازی تأمین‌کننده).",
                ],
                [
                    "بازگشت بانک قابل تکرار است و ممکن است چند بار ارسال شود.",
                    "کلید یکتایی عملیات و راستی‌آزمایی پرداخت الزامی است.",
                    "افزایش شمار بازگشت‌های تکراری یا اختلاف وضعیت پرداخت",
                    "ریسک صدور تکراری/ثبت تکراری بالا می‌رود و کنترل‌های مالی باید سخت‌گیرانه‌تر شود.",
                ],
                [
                    "APIها به‌صورت عمومی در معرض اینترنت هستند.",
                    "محدودیت نرخ، احراز هویت و ثبت ممیزی باید از ابتدا اعمال شود.",
                    "افزایش درخواست‌های غیرعادی یا خطاهای احراز هویت",
                    "در صورت نبود این کنترل‌ها، احتمال سوءاستفاده و اختلال بالا می‌رود.",
                ],
                [
                    "در فاز نخست، استقرار به‌صورت تک‌استقرار ماژولار انجام می‌شود.",
                    "سادگی عملیات و سرعت توسعه اولویت دارد، با حفظ مرزبندی دامنه‌ها.",
                    "افزایش نیاز به استقرار مستقل یا رشد تیم/بار",
                    "اگر نیاز به استقرار مستقل زودتر ایجاد شود، مرزبندی دامنه‌ها باید عملی‌تر و رابط‌ها دقیق‌تر شوند.",
                ],
            ],
            col_weights=[2, 2, 1, 2],
        )
    )

    el.append(make_p("محدوده", style="Heading2"))
    el.append(
        make_p(
            "از منظر کارکردی، سامانه سه جریان اصلی جست‌وجو، خرید و صدور، و استرداد را پوشش می‌دهد و در کنار آن "
            "قابلیت‌های مکمل مانند تور و اقامت، کیف‌پول، پشتیبانی و بازخورد را ارائه می‌کند. "
            "در جست‌وجو، نتایج چند تأمین‌کننده تجمیع و یکپارچه می‌شود و امکان فیلتر، مرتب‌سازی و صفحه‌بندی در اختیار کاربر قرار می‌گیرد. "
            "در خرید، پس از ثبت اطلاعات مسافر و بازتأیید ظرفیت و قیمت، پرداخت آغاز می‌شود و در صورت موفقیت، صدور بلیت و ارسال اعلان انجام می‌گردد. "
            "در استرداد نیز قوانین و جریمه‌ها اعمال می‌شود و بازپرداخت از مسیر مناسب (درگاه یا کیف‌پول) مدیریت می‌شود.",
            jc="both",
        )
    )
    el.append(
        make_p(
            "خارج از محدوده این سند، پیاده‌سازی سامانه‌های تأمین‌کنندگان و درگاه بانکی است. "
            "همچنین اگر برای پشتیبانی از یک سامانه بیرونی (مثل مرکز تماس یا سامانه تیکتینگ) استفاده شود، پیاده‌سازی آن نیز خارج از محدوده است. "
            "این سرویس‌های بیرونی در معماری از منظر قرارداد تبادل داده، مدیریت خطا، مهلت زمانی پاسخ و امنیت ارتباطات تحلیل می‌شوند. "
            "زیرساخت پیامک/ایمیل هم به‌عنوان سرویس بیرونی در نظر گرفته می‌شود.",
            jc="both",
        )
    )

    el.append(make_p("واژه‌نامه و تعاریف", style="Heading2"))
    el.append(
        make_tbl(
            ["اصطلاح", "تعریف"],
            [
                ["تأمین‌کننده", "سامانه بیرونی که گزینه‌های سفر/اقامت/تور را ارائه می‌کند و از طریق API با سامانه تبادل داده دارد."],
                ["درگاه پرداخت", "سامانه بانکی/پرداخت که شروع پرداخت، بازگشت پرداخت و راستی‌آزمایی تراکنش را ارائه می‌کند."],
                ["سفارش", "رکوردی که وضعیت خرید را از ایجاد تا پرداخت و صدور/استرداد نگه می‌دارد."],
                ["تراکنش", "ثبت عملیات مالی شامل پرداخت/بازپرداخت همراه با شناسه‌های پیگیری و وضعیت."],
                ["کلید یکتایی عملیات", "شناسه‌ای یکتا برای جلوگیری از اجرای تکراری یک عملیات در شرایط تلاش‌مجدد یا بازگشت تکراری بانک."],
                ["شناسه ردیابی", "شناسه‌ای برای پیگیری انتها به انتهای درخواست‌ها در لاگ‌ها و ابزارهای پایش (از لحظه ورود درخواست تا دریافت پاسخ از سرویس‌های بیرونی)."],
                ["کاهش سطح خدمت", "نمایش نتیجه ناقص/کنترل‌شده در صورت کندی یا قطع سرویس بیرونی، بدون از کار افتادن کل تجربه کاربر."],
                ["مهلت زمانی", "بازه زمانی حداکثر برای انتظار پاسخ از سرویس بیرونی؛ پس از آن، مسیر خطا/کاهش سطح خدمت فعال می‌شود."],
                ["تلاش‌مجدد کنترل‌شده", "تلاش دوباره برای خطاهای موقت با تعداد محدود و فاصله زمانی مناسب، بدون ایجاد بار اضافی یا تکرار ناخواسته."],
                ["نیازمند بررسی", "وضعیتی که نشان می‌دهد سامانه برای ادامه مسیر، به بررسی انسانی/پشتیبانی یا داده تکمیلی نیاز دارد."],
            ],
            col_weights=[1, 3],
        )
    )

    el.append(make_p("ذی‌نفعان و نیازهای کلیدی"))
    el.append(
        make_tbl(
            ["ذی‌نفع", "نیاز/انتظار"],
            [
                ["مشتری", "جست‌وجوی سریع و شفاف، پرداخت امن، صدور بلیت قابل اتکا، پشتیبانی پاسخگو"],
                ["تیم پشتیبانی", "دسترسی به وضعیت سفارش/پرداخت، ابزار پیگیری، ثبت وقایع و ممیزی"],
                ["تیم فنی", "معماری قابل نگهداری، جداسازی لایه یکپارچه‌سازی، مشاهده‌پذیری مناسب"],
                ["تأمین‌کننده", "ارتباط استاندارد و پایدار، کاهش درخواست‌های تکراری، تسویه‌حساب شفاف"],
                ["کسب‌وکار", "افزایش نرخ تبدیل، کاهش خطا، داده‌های قابل تحلیل، گزارش‌های مدیریتی"],
            ],
        )
    )

    # 2) نمایش معماری
    el.append(make_p("نمایش معماری", style="Heading1"))
    el.append(make_p("سبک معماری (انتخاب‌شده)", style="Heading2"))
    el.append(
        make_p(
            "در سمت سرور، معماری لایه‌ای انتخاب شده است. درخواست‌ها از طریق API وارد می‌شوند و پس از اعمال سیاست‌های عمومی "
            "(مثل اعتبارسنجی، احراز هویت و مجوز، و محدودیت نرخ) به لایه دامنه و خدمات می‌رسند.",
            jc="both",
        )
    )
    el.append(
        make_p(
            "لایه یکپارچه‌سازی مسئول ارتباط پایدار و قابل تغییر با سرویس‌های بیرونی (تأمین‌کنندگان، پرداخت، اعلان، پشتیبانی) است؛ "
            "و لایه داده هم ذخیره‌سازی و دسترسی به پایگاه داده و حافظه نهان را مدیریت می‌کند.",
            jc="both",
        )
    )
    el.append(
        make_p(
            "در فاز نخست، استقرار به‌صورت تک‌استقرار ماژولار انجام می‌شود تا تیم سریع‌تر به نسخه قابل ارائه برسد و عملیات هم ساده‌تر بماند. "
            "با این حال مرزبندی دامنه‌ها از ابتدا رعایت می‌شود تا در صورت رشد تیم یا افزایش بار، گذار تدریجی به ریزخدمت با کمترین بازطراحی ممکن باشد.",
            jc="both",
        )
    )
    el.append(make_fig_marker("2-1"))
    el.append(fig_caption("شکل ۲-۱: نمودار زمینه سامانه."))
    el.append(make_fig_marker("2-2"))
    el.append(fig_caption("شکل ۲-۲: نمودار کانتینرهای سامانه."))
    el.append(make_fig_marker("2-3"))
    el.append(fig_caption("شکل ۲-۳: نمودار اجزای سامانه سمت سرور."))

    # 3) اهداف و محدودیت‌های معماری
    el.append(make_p("اهداف و محدودیت‌های معماری", style="Heading1"))
    el.append(make_p("اهداف کیفیتی", style="Heading2"))
    el.append(
        make_tbl(
            ["ویژگی کیفی", "معیار ارزیابی", "اندازه/محدوده مطلوب", "راهکار معماری"],
            [
                ["کارایی", "صدک ۹۵ زمان پاسخ جست‌وجو", "کمتر از ۳ ثانیه", "حافظه نهان، صفحه‌بندی، زمان‌سنجی برای تأمین‌کنندگان"],
                ["دسترس‌پذیری", "تعداد رخداد قطعی", "کمتر از ۱ بار در هفته", "کاهش سطح خدمت، تلاش مجدد کنترل‌شده، قطع‌کننده مدار سبک"],
                ["مقیاس‌پذیری", "توان عملیاتی API", "مقیاس‌پذیری افقی", "API بدون حالت + حافظه نهان + پایگاه داده مشترک (در فاز نخست)"],
                ["امنیت", "رخداد نفوذ/تقلب", "حداقل", "راستی‌آزمایی پرداخت، امضا/توکن بازگشت بانک، ثبت وقایع ممیزی"],
                ["نگهداشت‌پذیری", "افزودن تأمین‌کننده جدید", "سریع و کم‌ریسک", "مبدل/روش کارخانه + آزمون قرارداد"],
                ["مشاهده‌پذیری", "قابلیت ردیابی", "کامل", "شناسه ردیابی/شناسه همبستگی در خرید/پرداخت"],
            ],
        )
    )

    el.append(make_p("محدودیت‌ها", style="Heading2"))
    el.append(
        make_p(
            "محدودیت اصلی سامانه، ناهمگنی و تغییرپذیری تأمین‌کنندگان است. هر تأمین‌کننده می‌تواند قرارداد، محدودیت نرخ و الگوی خطای متفاوتی داشته باشد و همین موضوع، "
            "طراحی یکپارچه‌سازی را به یکی از نقاط حساس سامانه تبدیل می‌کند.",
            jc="both",
        )
    )
    el.append(
        make_p(
            "در جست‌وجو اگر یک تأمین‌کننده کند یا قطع شود، تجربه کاربر نباید به‌طور کامل خراب شود. به همین دلیل، سامانه نتیجه سایر تأمین‌کنندگان را نمایش می‌دهد و "
            "فقط همان بخش را با پیام مناسب کنار می‌گذارد.",
            jc="both",
        )
    )
    el.append(
        make_p(
            "در پرداخت چون با پول و اعتماد کاربر سروکار داریم، صرف دریافت بازگشت بانک کافی نیست و تراکنش راستی‌آزمایی می‌شود. "
            "بازگشت‌های تکراری هم محتمل است؛ کلید یکتایی عملیات جلوی پردازش دوباره را می‌گیرد.",
            jc="both",
        )
    )
    el.append(
        make_p(
            "در کیف‌پول هم سازگاری موجودی و جلوگیری از دوباره‌خرج‌کردن حیاتی است. در کنار آن، رعایت حریم خصوصی و ثبت رویدادهای قابل ممیزی کمک می‌کند "
            "پاسخ‌گویی و پیگیری عملیاتی واقعاً شدنی باشد.",
            jc="both",
        )
    )

    el.append(make_p("شاخص‌های سنجش‌پذیری", style="Heading2"))
    el.append(
        make_p(
            "برای سنجش‌پذیر شدن اهداف کیفیتی، سنجه‌ها در نقاط مناسب اندازه‌گیری و ثبت می‌شوند. "
            "برای نمونه، زمان پاسخ جست‌وجو به تفکیک تأمین‌کننده و به‌صورت صدکی گزارش می‌شود و نرخ خطای پرداخت نیز بر اساس نتیجه راستی‌آزمایی و کدهای خطا قابل استخراج است. "
            "برای هر سفارش و تراکنش هم یک شناسه ردیابی در لاگ‌ها ثبت می‌شود تا مسیر رخدادها از ورود درخواست تا صدور قابل دنبال‌کردن باشد. "
            "معیارهای سنجش این سند چنین است: زمان پاسخ جست‌وجو (صدک ۹۵) کمتر از ۳ ثانیه، نرخ خطای بازگشت بانک کمتر از ۱ درصد، "
            "و ثبت رویدادهای کلیدی خرید، پرداخت و صدور برای هر سفارش.",
            jc="both",
        )
    )
    el.append(
        make_tbl(
            ["سنجه", "منبع داده", "نقطه اندازه‌گیری", "گزارش/داشبورد"],
            [
                ["زمان پاسخ جست‌وجو (صدک ۹۵)", "لاگ درخواست و زمان‌بندی", "API جست‌وجو و یکپارچه‌سازی تأمین‌کنندگان", "گزارش روزانه به تفکیک تأمین‌کننده و مسیر درخواست"],
                ["نرخ خطای بازگشت بانک", "لاگ پرداخت و نتیجه راستی‌آزمایی", "مسیر بازگشت و راستی‌آزمایی پرداخت", "هشدار در صورت عبور از آستانه و گزارش هفتگی"],
                ["نرخ خطای صدور", "لاگ صدور و پاسخ تأمین‌کننده", "یکپارچه‌سازی صدور", "گزارش وضعیت صدور به تفکیک تأمین‌کننده"],
                ["تعداد سفارش‌های نیازمند بررسی", "وضعیت سفارش و رویدادهای ممیزی", "دامنه سفارش/پرداخت", "نمای عملیاتی برای پشتیبانی و عملیات"],
            ],
            col_weights=[2, 2, 2, 2],
        )
    )

    el.append(make_p("تصمیم‌های معماری (خلاصه)", style="Heading2"))
    el.append(
        make_tbl(
            ["تصمیم", "دلیل", "پیامد"],
            [
                ["تک‌استقرار ماژولار در فاز اول", "سرعت توسعه و هزینه عملیاتی کمتر", "امکان گذار به ریزخدمت با مرزبندی دامنه"],
                ["هماهنگ‌ساز خرید (فرایند مرحله‌ای)", "پرهیز از تراکنش توزیع‌شده", "نیاز به سناریوهای جبرانی و تلاش‌مجدد"],
                ["مبدّل برای تأمین‌کنندگان", "ناهمگنی و تغییرپذیری بالا", "کاهش ریسک تغییرات با هزینه توسعه مبدل"],
                ["حافظه نهان برای جست‌وجو", "کاهش زمان پاسخ و هزینه تماس بیرونی", "نیاز به زمان اعتبار و سیاست بی‌اعتبارسازی"],
                ["کلید یکتایی عملیات پرداخت", "جلوگیری از پرداخت/صدور تکراری", "ذخیره و کنترل کلیدها در پایگاه داده"],
            ],
            col_weights=[2, 2, 3],
        )
    )

    el.append(make_p("گزینه‌های بررسی‌شده و ردشده", style="Heading2"))
    el.append(
        make_p(
            "در چند نقطه کلیدی، بیش از یک گزینه قابل اجرا وجود داشت. در این بخش، گزینه‌های رایج و دلیل رد شدن آن‌ها "
            "به‌طور خلاصه ثبت می‌شود تا در آینده، همان بحث‌ها تکرار نشود و اگر شرایط تغییر کرد، بازنگری هدفمند باشد.",
            jc="both",
        )
    )
    el.append(
        make_tbl(
            ["موضوع", "گزینه ردشده", "دلیل رد شدن", "یادداشت"],
            [
                ["پرداخت/صدور", "تراکنش توزیع‌شده سراسری", "پیچیدگی عملیاتی و ریسک شکست بالا در وابستگی‌های بیرونی", "به‌جای آن از فرایند مرحله‌ای و جبران استفاده شده است."],
                ["یکپارچه‌سازی تأمین‌کننده", "اتصال مستقیم بدون مبدل", "تغییرپذیری زیاد قراردادها و سخت شدن نگهداشت", "مبدل‌ها هزینه توسعه دارند ولی ریسک تغییرات را کنترل می‌کنند."],
                ["کارایی جست‌وجو", "عدم استفاده از حافظه نهان", "افزایش هزینه تماس بیرونی و افت تجربه کاربر", "زمان اعتبار کوتاه و بی‌اعتبارسازی کنترل‌شده در نظر گرفته شده است."],
                ["استقرار", "ریزخدمت از روز اول", "هزینه عملیاتی و پیچیدگی هماهنگی بالا برای تیم کوچک", "مرزبندی دامنه‌ها حفظ شده تا گذار تدریجی امکان‌پذیر باشد."],
            ],
            col_weights=[1, 2, 2, 2],
        )
    )

    el.append(make_p("ریسک‌های معماری و برنامه کنترل", style="Heading2"))
    el.append(
        make_p(
            "ریسک‌های اصلی این سامانه بیشتر از جنس وابستگی به سرویس‌های بیرونی و خطاهای مالی است. "
            "در جدول زیر، ریسک‌های مهم، پیامدها و کنترل‌ها فهرست شده است تا در آزمون و عملیات به‌صورت فعال پیگیری شوند.",
            jc="both",
        )
    )
    el.append(
        make_tbl(
            ["ریسک", "پیامد", "کنترل/کاهش", "نشانه‌های پایش"],
            [
                ["کندی/قطع تأمین‌کننده", "افت تجربه کاربر و کاهش نرخ تبدیل", "مهلت زمانی، کاهش سطح خدمت، تلاش‌مجدد کنترل‌شده", "افزایش زمان پاسخ یا نرخ خطا به تفکیک تأمین‌کننده"],
                ["بازگشت تکراری بانک", "پرداخت/صدور تکراری و زیان مالی", "کلید یکتایی عملیات + راستی‌آزمایی پرداخت", "تعداد بازگشت‌های تکراری و اختلاف وضعیت پرداخت"],
                ["شکست صدور پس از پرداخت موفق", "نارضایتی کاربر و بار پشتیبانی", "ثبت وضعیت نیازمند بررسی، تلاش‌مجدد کنترل‌شده، اطلاع‌رسانی", "افزایش سفارش‌های پرداخت تأیید شد اما صدور ناموفق"],
                ["افشای داده حساس", "ریسک حقوقی و از دست رفتن اعتماد", "حداقل‌سازی داده، رمزنگاری ستون‌های حساس، ممیزی", "رخدادهای دسترسی نامعمول و خطاهای احراز هویت"],
            ],
            col_weights=[2, 2, 2, 2],
        )
    )

    el.append(make_p("ردیابی‌پذیری تصمیم‌ها", style="Heading2"))
    el.append(
        make_p(
            "برای اینکه بتوان تصمیم‌ها را از روی سناریوها و سنجه‌ها پیگیری کرد، چند ارجاع اصلی در نظر گرفته شده است. "
            "به‌عنوان نمونه، تصمیم «کلید یکتایی عملیات» در سناریوی خرید و در بخش امنیت و عملیات تکرار شده تا اجرای آن قابل پیگیری باشد.",
            jc="both",
        )
    )
    el.append(
        make_tbl(
            ["موضوع", "جایی که توضیح داده شده", "جایی که آزمون/سنجش می‌شود"],
            [
                ["کلید یکتایی عملیات", "تصمیم‌های معماری، امنیت، دید سناریوها", "گزارش بازگشت تکراری بانک و آزمون سناریوهای B3"],
                ["کاهش سطح خدمت", "محدودیت‌ها، دید سناریوها", "گزارش نرخ حذف تأمین‌کننده و زمان پاسخ"],
                ["حافظه نهان جست‌وجو", "تصمیم‌های معماری، کارایی", "سنجش زمان پاسخ و نرخ Cache Hit"],
            ],
            col_weights=[2, 2, 2],
        )
    )

    # 4) دید سناریوها
    el.append(make_p("دید سناریوها", style="Heading1"))
    el.append(make_p("عینیت‌بخشی موارد کاربری", style="Heading2"))
    el.append(make_fig_marker("4-1"))
    el.append(fig_caption("شکل ۴-۱: نمودار موردکاربری در سطح سیستم (جامع)."))

    el.append(make_fig_marker("4-2"))
    el.append(fig_caption("شکل ۴-۲: نمودار توالی UC-01 (Cache Hit)."))

    el.append(make_fig_marker("4-3"))
    el.append(fig_caption("شکل ۴-۳: نمودار توالی UC-01 (Cache Miss + چند تأمین‌کننده + کنترل خطا)."))

    el.append(make_fig_marker("4-4"))
    el.append(fig_caption("شکل ۴-۴: نمودار فعالیت UC-01 (اعتبارسنجی، Cache، کاهش سطح خدمت، صفحه‌بندی)."))

    el.append(make_fig_marker("4-5"))
    el.append(fig_caption("شکل ۴-۵: نمودار توالی UC-02 (شروع خرید تا شروع پرداخت)."))

    el.append(make_fig_marker("4-6"))
    el.append(fig_caption("شکل ۴-۶: نمودار توالی UC-02 (بازگشت بانک و راستی‌آزمایی پرداخت)."))

    el.append(make_fig_marker("4-7"))
    el.append(fig_caption("شکل ۴-۷: نمودار توالی UC-02 (صدور، اعلان و مسیر جبرانی)."))

    el.append(make_fig_marker("4-8"))
    el.append(fig_caption("شکل ۴-۸: نمودار فعالیت UC-02 (با مسیرهای استثنا)."))

    el.append(make_fig_marker("4-9"))
    el.append(fig_caption("شکل ۴-۹: نمودار حالت سفارش (چرخه عمر سفارش از ایجاد تا پرداخت و صدور)."))
    el.append(make_p("حداقل اجزای مشخصات سناریو", style="Heading2"))
    el.append(
        make_tbl(
            ["مولفه", "توضیح"],
            [
                ["پیش‌شرط", "وضعیت‌هایی که قبل از شروع موردکاربری باید برقرار باشند (مثلاً ورود کاربر برای خرید)."],
                ["پس‌شرط", "نتیجه موفق و وضعیت سامانه پس از پایان (مثلاً ثبت سفارش و صدور بلیت)."],
                ["جریان اصلی", "گام‌های متوالی در حالت موفق از نگاه کاربر و سامانه."],
                ["جریان‌های جایگزین/خطا", "سناریوهای تغییر قیمت، عدم موجودی، شکست پرداخت، بازگشت تکراری بانک و سایر خطاهای عملیاتی."],
                ["قواعد یکتایی عملیات", "قانون استفاده از کلید یکتایی عملیات برای جلوگیری از ثبت/پرداخت/صدور تکراری."],
            ],
            col_weights=[1, 3],
        )
    )
    el.append(make_p("فهرست موارد کاربری سطح سیستم"))
    el.append(
        make_tbl(
            ["دسته‌بندی", "شناسه", "کنشگر", "موردکاربری کلان", "توضیح"],
            [
                ["بلیت", "UC-01", "مشتری", "جست‌وجوی خدمات سفر", "جست‌وجو/فیلتر/مرتب‌سازی گزینه‌ها"],
                ["بلیت", "UC-02", "مشتری", "خرید بلیت", "بازبینی ظرفیت/قیمت، پرداخت و صدور"],
                ["بلیت", "UC-03", "مشتری", "استرداد بلیت", "لغو و بازگشت وجه"],
                ["گردشگری", "UC-04", "مشتری", "رزرو تور", "انتخاب، پرداخت و تأیید رزرو"],
                ["گردشگری", "UC-05", "مشتری", "رزرو اقامتگاه/هتل", "انتخاب، پرداخت و تأیید رزرو"],
                ["خدمات جانبی", "UC-06", "مشتری", "بیمه مسافرتی", "محاسبه، پرداخت و صدور بیمه‌نامه"],
                ["خدمات جانبی", "UC-07", "مشتری", "درخواست/پیگیری ویزا", "ثبت درخواست، پرداخت و پیگیری"],
                ["پرداخت", "UC-08", "مشتری", "کیف‌پول", "افزایش موجودی، پرداخت و تراکنش"],
                ["پشتیبانی", "UC-09", "مشتری", "پشتیبانی", "ثبت و پیگیری تیکت"],
                ["بازخورد", "UC-10", "مشتری", "نظر/امتیاز", "ثبت امتیاز پس از خرید"],
            ],
        )
    )

    el.append(make_p("موردکاربری ۱: UC-01 جست‌وجوی خدمات سفر", style="Heading2"))
    el.append(
        make_p(
            "این موردکاربری جریان جست‌وجوی خدمات سفر را پوشش می‌دهد؛ کاربر معیارهای سفر را وارد می‌کند و سامانه، "
            "نتیجه چند تأمین‌کننده را دریافت و به‌صورت یکپارچه نمایش می‌دهد.",
            jc="both",
        )
    )
    el.append(make_label("کنشگرها"))
    el.append(make_tbl(["نقش", "توضیح"], [["کاربر", "ارسال معیارهای جست‌وجو و مشاهده نتایج"], ["تأمین‌کننده", "ارائه گزینه‌های سفر از طریق API"]], col_weights=[1, 3]))
    el.append(make_label("پیش‌شرط‌ها"))
    el.append(make_p("کاربر به سامانه دسترسی دارد و سرویس‌های تأمین‌کننده در دسترس هستند.", jc="both"))
    el.append(make_label("پس‌شرط‌ها"))
    el.append(make_p("نتایج جست‌وجو به کاربر نمایش داده می‌شود و در صورت امکان، نتیجه برای مدت کوتاه در حافظه نهان ذخیره می‌گردد.", jc="both"))
    el.append(make_label("جریان اصلی"))
    el.append(
        make_tbl(
            ["گام", "شرح"],
            [
                ["۱", "کاربر معیارهای جست‌وجو (مبدا، مقصد، تاریخ، تعداد مسافر و ...) را وارد و درخواست را ارسال می‌کند."],
                ["۲", "سامانه ورودی را اعتبارسنجی می‌کند (قالب تاریخ، کامل بودن داده‌های ضروری و ...) و کلید جست‌وجو را می‌سازد."],
                ["۳", "سامانه ابتدا حافظه نهان را بررسی می‌کند؛ اگر نتیجه معتبر موجود باشد، به گام ۷ می‌رود."],
                ["۴", "در صورت نبود نتیجه در حافظه نهان، سامانه درخواست‌های همزمان به API تأمین‌کنندگان ارسال می‌کند. برای جلوگیری از کند شدن تجربه کاربر، مهلت زمانی هر درخواست کنترل می‌شود."],
                ["۵", "سامانه پاسخ‌ها را یکپارچه‌سازی می‌کند (یکسان‌سازی قالب داده، حذف موارد تکراری، مرتب‌سازی و اعمال فیلترها). در صورت نیاز، نتایج ناقص نیز با اعلام کاهش سطح خدمت نمایش داده می‌شود."],
                ["۶", "سامانه نتیجه را با زمان اعتبار مناسب در حافظه نهان ذخیره می‌کند."],
                ["۷", "سامانه نتایج صفحه‌بندی‌شده را به کاربر نمایش می‌دهد."],
            ],
            col_weights=[1, 3],
        )
    )
    el.append(make_label("جریان‌های جایگزین و خطا"))
    el.append(
        make_tbl(
            ["شناسه", "شرط/رویداد", "رفتار سامانه"],
            [
                ["A1", "ورودی نامعتبر", "سامانه خطای اعتبارسنجی را برمی‌گرداند و از ارسال درخواست به تأمین‌کنندگان خودداری می‌کند."],
                ["A2", "پاسخ دیرهنگام/عدم پاسخ یک تأمین‌کننده", "سامانه آن تأمین‌کننده را از نتیجه حذف می‌کند و نتیجه باقی‌مانده را با پیام هشدار/کاهش سطح خدمت نمایش می‌دهد."],
                ["A3", "خطای موقت تأمین‌کننده", "سامانه یک تلاش‌مجدد کنترل‌شده انجام می‌دهد؛ در صورت تداوم خطا، نتیجه بدون آن تأمین‌کننده ارائه می‌شود."],
            ],
            col_weights=[1, 2, 3],
        )
    )
    el.append(make_p("نیازمندی‌های غیرعملکردی مرتبط"))
    el.append(
        make_p(
            "هدف عملیاتی این است که زمان پاسخ جست‌وجو در شرایط معمول زیر ۳ ثانیه باشد (صدک ۹۵). برای اینکه این هدف قابل پیگیری بماند، "
            "برای هر جست‌وجو ثبت رویداد و شناسه ردیابی انجام می‌شود تا تفکیک زمان پاسخ هر تأمین‌کننده و تحلیل خطاها ممکن باشد. "
            "برای هر جست‌وجو، زمان شروع و پایان، نام تأمین‌کننده و نتیجه (موفق/ناموفق) ثبت می‌شود. "
            "همچنین نرخ خطای هر تأمین‌کننده به‌صورت دوره‌ای گزارش می‌شود تا تصمیم‌های عملیاتی (مثل محدودسازی یا قطع موقت) قابل انجام باشد.",
            jc="both",
        )
    )
    el.append(make_p("موردکاربری ۲: UC-02 خرید بلیت", style="Heading2"))
    el.append(
        make_p(
            "این موردکاربری جریان خرید بلیت را پوشش می‌دهد و چند گام وابسته (بازبینی ظرفیت و قیمت، پرداخت، بازگشت بانک، راستی‌آزمایی و صدور) "
            "و مدیریت خطاهای حساس را شامل می‌شود.",
            jc="both",
        )
    )
    el.append(make_label("کنشگرها"))
    el.append(
        make_tbl(
            ["نقش", "توضیح"],
            [
                ["کاربر", "انتخاب گزینه سفر، ورود اطلاعات مسافر و انجام پرداخت"],
                ["تأمین‌کننده", "بازبینی ظرفیت/قیمت و صدور بلیت از طریق API"],
                ["درگاه پرداخت", "شروع پرداخت، بازگشت پرداخت و راستی‌آزمایی تراکنش"],
                ["سرویس اعلان", "ارسال پیامک/ایمیل تأیید خرید"],
            ],
            col_weights=[1, 3],
        )
    )
    el.append(make_label("پیش‌شرط‌ها"))
    el.append(make_p("کاربر وارد سامانه شده است و یک گزینه سفر معتبر انتخاب کرده است.", jc="both"))
    el.append(make_label("پس‌شرط‌ها"))
    el.append(
        make_p(
            "در حالت موفق، سفارش در وضعیت صدور انجام شد قرار می‌گیرد و بلیت/کد پیگیری به کاربر اعلام می‌شود. "
            "در حالت ناموفق، وضعیت سفارش طوری ثبت می‌شود که پیگیری، تلاش‌مجدد کنترل‌شده یا اقدام پشتیبانی ممکن باشد.",
            jc="both",
        )
    )
    el.append(make_label("قواعد کلیدی"))
    el.append(
        make_tbl(
            ["قانون", "توضیح"],
            [
                ["بازبینی قبل از پرداخت", "قبل از شروع پرداخت، ظرفیت و قیمت دوباره از تأمین‌کننده استعلام می‌شود تا از خرید با داده قدیمی جلوگیری شود."],
                ["کلید یکتایی عملیات", "برای جلوگیری از پرداخت/ثبت تکراری، روی شروع پرداخت و بازگشت بانک کلید یکتا اعمال می‌شود."],
                ["راستی‌آزمایی پرداخت", "صرفاً بازگشت بانک کافی نیست؛ وضعیت پرداخت از طریق API درگاه راستی‌آزمایی می‌شود و نتیجه در سفارش ثبت می‌گردد."],
            ],
            col_weights=[1, 3],
        )
    )
    el.append(make_label("جریان اصلی"))
    el.append(
        make_tbl(
            ["گام", "شرح"],
            [
                ["۱", "کاربر گزینه سفر را انتخاب کرده و اطلاعات مسافران را وارد می‌کند."],
                ["۲", "سامانه اطلاعات مسافر را اعتبارسنجی می‌کند و درخواست بازبینی ظرفیت/قیمت را به تأمین‌کننده ارسال می‌کند."],
                ["۳", "در صورت تأیید، سامانه یک سفارش با وضعیت در انتظار پرداخت ایجاد می‌کند و کلید یکتایی عملیات پرداخت را ثبت می‌کند."],
                ["۴", "سامانه درخواست شروع پرداخت را به درگاه پرداخت ارسال و کاربر را به صفحه پرداخت هدایت می‌کند."],
                ["۵", "کاربر پرداخت را انجام می‌دهد و درگاه پرداخت کاربر/سامانه را به مسیر بازگشت هدایت می‌کند."],
                ["۶", "سامانه بازگشت بانک را دریافت می‌کند و با استفاده از کلید یکتایی عملیات، از پردازش تکراری جلوگیری می‌کند."],
                ["۷", "سامانه پرداخت را از طریق API درگاه راستی‌آزمایی می‌کند؛ در صورت موفقیت، وضعیت سفارش به پرداخت تأیید شد تغییر می‌کند."],
                ["۸", "سامانه درخواست صدور را به تأمین‌کننده ارسال می‌کند؛ در صورت موفقیت، بلیت/کد پیگیری ثبت و وضعیت سفارش به صدور انجام شد تغییر می‌کند."],
                ["۹", "سامانه اعلان تأیید خرید را برای کاربر ارسال می‌کند."],
            ],
            col_weights=[1, 3],
        )
    )
    el.append(make_label("جریان‌های جایگزین و خطا"))
    el.append(
        make_tbl(
            ["شناسه", "شرط/رویداد", "رفتار سامانه"],
            [
                ["B1", "عدم تأیید ظرفیت/تغییر قیمت در بازبینی", "سامانه خرید را متوقف می‌کند و پیام تغییر قیمت/عدم موجودی را به کاربر نمایش می‌دهد."],
                ["B2", "شکست پرداخت یا انصراف کاربر", "سامانه وضعیت سفارش را ناموفق/لغو ثبت می‌کند و امکان تلاش مجدد را فراهم می‌نماید."],
                ["B3", "بازگشت تکراری بانک", "سامانه با کلید یکتایی عملیات، فقط یک‌بار پردازش را انجام می‌دهد و درخواست‌های تکراری را بی‌اثر می‌کند."],
                ["B4", "راستی‌آزمایی ناموفق/مبهم", "سامانه وضعیت سفارش را نیازمند بررسی ثبت می‌کند و در صورت نیاز تیکت پشتیبانی ایجاد می‌شود."],
                ["B5", "شکست صدور پس از پرداخت موفق", "سامانه وضعیت پرداخت تأیید شد، اما صدور ناموفق را ثبت می‌کند؛ تلاش‌مجدد کنترل‌شده انجام می‌دهد و در صورت تداوم، پشتیبانی را مطلع می‌کند."],
            ],
            col_weights=[1, 2, 3],
        )
    )
    # نمودارهای UC-01/UC-02 در ابتدای این بخش به‌صورت شکل ۴-۲ و شکل ۴-۳ ارجاع داده شده‌اند.

    # 5) دید منطقی
    el.append(make_p("دید منطقی", style="Heading1"))
    el.append(make_p("مدل لایه‌ای", style="Heading2"))
    el.append(
        make_p(
            "در نمای منطقی، سامانه به پنج لایه اصلی تقسیم می‌شود: "
            "لایه ارائه (کلاینت‌های وب و موبایل) برای نمایش و تعامل، "
            "لایه API برای ورودی و خروجی، اعتبارسنجی و سیاست‌های عمومی، "
            "لایه دامنه و خدمات برای پیاده‌سازی منطق کسب‌وکار، "
            "لایه یکپارچه‌سازی برای تطبیق قراردادها و مدیریت خطاهای سرویس‌های بیرونی، "
            "و در نهایت لایه داده و زیرساخت برای مدیریت پایگاه داده، حافظه نهان و کارهای پس‌زمینه.",
            jc="both",
        )
    )
    el.append(make_p("دامنه‌های مرزبندی‌شده", style="Heading2"))
    el.append(
        make_tbl(
            ["دامنه", "مسئولیت", "داده‌های اصلی"],
            [
                ["هویت", "ورود/ثبت‌نام/پروفایل", "کاربر، نشست"],
                ["جست‌وجو", "جست‌وجو و نتایج", "کلید جست‌وجو، گزینه سفر"],
                ["سفارش", "سفارش و وضعیت‌ها", "سفارش، مسافر"],
                ["پرداخت", "تراکنش و راستی‌آزمایی پرداخت", "تراکنش، کلید یکتایی عملیات"],
                ["صدور", "صدور بلیت/واچر", "بلیت"],
                ["گردشگری", "تور/اقامتگاه", "رزرو تور، رزرو اقامت"],
                ["کیف‌پول", "موجودی و تراکنش کیف‌پول", "کیف‌پول، تراکنش کیف‌پول"],
                ["پشتیبانی", "تیکت و پیگیری", "تیکت پشتیبانی"],
                ["بازخورد", "امتیاز و بازخورد", "بازخورد"],
            ],
        )
    )
    el.append(make_p("CRC (تحلیل) — مشترک", style="Heading2"))
    el.append(
        make_tbl(
            ["کلاس", "مسئولیت‌ها", "همکاران"],
            [
                ["SearchController", "دریافت ورودی جست‌وجو، اعتبارسنجی اولیه، مدیریت خروجی صفحه‌بندی", "SearchService، CacheClient"],
                ["SearchService", "اجرای جست‌وجو، تجمیع نتایج، حذف تکراری، مرتب‌سازی/فیلتر", "ProviderAdapter، CacheClient"],
                ["BookingController", "شروع خرید، دریافت اطلاعات مسافر، ایجاد سفارش و نمایش وضعیت", "BookingService، PaymentService"],
                ["BookingService", "ساخت/به‌روزرسانی سفارش، مدیریت چرخه وضعیت، قواعد انقضا", "BookingRepository، ProviderAdapter"],
                ["PaymentService", "شروع پرداخت، کنترل یکتایی عملیات، راستی‌آزمایی تراکنش", "PaymentGatewayClient، TransactionRepository"],
                ["IssueService", "ارسال درخواست صدور و ثبت نتیجه، مدیریت تلاش‌مجدد", "ProviderAdapter، NotificationService"],
                ["RefundService", "بررسی قوانین استرداد، ثبت درخواست و مدیریت بازپرداخت", "ProviderAdapter، PaymentGatewayClient، WalletService"],
                ["ProviderAdapter", "تطبیق قرارداد هر تأمین‌کننده، نگاشت داده و مدیریت خطا", "HTTPClient، ProviderMapper"],
                ["NotificationService", "ارسال پیامک/ایمیل و پیگیری خطاهای ارسال", "NotifyClient، OutboxRepository"],
                ["SupportService", "ثبت و پیگیری تیکت‌های عملیاتی", "SupportClient، AuditLog"],
            ],
            col_weights=[2, 4, 3],
        )
    )
    el.append(make_fig_marker("5-3"))
    el.append(fig_caption("شکل ۵-۳: کارت‌های CRC (تحلیل) — نمایش خلاصه مسئولیت‌ها و همکاران."))

    el.append(make_p("الگوهای طراحی کلیدی", style="Heading2"))
    el.append(
        make_p(
            "برای کاهش پیچیدگی و افزایش توسعه‌پذیری، در طراحی سامانه از الگوهای زیر استفاده می‌شود: "
            "الگوی نما برای ارائه نقاط ورود یکتا به جریان‌های مهم (جست‌وجو و خرید)، "
            "الگوی راهبرد برای سیاست‌های قابل تغییر (رتبه‌بندی، قیمت‌گذاری، قوانین استرداد)، "
            "الگوی روش کارخانه و مبدل برای اضافه/تعویض کردن تأمین‌کنندگان و درگاه‌ها با حداقل تغییرات، "
            "و در نهایت هماهنگ‌ساز فرایند خرید برای مدیریت مراحل خرید بدون نیاز به تراکنش توزیع‌شده.",
            jc="both",
        )
    )
    el.append(make_fig_marker("5-1"))
    el.append(fig_caption("شکل ۵-۱: نمودار کلاس (تحلیلی) — کلاس‌های کلیدی و رابطه‌ها."))
    el.append(make_fig_marker("5-2"))
    el.append(fig_caption("شکل ۵-۲: نمودار کلاس (طراحی) — تمرکز روی رابط‌ها و عملیات."))

    # 6) دید فرایند
    el.append(make_p("دید فرایند", style="Heading1"))
    el.append(make_p("همزمانی و زمان‌بندی", style="Heading2"))
    el.append(
        make_p(
            "در جست‌وجو، فراخوانی تأمین‌کنندگان به‌صورت موازی انجام می‌شود و برای کنترل زمان پاسخ، "
            "مهلت زمانی پاسخ و سیاست کاهش سطح خدمت اعمال می‌گردد. "
            "در خرید، یک سفارش موقت با زمان انقضا ثبت می‌شود تا هم وضعیت سفارش قابل ردیابی باشد و هم منابع به‌صورت نامحدود قفل نشوند. "
            "پس از راستی‌آزمایی پرداخت، فرآیند صدور اجرا می‌گردد. "
            "در استرداد، ابتدا امکان‌پذیری و قوانین بررسی می‌شود و سپس بازپرداخت از مسیر مناسب (درگاه یا کیف‌پول) انجام می‌گردد.",
            jc="both",
        )
    )
    el.append(make_p("کارهای پس‌زمینه (اختیاری)", style="Heading2"))
    el.append(
        make_tbl(
            ["کار", "دلیل", "سیاست"],
            [
                ["تلاش‌مجدد صدور بلیت", "خطای موقت تأمین‌کننده", "۳ بار تلاش با افزایش تدریجی فاصله زمانی"],
                ["پاکسازی رزروهای منقضی", "آزادسازی منابع", "هر ۵ دقیقه"],
                ["ارسال مجدد اعلان‌های ناموفق", "پایداری پیامک/ایمیل", "تلاش‌مجدد + صف پیام‌های مرده"],
            ],
        )
    )
    el.append(make_p("یکتایی عملیات و سازگاری", style="Heading2"))
    el.append(
        make_p(
            "برای کنترل اجرای تکراری عملیات (به‌ویژه در بازگشت بانک و تلاش‌مجددها)، کلید یکتایی عملیات محور کار است. "
            "برای اینکه اعلان‌ها در خطاهای موقت از بین نروند، به‌کارگیری الگوی صندوق خروجی کمک می‌کند. "
            "ثبت شناسه ردیابی برای سفارش و تراکنش هم امکان پیگیری انتها به انتها را فراهم می‌کند.",
            jc="both",
        )
    )

    # 7) دید فیزیکی
    el.append(make_p("دید فیزیکی (استقرار)", style="Heading1"))
    el.append(make_fig_marker("7-1"))
    el.append(fig_caption("شکل ۷-۱: نمودار استقرار."))

    # 8) دید توسعه و پیاده‌سازی
    el.append(make_p("دید توسعه و پیاده‌سازی", style="Heading1"))
    el.append(make_p("ساختار ماژول‌ها", style="Heading2"))
    el.append(
        make_tbl(
            ["ماژول/بسته", "توضیح", "فناوری‌ها", "وابستگی‌ها"],
            [
                ["رابط وب (ui-web)", "نمایش و تعامل کاربر در وب", "React", "سامانه سمت سرور"],
                ["رابط موبایل (ui-mobile)", "نمایش و تعامل کاربر در موبایل", "Flutter", "سامانه سمت سرور"],
                ["سامانه سمت سرور (api)", "ارائه خدمات و منطق کسب‌وکار", "Python (FastAPI)", "داده/حافظه نهان/یکپارچه‌سازی/کارهای پس‌زمینه"],
                ["پایگاه داده (db)", "ذخیره‌سازی پایدار", "PostgreSQL", "ندارد"],
                ["حافظه نهان (cache)", "کاهش زمان پاسخ و بار روی تأمین‌کننده/داده", "Redis", "ندارد"],
                ["کارهای پس‌زمینه (queue)", "تلاش‌مجدد، زمان‌بندی و پردازش ناهمگام", "Celery/RQ", "Redis"],
                ["دامنه‌ها (domain-*)", "بسته‌های منطق دامنه", "بسته‌های پایتون", "مخزن‌ها/یکپارچه‌سازی"],
                ["یکپارچه‌سازی تأمین‌کنندگان (integrations-providers)", "ارتباط و تبدیل داده", "کلاینت HTTP + نگاشت داده", "ندارد"],
                ["یکپارچه‌سازی پرداخت (integrations-bank)", "شروع/بازگشت/راستی‌آزمایی", "کلاینت HTTP + بررسی امضا", "ندارد"],
                ["یکپارچه‌سازی پشتیبانی (integrations-support)", "تیکت/پیگیری", "کلاینت HTTP", "ندارد"],
                ["یکپارچه‌سازی اعلان (integrations-notify)", "پیامک/ایمیل", "کلاینت HTTP", "ندارد"],
            ],
            col_weights=[2, 2, 2, 2],
        )
    )
    el.append(make_p("قراردادهای API (نمای کلی)", style="Heading2"))
    el.append(
        make_tbl(
            ["مسیر خدمت", "توضیح"],
            [
                ["GET /search", "جستجو و نتایج صفحه‌بندی"],
                ["POST /bookings", "ایجاد سفارش/شروع پرداخت"],
                ["POST /payments/callback", "بازگشت بانک و راستی‌آزمایی پرداخت"],
                ["POST /refunds", "ثبت استرداد"],
                ["POST /wallet/topup", "افزایش موجودی کیف پول"],
                ["POST /support/tickets", "ثبت تیکت"],
            ],
        )
    )

    # بهره‌برداری و عملیات (برای اینکه سند فقط توسعه‌محور نباشد)
    el.append(make_p("بهره‌برداری و عملیات", style="Heading1"))
    el.append(make_p("پایش و هشداردهی", style="Heading2"))
    el.append(
        make_p(
            "برای جلوگیری از غافلگیر شدن در مسیرهای حساس، پایش باید روی سنجه‌های تجربه کاربر و سنجه‌های وابستگی‌های بیرونی متمرکز باشد. "
            "به‌طور خاص، جست‌وجو و پرداخت باید هم از منظر کارایی و هم از منظر خطاها به‌صورت پیوسته زیر نظر باشد و در صورت عبور از آستانه‌ها، "
            "هشدار عملیاتی ایجاد شود. این هشدارها باید قابل اقدام باشند؛ یعنی همراه با شناسه ردیابی و زمینه کافی ثبت شوند.",
            jc="both",
        )
    )
    el.append(
        make_tbl(
            ["حوزه", "سنجه/هشدار", "آستانه", "اقدام"],
            [
                ["جست‌وجو", "زمان پاسخ صدک ۹۵", "بیش از ۳ ثانیه", "بررسی تأمین‌کننده‌های کند، فعال‌سازی کاهش سطح خدمت"],
                ["تأمین‌کنندگان", "نرخ خطا", "بیش از ۵٪ در ۱۰ دقیقه", "محدودسازی یا قطع موقت تأمین‌کننده و ثبت رخداد"],
                ["پرداخت", "راستی‌آزمایی ناموفق/مبهم", "افزایش غیرعادی", "بررسی درگاه، فعال‌سازی مسیر نیازمند بررسی"],
                ["صدور", "صدور ناموفق پس از پرداخت موفق", "بیش از ۱٪", "تلاش‌مجدد کنترل‌شده و اطلاع‌رسانی به پشتیبانی"],
            ],
            col_weights=[1, 2, 1, 2],
        )
    )

    el.append(make_p("پشتیبان‌گیری و بازیابی", style="Heading2"))
    el.append(
        make_p(
            "برای داده‌های تراکنشی، داشتن برنامه پشتیبان‌گیری و بازیابی ضروری است. "
            "هدف این برنامه این است که در رخدادهای عملیاتی، داده از دست نرود و سامانه در زمان قابل قبول به کار برگردد. "
            "این بخش، اهداف عملیاتی را مشخص می‌کند تا در پیاده‌سازی و عملیات مبنا قرار گیرد.",
            jc="both",
        )
    )
    el.append(
        make_tbl(
            ["موضوع", "هدف", "توضیح"],
            [
                ["هدف زمان بازیابی (RTO)", "۱ ساعت", "حداکثر زمان قابل قبول برای بازگشت سرویس پس از رخداد"],
                ["هدف نقطه بازیابی (RPO)", "۱۵ دقیقه", "حداکثر داده قابل از دست‌رفتن در بدترین حالت"],
                ["پشتیبان‌گیری پایگاه داده", "روزانه + نقطه‌ای", "تهیه نسخه کامل و امکان بازیابی نقطه‌ای در بازه کوتاه"],
                ["تمرین بازیابی", "ماهیانه", "بازیابی آزمایشی برای اطمینان از قابل اجرا بودن فرایند"],
            ],
            col_weights=[2, 1, 3],
        )
    )

    el.append(make_p("مدیریت رخداد و پیگیری", style="Heading2"))
    el.append(
        make_p(
            "برای رخدادهای مالی و عملیاتی، مهم است که ثبت رخداد و پیگیری، به وضعیت‌های قابل فهم تبدیل شود. "
            "به همین دلیل، وضعیت‌هایی مانند «نیازمند بررسی» در سفارش تعریف شده است تا هم پشتیبانی و هم عملیات بتوانند بدون حدس و گمان، "
            "کار را پیگیری کنند و روند حل مسئله قابل گزارش باشد.",
            jc="both",
        )
    )
    el.append(
        make_tbl(
            ["رخداد", "داده‌های لازم برای پیگیری", "خروجی/اقدام"],
            [
                ["پرداخت مبهم", "شناسه سفارش، شناسه تراکنش، شناسه پیگیری درگاه", "انتقال به وضعیت نیازمند بررسی و ایجاد وظیفه پشتیبانی"],
                ["شکست صدور", "شناسه سفارش، نام تأمین‌کننده، کد خطا", "تلاش‌مجدد کنترل‌شده و اطلاع‌رسانی به کاربر/پشتیبانی"],
                ["کندی جست‌وجو", "شناسه ردیابی، تفکیک زمان پاسخ تأمین‌کننده‌ها", "محدودسازی تأمین‌کننده کند و ثبت رخداد عملیاتی"],
            ],
            col_weights=[1, 2, 2],
        )
    )

    el.append(make_p("ثبت رویدادها و ممیزی", style="Heading3"))
    el.append(
        make_p(
            "برای اینکه پیگیری رخدادها به حدس و گمان وابسته نباشد، رویدادهای کلیدی در دو سطح ثبت می‌شوند: "
            "ثبت عملیاتی (برای عیب‌یابی و پایش) و ثبت ممیزی (برای پاسخ‌گویی و پیگیری مالی). "
            "حداقلِ مورد انتظار این است که هر رویداد، همراه با شناسه‌های اصلی و نتیجه عملیات ثبت شود تا با یک شناسه سفارش یا تراکنش، "
            "بتوان مسیر رخداد را بازسازی کرد.",
            jc="both",
        )
    )
    el.append(
        make_tbl(
            ["رویداد", "محل ثبت", "حداقل فیلدهای ثبت‌شونده"],
            [
                ["ایجاد سفارش", "لاگ عملیاتی + ممیزی", "bookingId، userId، createdAt، status"],
                ["شروع پرداخت", "لاگ عملیاتی + ممیزی", "bookingId، transactionId، amount، idempotencyKey"],
                ["بازگشت بانک", "لاگ عملیاتی + ممیزی", "transactionId، gatewayRef، نتیجه راستی‌آزمایی، تکراری/غیرتکراری"],
                ["تأیید پرداخت", "لاگ عملیاتی + ممیزی", "bookingId، transactionId، وضعیت نهایی پرداخت، زمان تأیید"],
                ["شروع صدور", "لاگ عملیاتی", "bookingId، providerName، requestId/traceId"],
                ["نتیجه صدور", "لاگ عملیاتی + ممیزی", "bookingId، providerRef، issueStatus، کد خطا/پیام"],
                ["شروع استرداد", "لاگ عملیاتی + ممیزی", "bookingId، مبلغ بازپرداخت، علت/قانون"],
                ["نتیجه بازپرداخت", "لاگ عملیاتی + ممیزی", "transactionId، gatewayRef، refundStatus، زمان ثبت"],
            ],
            col_weights=[2, 2, 3],
        )
    )

    # 9) دید داده
    el.append(make_p("دید داده (اختیاری)", style="Heading1"))
    el.append(make_fig_marker("9-1"))
    el.append(fig_caption("شکل ۹-۱: نمودار موجودیت-رابطه (مدل داده)."))
    el.append(make_p("سیاست‌های داده", style="Heading2"))
    el.append(
        make_p(
            "در این سامانه، اصل حداقل‌سازی داده رعایت می‌شود و فقط داده‌های ضروری نگهداری خواهد شد. "
            "ارتباطات شبکه‌ای با TLS محافظت می‌شوند و در صورت نیاز، ستون‌های حساس در پایگاه داده نیز رمزنگاری می‌گردند. "
            "همچنین لاگ تراکنش‌ها و لاگ ممیزی به شکلی نگهداری می‌شوند که پاسخ‌گویی پشتیبانی و پیگیری رخدادها امکان‌پذیر باشد."
            ,
            jc="both",
        )
    )

    # 10) کارایی
    el.append(make_p("کارایی", style="Heading1"))
    el.append(make_p("راهکارهای کارایی", style="Heading2"))
    el.append(
        make_p(
            "برای دستیابی به کارایی مناسب، به‌کارگیری حافظه نهان برای نتایج جست‌وجو با زمان اعتبار کوتاه‌مدت، "
            "صفحه‌بندی و کنترل اندازه پاسخ‌ها، اعمال مهلت زمانی و تلاش‌مجدد کنترل‌شده برای سرویس‌های بیرونی، "
            "و بهینه‌سازی دسترسی به پایگاه داده با ایندکس‌ها و استخر اتصال ضروری است.",
            jc="both",
        )
    )
    el.append(make_p("روش سنجش کارایی", style="Heading2"))
    el.append(
        make_p(
            "برای سنجش کارایی، سنجه‌ها از محل‌های درست جمع‌آوری می‌شوند تا هم تجربه کاربر دیده شود و هم سهم هر وابستگی بیرونی قابل تفکیک باشد.",
            jc="both",
        )
    )
    el.append(
        make_tbl(
            ["سنجه", "جایی که اندازه‌گیری می‌شود", "چیزی که باید ثبت شود"],
            [
                ["زمان پاسخ جست‌وجو", "API جست‌وجو", "زمان کل + تفکیک زمان تماس با هر تأمین‌کننده + شناسه ردیابی"],
                ["زمان بازگشت/راستی‌آزمایی پرداخت", "API پرداخت", "زمان بازگشت بانک + زمان راستی‌آزمایی + نتیجه نهایی"],
                ["زمان صدور", "یکپارچه‌سازی صدور", "زمان درخواست تا پاسخ تأمین‌کننده + کد خطا/موفقیت"],
            ],
            col_weights=[2, 2, 3],
        )
    )
    el.append(make_p("تنظیمات اجرایی", style="Heading2"))
    el.append(
        make_tbl(
            ["موضوع", "مقدار", "توضیح"],
            [
                ["زمان اعتبار حافظه نهان جست‌وجو", "۱۰ دقیقه", "کاهش بار تماس با تأمین‌کنندگان و بهبود زمان پاسخ در بازه کوتاه"],
                ["مهلت زمانی پاسخ تأمین‌کننده", "۲٫۵ ثانیه", "عدم وابستگی تجربه کاربر به تأمین‌کننده کند؛ نمایش نتیجه با کاهش سطح خدمت"],
                ["تلاش‌مجدد تماس با تأمین‌کننده", "۱ بار", "فقط برای خطاهای موقت؛ با فاصله زمانی کوتاه"],
                ["زمان انقضای سفارشِ در انتظار پرداخت", "۱۵ دقیقه", "جلوگیری از قفل شدن منابع و مدیریت سفارش‌های نیمه‌کاره"],
                ["پنجره اعتبار کلید یکتایی عملیات پرداخت", "۲۴ ساعت", "بی‌اثر کردن بازگشت‌های تکراری و جلوگیری از پردازش دوباره"],
            ],
            col_weights=[2, 1, 3],
        )
    )
    el.append(make_p("نقاط گلوگاه و کنترل", style="Heading2"))
    el.append(
        make_tbl(
            ["گلوگاه", "ریسک", "کنترل"],
            [
                ["تأمین‌کننده کند/قطع", "افزایش زمان پاسخ", "کاهش سطح خدمت + قطع‌کننده مدار سبک"],
                ["بازگشت بانک", "تکرار/تقلب", "راستی‌آزمایی + کلید یکتایی عملیات"],
                ["صدور بلیت", "خطای موقت", "تلاش‌مجدد خودکار + تیکت پشتیبانی"],
            ],
        )
    )

    # 11) کیفیت
    el.append(make_p("کیفیت", style="Heading1"))
    el.append(make_p("امنیت", style="Heading2"))
    el.append(
        make_p(
            "احراز هویت کاربران با توکن انجام می‌شود و برای مسیرهای حساس (مثل شروع پرداخت و استرداد) محدودیت نرخ درخواست در نظر گرفته می‌شود. "
            "مجوزدهی هم بر اساس نقش‌ها (کاربر، پشتیبان، مدیر و تأمین‌کننده) اعمال می‌شود تا هر نقش فقط به قابلیت‌های مرتبط دسترسی داشته باشد.",
            jc="both",
        )
    )
    el.append(make_p("دارایی‌های حساس و تهدیدهای اصلی", style="Heading3"))
    el.append(
        make_tbl(
            ["دارایی/حوزه", "تهدیدهای رایج", "کنترل‌ها"],
            [
                ["حساب کاربری و نشست", "دسترسی غیرمجاز، حملات حدس گذرواژه", "توکن امن، محدودیت نرخ، قفل موقت، ثبت ممیزی"],
                ["پرداخت و سفارش", "پردازش تکراری، تقلب، دستکاری بازگشت", "کلید یکتایی عملیات، راستی‌آزمایی، امضا/توکن بازگشت"],
                ["داده‌های شخصی", "افشای داده، دسترسی بیش از حد", "حداقل‌سازی داده، نقش‌ها و مجوزها، رمزنگاری ستون‌های حساس"],
                ["یکپارچه‌سازی بیرونی", "قطع سرویس، پاسخ مخرب/نامعتبر", "مهلت زمانی، اعتبارسنجی پاسخ، کاهش سطح خدمت"],
            ],
            col_weights=[2, 2, 3],
        )
    )
    el.append(make_p("هدف سطح خدمت (SLO) و توافق سطح خدمت (SLA)", style="Heading2"))
    el.append(
        make_p(
            "برای اینکه کیفیت قابل داوری باشد، چند هدف سطح خدمت به‌عنوان مبنای ارزیابی در نظر گرفته شده است. "
            "این اهداف در فازهای بعدی می‌تواند به توافق عملیاتی تبدیل شود.",
            jc="both",
        )
    )
    el.append(
        make_tbl(
            ["موضوع", "هدف", "روش ارزیابی"],
            [
                ["زمان پاسخ جست‌وجو", "صدک ۹۵ کمتر از ۳ ثانیه", "اندازه‌گیری از API جست‌وجو با شناسه ردیابی"],
                ["نرخ موفقیت پرداخت", "بیش از ۹۹٪", "نتیجه راستی‌آزمایی پرداخت در بازه زمانی"],
                ["نرخ شکست صدور پس از پرداخت موفق", "کمتر از ۱٪", "شمارش سفارش‌های پرداخت تأیید شد که صدور ناموفق دارند"],
            ],
            col_weights=[2, 2, 3],
        )
    )
    el.append(
        make_p(
            "در پرداخت، صرف دریافت بازگشت بانک کافی نیست و تراکنش از طریق API درگاه راستی‌آزمایی می‌شود. "
            "برای جلوگیری از پردازش تکراری هم کلید یکتایی عملیات روی شروع پرداخت و بازگشت بانک کنترل می‌گردد.",
            jc="both",
        )
    )
    el.append(
        make_p(
            "در ذخیره‌سازی، اطلاعات حساس حداقلی نگهداری می‌شود؛ رمز عبور با هش امن ذخیره می‌شود و اطلاعات کارت در سامانه نگهداری نمی‌شود. "
            "برای پیگیری و پاسخ‌گویی، رویدادهای کلیدی مانند ایجاد سفارش، شروع پرداخت، تأیید پرداخت، صدور بلیت و استرداد با شناسه‌های مرتبط ثبت می‌شوند "
            "(برای نمونه: شناسه سفارش (bookingId)، شناسه پیگیری درگاه (gatewayRef)، و کلید یکتایی عملیات (idempotencyKey)).",
            jc="both",
        )
    )
    el.append(make_p("قابلیت نگهداری و توسعه‌پذیری", style="Heading2"))
    el.append(
        make_p(
            "بیشترین تغییرپذیری سامانه در لایه یکپارچه‌سازی رخ می‌دهد، چون هر تأمین‌کننده قرارداد و محدودیت‌های خودش را دارد. "
            "برای کنترل این تغییرپذیری، برای هر تأمین‌کننده یک مبدل مستقل در نظر گرفته می‌شود و ورودی/خروجی آن با نمونه پیام‌ها و قواعد نگاشت به‌صورت روشن ثبت می‌گردد.",
            jc="both",
        )
    )
    el.append(
        make_p(
            "برای کاهش ریسک تغییرات، آزمون قرارداد روی پیام‌های کلیدی (جست‌وجو، بازبینی، صدور) انجام می‌شود تا با تغییرات ناگهانی تأمین‌کننده سریع‌تر متوجه شکست‌ها شویم. "
            "در سمت سامانه نیز مرزبندی دامنه‌ها رعایت می‌شود تا تغییرات یک بخش، کمترین اثر را روی بخش‌های دیگر داشته باشد.",
            jc="both",
        )
    )
    el.append(make_p("مشاهده‌پذیری", style="Heading2"))
    el.append(
        make_p(
            "برای پیگیری خطاها، در مسیرهای حساس یک شناسه ردیابی تولید می‌شود و در تمام لاگ‌ها حمل می‌گردد. "
            "حداقل داده‌ای که برای پیگیری عملیاتی نیاز داریم شامل زمان پاسخ جست‌وجو به تفکیک تأمین‌کننده، نرخ خطای بازگشت و راستی‌آزمایی پرداخت، و وضعیت صدور بلیت است.",
            jc="both",
        )
    )
    el.append(
        make_p(
            "در عمل، ثبت چند شناسه کلیدی کمک می‌کند وقتی کاربر یا پشتیبانی گزارش می‌دهد، مسیر رخداد سریع‌تر دنبال شود "
            "(برای نمونه: شناسه سفارش (bookingId)، شناسه تراکنش (transactionId)، نام تأمین‌کننده (providerName)، و شناسه پیگیری درگاه (gatewayRef)). "
            "اگر زیرساخت فراهم باشد، ردیابی توزیع‌شده هم برای مسیر خرید-پرداخت-صدور اضافه می‌شود تا گلوگاه‌ها دقیق‌تر دیده شوند.",
            jc="both",
        )
    )

    return el


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--embed-images", action="store_true", help="Embed diagram PNGs into the output .docx")
    parser.add_argument("--no-autogen-diagrams", action="store_true", help="Do not auto-generate placeholder diagrams")
    parser.add_argument("--out", default="SAD-Final.docx", help="Output .docx path")
    args = parser.parse_args()

    template_path = Path("SAD-Template.docx")
    out_path = Path(args.out)
    if not template_path.exists():
        raise SystemExit("Missing SAD-Template.docx")

    with zipfile.ZipFile(template_path, "r") as zin:
        file_bytes = {name: zin.read(name) for name in zin.namelist()}

    doc_xml = file_bytes.get("word/document.xml")
    if doc_xml is None:
        raise SystemExit("Template missing word/document.xml")

    root = ET.fromstring(doc_xml)

    # Fill cover placeholders (best-effort)
    replace_first_paragraph_text(root, "سازمان ...", "سازمان آژانس مسافرتی مارکوپولو")
    replace_first_paragraph_text(root, "سامانه ...", "سامانه فروش/رزرو خدمات سفر (وب/موبایل)")
    replace_first_paragraph_text(root, "پاییز 1404", "زمستان ۱۴۰۴")
    fill_cover_group_members(root, get_group_members())

    # Fill history table
    fill_history_table(root)

    # Replace body content starting from the template's first main section heading
    body = root.find("w:body", NS)
    if body is None:
        raise SystemExit("Invalid document.xml (no w:body)")
    children = list(body)

    start_idx = None
    for i, el in enumerate(children):
        if el.tag == _qn("w:p") and _p_text(el) == "کليات سند":
            start_idx = i
            break
    if start_idx is None:
        raise SystemExit("Could not find 'کليات سند' heading in template")

    sectPr = body.find("w:sectPr", NS)
    # Remove everything from start_idx to before sectPr (or end)
    for el in children[start_idx:]:
        if sectPr is not None and el is sectPr:
            break
        body.remove(el)

    # Insert filled content
    insert_pos = start_idx
    for new_el in build_sad_content(fig_caption_red=not args.embed_images):
        body.insert(insert_pos, new_el)
        insert_pos += 1

    # Optionally embed diagrams from ./diagrams into the document.
    embed_figures(
        root,
        file_bytes,
        diagrams_dir=Path("diagrams"),
        autogen=not args.no_autogen_diagrams,
        embed_images=args.embed_images,
    )

    # Keep the template TOC field and make sure it updates on open.
    ensure_toc_field(root, file_bytes)

    # Fix header/footer placeholders (e.g., '...') after all edits.
    _update_header_footer_xml(file_bytes)

    # Write output docx (preserve all other parts)
    new_doc_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    file_bytes["word/document.xml"] = new_doc_xml

    tmp_out = out_path.with_suffix(out_path.suffix + ".tmp")
    with zipfile.ZipFile(tmp_out, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in file_bytes.items():
            zout.writestr(name, data)

    tmp_out.replace(out_path)
    print(f"Wrote {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
