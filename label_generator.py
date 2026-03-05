"""
label_generator.py – Core engine for barcode label PDF generation.

Reads Excel data, generates Code 128 barcodes, and produces a PDF
with labels arranged in a 3×8 grid on A4 sheets (70×37 mm each).
"""

import io
import json
import os
import xml.etree.ElementTree as ET
from typing import Any

import pandas as pd
import barcode
from barcode.writer import ImageWriter, SVGWriter
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

CONFIG_FILE = "config.json"

# A4 dimensions in mm
A4_WIDTH_MM = 210
A4_HEIGHT_MM = 297

# Label grid
COLS = 3
ROWS = 8
LABELS_PER_PAGE = COLS * ROWS  # 24

# Label dimensions (mm)
LABEL_WIDTH = 70.0
LABEL_HEIGHT = 37.0

# Padding within each label cell (keeps content off pre-cut borders)
LABEL_PAD_X = 2.0   # mm inset from left/right edges
LABEL_PAD_Y = 3.0   # mm inset from top/bottom edges (keeps content inside printer margins)

# Margins to center the grid on the page
MARGIN_LEFT = (A4_WIDTH_MM - COLS * LABEL_WIDTH) / 2  # ≈ 0 mm
MARGIN_TOP = (A4_HEIGHT_MM - ROWS * LABEL_HEIGHT) / 2  # ≈ 0.5 mm

# Barcode sizing
BARCODE_MAX_WIDTH_RATIO = 0.88  # max 88% of label width (accounts for padding)
BARCODE_HEIGHT_MM = 20.0  # default barcode height (profiles may override)

# Font settings
FONT_TOP = "Helvetica"
FONT_TOP_SIZE = 9
FONT_BOTTOM = "Helvetica"
FONT_BOTTOM_SIZE = 9

# Profiles -------------------------------------------------------------------
# Each profile defines:
#   top_fields  – list of {"key": logical name, "prefix": optional prefix string}
#   bottom_field – logical name for text below barcode (e.g. QVC)
#   description – human-readable description

PROFILES = {
    "COLLI": {
        "top_fields": [
            {"key": "Testo Superiore 1", "prefix": ""},
            {"key": "Testo Superiore 2", "prefix": "PO: "},
            {"key": "Testo Superiore 3", "prefix": "Quantità: "},
        ],
        "bottom_field": "Testo Inferiore",
        "has_repeat": False,
        "description": "CARTONE + PO + Quantità in alto, Barcode al centro, QVC in basso",
        # Layout overrides for 3-line labels
        "line_spacing_mm": 3.0,
        "font_top_size": 9,
        "barcode_height_mm": 14.5,
        # Preset column mapping
        "default_mapping": {
            "Codice Barcode": "QVC",
            "Testo Superiore 1": "CARTONE",
            "Testo Superiore 2": "PO",
            "Testo Superiore 3": "Quantità",
            "Testo Inferiore": "QVC",
        },
    },
    "SKT": {
        "top_fields": [
            {"key": "Testo Superiore 1", "prefix": ""},
            {"key": "Testo Superiore 2", "prefix": "PO: "},
        ],
        "bottom_field": "Testo Inferiore",
        "has_repeat": True,
        "repeat_field": "Numero Copie",
        "description": "SKT + PO in alto, Barcode QVC al centro, QVC in basso (Qta = n° copie)",
        "line_spacing_mm": 3.0,
        "barcode_height_mm": 17.5,
        # Preset column mapping
        "default_mapping": {
            "Codice Barcode": "Codice QVC",
            "Testo Superiore 1": "SKT",
            "Testo Superiore 2": "Numero PO",
            "Testo Inferiore": "Codice QVC",
            "Numero Copie": "Qta",
        },
    },
}


# ---------------------------------------------------------------------------
# Excel helpers
# ---------------------------------------------------------------------------

def read_excel_headers(path: str) -> list[str]:
    """Return the column headers of the Excel file."""
    df = pd.read_excel(path, nrows=0)
    return list(df.columns)


def read_excel_data(path: str) -> list[dict[str, Any]]:
    """Return all rows as a list of dicts."""
    df = pd.read_excel(path)
    return df.to_dict(orient="records")


# ---------------------------------------------------------------------------
# Config persistence
# ---------------------------------------------------------------------------

def _config_path(directory: str) -> str:
    return os.path.join(directory, CONFIG_FILE)


def load_config(directory: str) -> dict | None:
    """Load saved mapping config from *directory*, or return None."""
    p = _config_path(directory)
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    return None


def save_config(directory: str, cfg: dict) -> None:
    """Persist mapping config to *directory*/config.json."""
    p = _config_path(directory)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


def config_matches(config: dict, headers: list[str]) -> bool:
    """Check if the saved config's mapped columns are still in the headers."""
    mapping = config.get("mapping", {})
    for col in mapping.values():
        if col not in headers:
            return False
    return True


# ---------------------------------------------------------------------------
# Barcode generation
# ---------------------------------------------------------------------------

def generate_barcode_image(code_value: str) -> ImageReader:
    """
    Generate a Code 128 barcode as an in-memory PNG and return a
    ReportLab-compatible ImageReader.  Human-readable text is disabled;
    we print it ourselves for better layout control.
    """
    code_value = str(code_value).strip()
    code128 = barcode.get("code128", code_value, writer=ImageWriter())

    buf = io.BytesIO()
    code128.write(buf, options={
        "write_text": False,       # no built-in text
        "module_height": 15.0,     # mm height of bars
        "module_width": 0.33,      # mm width of narrowest bar
        "quiet_zone": 2.0,         # mm quiet zone
        "dpi": 300,
    })
    buf.seek(0)
    return ImageReader(buf)


def _generate_barcode_svg_data(code_value: str) -> tuple[list[dict], float, float]:
    """
    Generate a Code 128 barcode as SVG and return
    (rect_list, width_mm, height_mm).
    Each rect dict has keys: x, y, width, height (all floats in mm).
    """
    code_value = str(code_value).strip()
    code128 = barcode.get("code128", code_value, writer=SVGWriter())

    buf = io.BytesIO()
    code128.write(buf, options={
        "write_text": False,
        "module_height": 15.0,
        "module_width": 0.33,
        "quiet_zone": 2.0,
    })
    svg_bytes = buf.getvalue()

    SVG_NS = "http://www.w3.org/2000/svg"
    root = ET.fromstring(svg_bytes)

    # Dimensions from width/height attributes (e.g. "85.0mm")
    w_str = root.get("width", "0").replace("mm", "").strip()
    h_str = root.get("height", "0").replace("mm", "").strip()
    bc_width = float(w_str)
    bc_height = float(h_str)

    def _strip_mm(val: str) -> float:
        """Parse an SVG dimension value, stripping optional 'mm' suffix."""
        return float(str(val).replace("mm", "").strip())

    rects: list[dict] = []
    for elem in root.iter(f"{{{SVG_NS}}}rect"):
        # Skip background rects that use percentage dimensions
        w_val = elem.get("width", "0")
        h_val = elem.get("height", "0")
        if "%" in str(w_val) or "%" in str(h_val):
            continue
        try:
            rects.append({
                "x": _strip_mm(elem.get("x", "0")),
                "y": _strip_mm(elem.get("y", "0")),
                "width": _strip_mm(w_val),
                "height": _strip_mm(h_val),
            })
        except ValueError:
            continue

    return rects, bc_width, bc_height


# ---------------------------------------------------------------------------
# Record expansion (for repeat/copies field)
# ---------------------------------------------------------------------------

def _expand_records(
    records: list[dict[str, Any]],
    mapping: dict[str, str],
    profile: str,
) -> list[dict[str, Any]]:
    """
    If the profile has ``has_repeat=True``, duplicate each record N times
    where N is read from the repeat-field column.  Otherwise return records
    unchanged.
    """
    prof = PROFILES[profile]
    if not prof.get("has_repeat", False):
        return records

    repeat_key = prof.get("repeat_field", "")
    repeat_col = mapping.get(repeat_key, "")
    if not repeat_col:
        return records

    expanded: list[dict[str, Any]] = []
    for rec in records:
        try:
            n = int(rec.get(repeat_col, 1))
        except (ValueError, TypeError):
            n = 1
        expanded.extend([rec] * max(n, 1))
    return expanded


# ---------------------------------------------------------------------------
# PDF generation
# ---------------------------------------------------------------------------

def _label_origin(index: int, offset_x: float, offset_y: float) -> tuple[float, float]:
    """
    Return (x, y) in points for the top-left corner of label *index*
    (0-based) on the current page, accounting for offsets.
    """
    col = index % COLS
    row = index // COLS

    x_mm = MARGIN_LEFT + col * LABEL_WIDTH + offset_x
    # Y is measured from the BOTTOM in ReportLab, so we invert.
    y_mm = A4_HEIGHT_MM - MARGIN_TOP - (row + 1) * LABEL_HEIGHT - offset_y

    return x_mm * mm, y_mm * mm


def generate_pdf(
    records: list[dict[str, Any]],
    mapping: dict[str, str],
    profile: str,
    start_pos: int,
    offset_x: float,
    offset_y: float,
    output_path: str,
) -> str:
    """
    Generate a PDF of barcode labels.

    Parameters
    ----------
    records : list of dicts
        Data rows from Excel.
    mapping : dict
        Maps logical field names → Excel column names.
        Required key: "Codice Barcode".
        Optional keys: "Testo Superiore 1/2/3", "Codice QVC".
    profile : str
        "COLLI" or "SKT".
    start_pos : int
        1-based starting label position (1-24).
    offset_x, offset_y : float
        Calibration offsets in mm.
    output_path : str
        Destination PDF file path.

    Returns
    -------
    str  –  The absolute path to the generated PDF.
    """
    prof = PROFILES[profile]
    top_fields = prof["top_fields"]
    bottom_field_key = prof.get("bottom_field", "")

    # Expand records for repeat/copies
    records = _expand_records(records, mapping, profile)

    c = canvas.Canvas(output_path, pagesize=A4)
    c.setTitle("Etichette Barcode")

    label_idx = start_pos - 1  # 0-based position on first page
    page_started = True

    for rec in records:
        # Start a new page when the current one is full
        if label_idx >= LABELS_PER_PAGE:
            c.showPage()
            label_idx = 0
            page_started = True

        x, y = _label_origin(label_idx, offset_x, offset_y)

        # ---- Layout zones (uniform spacing) --------------------------------
        label_top = y + LABEL_HEIGHT * mm
        label_bottom = y
        cx = x + (LABEL_WIDTH * mm) / 2  # horizontal center

        content_top = label_top - LABEL_PAD_Y * mm
        content_bottom = label_bottom + LABEL_PAD_Y * mm

        # Profile overrides
        line_spacing = prof.get("line_spacing_mm", 3.0) * mm
        font_size = prof.get("font_top_size", FONT_TOP_SIZE)
        barcode_h_mm = prof.get("barcode_height_mm", BARCODE_HEIGHT_MM)

        n_top = len(top_fields)
        font_h = font_size * 0.3528 * mm  # pt → mm (approx)
        gap = 1.5 * mm  # uniform gap between zones

        # Zone boundaries (ReportLab: y=0 is page bottom)
        top_zone_bottom = content_top - n_top * line_spacing
        bottom_baseline = content_bottom + 1.0 * mm
        bottom_zone_top = bottom_baseline + font_h

        # Barcode zone: space between top text and bottom text
        barcode_area_top = top_zone_bottom - gap
        barcode_area_bottom = bottom_zone_top + gap
        barcode_available = barcode_area_top - barcode_area_bottom

        # ---- Top text lines ------------------------------------------------
        c.setFont(FONT_TOP, font_size)
        for i, field_def in enumerate(top_fields):
            field_key = field_def["key"]
            prefix = field_def.get("prefix", "")
            col_name = mapping.get(field_key, "")
            if col_name and col_name in rec:
                text = prefix + str(rec[col_name])
            else:
                text = ""
            text_y = content_top - (i + 1) * line_spacing
            c.drawCentredString(cx, text_y, text)

        # ---- Barcode -------------------------------------------------------
        barcode_col = mapping.get("Codice Barcode", "")
        barcode_value = str(rec.get(barcode_col, "")).strip()

        if barcode_value:
            img = generate_barcode_image(barcode_value)

            iw, ih = img.getSize()
            max_w = LABEL_WIDTH * BARCODE_MAX_WIDTH_RATIO * mm
            scale = max_w / iw
            draw_w = iw * scale
            draw_h = ih * scale

            # Clamp to configured max height
            max_h = barcode_h_mm * mm
            if draw_h > max_h:
                scale2 = max_h / draw_h
                draw_w *= scale2
                draw_h = max_h

            # Dynamic clamp: never exceed available space
            if barcode_available > 0 and draw_h > barcode_available:
                clamp_scale = barcode_available / draw_h
                draw_w *= clamp_scale
                draw_h = barcode_available

            # Center barcode in its zone
            barcode_mid = (barcode_area_top + barcode_area_bottom) / 2
            barcode_y = barcode_mid - draw_h / 2
            barcode_x = x + (LABEL_WIDTH * mm - draw_w) / 2
            c.drawImage(img, barcode_x, barcode_y, width=draw_w, height=draw_h,
                        preserveAspectRatio=True, anchor="c")

        # ---- Bottom text (Testo Inferiore) ---------------------------------
        bottom_col = mapping.get(bottom_field_key, "")
        if bottom_col and bottom_col in rec:
            bottom_text = str(rec[bottom_col])
            c.setFont(FONT_BOTTOM, font_size)
            c.drawCentredString(cx, bottom_baseline, bottom_text)

        # ---- Optional: draw light border for debugging (uncomment) ----------
        # c.setStrokeColorRGB(0.85, 0.85, 0.85)
        # c.rect(x, y, LABEL_WIDTH * mm, LABEL_HEIGHT * mm, stroke=1, fill=0)

        label_idx += 1

    c.showPage()
    c.save()
    return os.path.abspath(output_path)


# ---------------------------------------------------------------------------
# SVG generation
# ---------------------------------------------------------------------------

def generate_svg(
    records: list[dict[str, Any]],
    mapping: dict[str, str],
    profile: str,
    start_pos: int,
    offset_x: float,
    offset_y: float,
    output_path: str,
) -> list[str]:
    """
    Generate one SVG file per page of barcode labels (vector, editable
    in Canva / Illustrator / Inkscape).

    Returns a list of absolute paths to the generated SVG files.
    """
    SVG_NS = "http://www.w3.org/2000/svg"
    prof = PROFILES[profile]
    top_fields = prof["top_fields"]
    bottom_field_key = prof.get("bottom_field", "")

    # Expand records for repeat/copies
    records = _expand_records(records, mapping, profile)

    # ---- Split records into pages ----------------------------------------
    pages: list[list[tuple[int, dict]]] = []
    slot = start_pos - 1   # 0-based label slot on current page
    rec_i = 0

    while rec_i < len(records):
        page: list[tuple[int, dict]] = []
        while slot < LABELS_PER_PAGE and rec_i < len(records):
            page.append((slot, records[rec_i]))
            slot += 1
            rec_i += 1
        pages.append(page)
        slot = 0

    if not pages:
        pages = [[]]

    base, _ = os.path.splitext(output_path)
    output_files: list[str] = []

    for page_num, page_data in enumerate(pages):
        ET.register_namespace("", SVG_NS)
        svg = ET.Element("svg", {
            "xmlns": SVG_NS,
            "width": f"{A4_WIDTH_MM}mm",
            "height": f"{A4_HEIGHT_MM}mm",
            "viewBox": f"0 0 {A4_WIDTH_MM} {A4_HEIGHT_MM}",
        })

        # White background
        ET.SubElement(svg, "rect", {
            "width": str(A4_WIDTH_MM),
            "height": str(A4_HEIGHT_MM),
            "fill": "white",
        })

        for label_idx, rec in page_data:
            col = label_idx % COLS
            row = label_idx // COLS

            lx = MARGIN_LEFT + col * LABEL_WIDTH + offset_x
            ly = MARGIN_TOP + row * LABEL_HEIGHT + offset_y

            # ---- Layout zones (uniform spacing) --------------------------
            content_top = ly + LABEL_PAD_Y
            content_bottom = ly + LABEL_HEIGHT - LABEL_PAD_Y
            cx = lx + LABEL_WIDTH / 2

            # Profile overrides
            line_spacing = prof.get("line_spacing_mm", 3.0)
            fs_svg = prof.get("font_top_size", FONT_TOP_SIZE) * 0.3528  # pt → mm

            n_top = len(top_fields)
            font_h = fs_svg
            gap = 1.5  # mm uniform gap between zones

            # Zone boundaries (SVG: y increases downward)
            top_zone_bottom = content_top + n_top * line_spacing
            bottom_baseline = content_bottom - 1.0
            bottom_zone_top = bottom_baseline - font_h

            # Barcode zone
            barcode_area_top = top_zone_bottom + gap
            barcode_area_bottom = bottom_zone_top - gap
            barcode_available = barcode_area_bottom - barcode_area_top

            # ---- Top text lines ------------------------------------------
            for i, field_def in enumerate(top_fields):
                field_key = field_def["key"]
                prefix = field_def.get("prefix", "")
                col_name = mapping.get(field_key, "")
                text_val = prefix + str(rec[col_name]) if (col_name and col_name in rec) else ""

                ty = content_top + (i + 1) * line_spacing
                t = ET.SubElement(svg, "text", {
                    "x": str(cx),
                    "y": str(ty),
                    "text-anchor": "middle",
                    "font-family": "Helvetica, Arial, sans-serif",
                    "font-weight": "bold",
                    "font-size": f"{fs_svg:.2f}",
                    "fill": "black",
                })
                t.text = text_val

            # ---- Barcode -------------------------------------------------
            barcode_col = mapping.get("Codice Barcode", "")
            barcode_value = str(rec.get(barcode_col, "")).strip()

            if barcode_value:
                rects, bc_w, bc_h = _generate_barcode_svg_data(barcode_value)

                max_w = LABEL_WIDTH * BARCODE_MAX_WIDTH_RATIO
                scale = max_w / bc_w if bc_w > 0 else 1.0

                scaled_h = bc_h * scale
                max_h = prof.get("barcode_height_mm", BARCODE_HEIGHT_MM)
                if scaled_h > max_h:
                    scale *= max_h / scaled_h
                    scaled_h = max_h

                scaled_w = bc_w * scale

                # Dynamic clamp: never exceed available space
                if barcode_available > 0 and scaled_h > barcode_available:
                    clamp_scale = barcode_available / scaled_h
                    scale *= clamp_scale
                    scaled_w = bc_w * scale
                    scaled_h = barcode_available

                # Center barcode in its zone
                barcode_mid = (barcode_area_top + barcode_area_bottom) / 2
                bc_x = lx + (LABEL_WIDTH - scaled_w) / 2
                bc_y = barcode_mid - scaled_h / 2

                g = ET.SubElement(svg, "g", {
                    "transform": f"translate({bc_x:.3f},{bc_y:.3f}) scale({scale:.6f})",
                })
                for r in rects:
                    ET.SubElement(g, "rect", {
                        "x": str(r["x"]),
                        "y": str(r["y"]),
                        "width": str(r["width"]),
                        "height": str(r["height"]),
                        "fill": "black",
                    })

            # ---- Bottom text (Testo Inferiore) ---------------------------
            bottom_col = mapping.get(bottom_field_key, "")
            if bottom_col and bottom_col in rec:
                bottom_text = str(rec[bottom_col])
                t = ET.SubElement(svg, "text", {
                    "x": str(cx),
                    "y": str(bottom_baseline),
                    "text-anchor": "middle",
                    "font-family": "Helvetica, Arial, sans-serif",
                    "font-weight": "bold",
                    "font-size": f"{fs_svg:.2f}",
                    "fill": "black",
                })
                t.text = bottom_text

        # Write SVG file
        if len(pages) == 1:
            svg_path = f"{base}.svg"
        else:
            svg_path = f"{base}_pagina{page_num + 1}.svg"

        tree = ET.ElementTree(svg)
        ET.indent(tree, space="  ")
        with open(svg_path, "w", encoding="utf-8") as f:
            f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
            tree.write(f, encoding="unicode", xml_declaration=False)

        output_files.append(os.path.abspath(svg_path))

    return output_files
