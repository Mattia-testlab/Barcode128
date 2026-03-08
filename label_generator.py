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
from reportlab.pdfbase import pdfmetrics
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
LABEL_HEIGHT = 37.125  # 297/8 exactly (must match physical grid)

# Padding within each label cell (keeps content off pre-cut borders)
LABEL_PAD_X = 2.0   # mm inset from left/right edges
LABEL_PAD_Y = 2.0   # mm inset from top/bottom edges (safe print margin)

# Margins to center the grid on the page
MARGIN_LEFT = (A4_WIDTH_MM - COLS * LABEL_WIDTH) / 2  # ≈ 0 mm
MARGIN_TOP = (A4_HEIGHT_MM - ROWS * LABEL_HEIGHT) / 2  # ≈ 0.5 mm

# Barcode sizing
BARCODE_MAX_WIDTH_RATIO = 0.95  # max 95% of label width (slightly wider, still safe on pre-cut labels)
BARCODE_HEIGHT_MM = 20.0  # default barcode height (profiles may override)
BOTTOM_FONT_SCALE = 1.30

# Font settings
FONT_TOP = "Helvetica-Bold"
FONT_TOP_SIZE = 11
FONT_BOTTOM = "Helvetica-Bold"
FONT_BOTTOM_SIZE = 11

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
        "line_spacing_mm": 3.6,
        "font_top_size": 11,
        "barcode_height_mm": 25.0,
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
        "line_spacing_mm": 3.6,
        "font_top_size": 11,
        "barcode_height_mm": 25.0,
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


def _cell_to_text(value: Any) -> str:
    """
    Convert an Excel cell value to printable text.
    Returns an empty string for missing values (None/NaN).
    """
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass
    return str(value).strip()


def _fit_font_size_pt(
    text: str,
    font_name: str,
    desired_size_pt: float,
    max_width_pt: float,
    min_size_pt: float = 6.0,
) -> float:
    """
    Reduce font size only when needed to keep text within max width.
    """
    if not text:
        return desired_size_pt

    text_width = pdfmetrics.stringWidth(text, font_name, desired_size_pt)
    if text_width <= max_width_pt or text_width <= 0:
        return desired_size_pt

    fitted = desired_size_pt * (max_width_pt / text_width)
    return max(min_size_pt, min(desired_size_pt, fitted))


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
        "module_height": 30.0,     # mm height of bars (tall for ≥25mm on label)
        "module_width": 0.38,      # mm width of narrowest bar (wider for reliability)
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
        "module_height": 30.0,
        "module_width": 0.38,
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
    layout_overrides: dict | None = None,
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
    layout_overrides : dict, optional
        Override layout params: pad_y_mm, gap_mm, line_spacing_mm,
        font_size_pt, barcode_height_mm.

    Returns
    -------
    str  –  The absolute path to the generated PDF.
    """
    lo = layout_overrides or {}
    prof = PROFILES[profile]
    top_fields = prof["top_fields"]
    bottom_field_key = prof.get("bottom_field", "")

    # Expand records for repeat/copies
    records = _expand_records(records, mapping, profile)

    c = canvas.Canvas(output_path, pagesize=A4)
    c.setTitle("Etichette Barcode")

    label_idx = start_pos - 1  # 0-based position on first page
    page_started = True

    # Resolve layout values: override > profile > global default
    pad_y = lo.get("pad_y_mm", LABEL_PAD_Y)
    gap_val = lo.get("gap_mm", 1.5)
    ls_mm = lo.get("line_spacing_mm", prof.get("line_spacing_mm", 3.6))
    font_size = lo.get("font_size_pt", prof.get("font_top_size", FONT_TOP_SIZE))
    barcode_h_mm = lo.get("barcode_height_mm", prof.get("barcode_height_mm", BARCODE_HEIGHT_MM))

    for rec in records:
        # Start a new page when the current one is full
        if label_idx >= LABELS_PER_PAGE:
            c.showPage()
            label_idx = 0
            page_started = True

        x, y = _label_origin(label_idx, offset_x, offset_y)

        # ---- Layout zones (exact bounding box math) ------------------------
        label_top = y + LABEL_HEIGHT * mm
        label_bottom = y
        cx = x + (LABEL_WIDTH * mm) / 2

        content_top = label_top - pad_y * mm
        content_bottom = label_bottom + pad_y * mm

        n_top = len(top_fields)
        font_h = font_size * 0.3528 * mm
        cap_h = font_h * 0.7
        gap = gap_val * mm

        # Top text bounding box (ReportLab Y goes up)
        y_first_line = content_top - cap_h
        y_last_line = y_first_line - (n_top - 1) * (ls_mm * mm)
        top_zone_bottom = y_last_line - font_h * 0.2

        # Bottom text bounding box
        bottom_font_size = font_size * BOTTOM_FONT_SCALE
        bottom_font_h = bottom_font_size * 0.3528 * mm
        bottom_baseline = content_bottom + bottom_font_h * 0.2
        bottom_zone_top = bottom_baseline + bottom_font_h * 0.7
        max_text_width = (LABEL_WIDTH - 2 * LABEL_PAD_X) * mm

        # Barcode area exactly between top text and bottom text
        barcode_area_top = top_zone_bottom - gap
        barcode_area_bottom = bottom_zone_top + gap
        barcode_available = barcode_area_top - barcode_area_bottom

        # ---- Top text lines ------------------------------------------------
        c.setFont(FONT_TOP, font_size)
        for i, field_def in enumerate(top_fields):
            field_key = field_def["key"]
            prefix = field_def.get("prefix", "")
            col_name = mapping.get(field_key, "")
            text_value = _cell_to_text(rec.get(col_name)) if col_name else ""
            text = f"{prefix}{text_value}" if text_value else ""
            text_y = y_first_line - i * (ls_mm * mm)
            top_font_size = _fit_font_size_pt(text, FONT_TOP, font_size, max_text_width)
            c.setFont(FONT_TOP, top_font_size)
            c.drawCentredString(cx, text_y, text)

        # ---- Barcode -------------------------------------------------------
        barcode_col = mapping.get("Codice Barcode", "")
        barcode_value = _cell_to_text(rec.get(barcode_col))

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
        bottom_text = _cell_to_text(rec.get(bottom_col)) if bottom_col else ""
        if bottom_text:
            bottom_draw_size = _fit_font_size_pt(
                bottom_text,
                FONT_BOTTOM,
                bottom_font_size,
                max_text_width,
                min_size_pt=6.5,
            )
            c.setFont(FONT_BOTTOM, bottom_draw_size)
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
    layout_overrides: dict | None = None,
) -> list[str]:
    """
    Generate one SVG file per page of barcode labels (vector, editable
    in Canva / Illustrator / Inkscape).

    Returns a list of absolute paths to the generated SVG files.
    """
    lo = layout_overrides or {}
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

            # ---- Layout zones (exact bounding box math) ------------------------
            # Resolve layout values: override > profile > global default
            pad_y = lo.get("pad_y_mm", LABEL_PAD_Y)
            gap_val = lo.get("gap_mm", 1.5)
            ls_mm = lo.get("line_spacing_mm", prof.get("line_spacing_mm", 3.6))
            fs_pt = lo.get("font_size_pt", prof.get("font_top_size", FONT_TOP_SIZE))
            bc_h_mm = lo.get("barcode_height_mm", prof.get("barcode_height_mm", BARCODE_HEIGHT_MM))

            # SVG coordinates (Y=0 is top, Y increases downwards)
            content_top = ly + pad_y
            content_bottom = ly + LABEL_HEIGHT - pad_y
            cx = lx + LABEL_WIDTH / 2

            n_top = len(top_fields)
            font_h = fs_pt * 0.3528
            cap_h = font_h * 0.7
            gap = gap_val

            # Top text bounding box (Y goes down)
            y_first_line = content_top + cap_h
            y_last_line = y_first_line + (n_top - 1) * ls_mm
            top_zone_bottom = y_last_line + font_h * 0.2
            
            # Bottom text bounding box
            bottom_font_size = fs_pt * BOTTOM_FONT_SCALE
            bottom_font_h = bottom_font_size * 0.3528
            bottom_baseline = content_bottom - bottom_font_h * 0.2
            bottom_zone_top = bottom_baseline - bottom_font_h * 0.7
            
            # Barcode area (SVG: top is smaller Y, bottom is larger Y)
            barcode_area_top = top_zone_bottom + gap
            barcode_area_bottom = bottom_zone_top - gap
            barcode_available = barcode_area_bottom - barcode_area_top

            # ---- Top text lines ------------------------------------------
            for i, field_def in enumerate(top_fields):
                field_key = field_def["key"]
                prefix = field_def.get("prefix", "")
                col_name = mapping.get(field_key, "")
                text_value = _cell_to_text(rec.get(col_name)) if col_name else ""
                text_val = f"{prefix}{text_value}" if text_value else ""

                ty = y_first_line + i * ls_mm
                t = ET.SubElement(svg, "text", {
                    "x": f"{cx:.3f}",
                    "y": f"{ty:.3f}",
                    "text-anchor": "middle",
                    "font-family": "Helvetica, Arial, sans-serif",
                    "font-weight": "bold",
                    "font-size": f"{fs_pt * 0.3528:.2f}",
                    "fill": "black",
                })
                t.text = text_val

            # ---- Barcode -------------------------------------------------
            barcode_col = mapping.get("Codice Barcode", "")
            barcode_value = _cell_to_text(rec.get(barcode_col))

            if barcode_value:
                rects, bc_w, bc_h = _generate_barcode_svg_data(barcode_value)

                max_w = LABEL_WIDTH * BARCODE_MAX_WIDTH_RATIO
                scale = max_w / bc_w if bc_w > 0 else 1.0

                scaled_h = bc_h * scale
                max_h = bc_h_mm
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
            bottom_text = _cell_to_text(rec.get(bottom_col)) if bottom_col else ""
            if bottom_text:
                t = ET.SubElement(svg, "text", {
                    "x": str(cx),
                    "y": str(bottom_baseline),
                    "text-anchor": "middle",
                    "font-family": "Helvetica, Arial, sans-serif",
                    "font-weight": "bold",
                    "font-size": f"{bottom_font_size * 0.3528:.2f}",
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


# ---------------------------------------------------------------------------
# Live preview (PIL image for Streamlit / GUI)
# ---------------------------------------------------------------------------

def generate_preview_image(
    n_top_lines: int = 3,
    layout_overrides: dict | None = None,
    scale_factor: float = 6.0,
) -> "Image.Image":
    """
    Generate a single-label preview as a PIL Image showing the layout zones.

    Parameters
    ----------
    n_top_lines : int
        Number of top text lines (3 for COLLI, 2 for SKT).
    layout_overrides : dict, optional
        Same keys as generate_pdf: pad_y_mm, gap_mm, line_spacing_mm,
        font_size_pt, barcode_height_mm.
    scale_factor : float
        Pixels per mm for rendering (default 6 → 420×222 px).

    Returns
    -------
    PIL.Image.Image
    """
    from PIL import Image, ImageDraw, ImageFont

    lo = layout_overrides or {}
    pad_y    = lo.get("pad_y_mm", LABEL_PAD_Y)
    gap_val  = lo.get("gap_mm", 1.5)
    ls       = lo.get("line_spacing_mm", 3.0)
    fs_pt    = lo.get("font_size_pt", FONT_TOP_SIZE)
    bc_h_cfg = lo.get("barcode_height_mm", 14.5)

    sf = scale_factor
    W = int(LABEL_WIDTH * sf)
    H = int(LABEL_HEIGHT * sf)

    # Colours
    BG_COL      = (255, 255, 255)
    BORDER_COL  = (180, 180, 180)
    PAD_COL     = (235, 245, 255)   # light blue for padding zones
    TEXT_COL    = (220, 235, 255)   # slightly darker blue for text zones
    GAP_COL     = (255, 245, 220)   # warm yellow for gap
    BC_COL      = (230, 255, 230)   # green for barcode zone
    ANNOT_COL   = (100, 100, 100)
    LINE_COL    = (60, 60, 60)
    LABEL_TXT   = (40, 40, 40)

    img = Image.new("RGB", (W + 120, H), BG_COL)  # extra space for annotations on right
    draw = ImageDraw.Draw(img)

    # Try to get a nice font, fall back to default
    try:
        font_sm = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", int(9 * sf / 6))
        font_xs = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", int(7 * sf / 6))
    except Exception:
        font_sm = ImageFont.load_default()
        font_xs = font_sm

    # Zone calculations (in mm, then convert to px)
    content_top    = pad_y
    content_bottom = LABEL_HEIGHT - pad_y
    content_h      = content_bottom - content_top
    font_h_mm      = fs_pt * 0.3528

    top_text_h     = n_top_lines * ls
    top_zone_bottom = content_top + top_text_h

    bottom_baseline_off = 1.0
    bottom_zone_h  = font_h_mm + bottom_baseline_off
    bottom_zone_top = content_bottom - bottom_zone_h

    barcode_area_top    = top_zone_bottom + gap_val
    barcode_area_bottom = bottom_zone_top - gap_val
    barcode_available   = barcode_area_bottom - barcode_area_top
    barcode_h_actual    = min(bc_h_cfg, max(barcode_available, 0))

    # Helper: mm -> px (from top of label)
    def ypx(mm_from_top):
        return int(mm_from_top * sf)

    def xpx(mm_from_left):
        return int(mm_from_left * sf)

    lbl_w = int(LABEL_WIDTH * sf)

    # ── Draw zones ──────────────────────────────────────────────────

    # Padding top
    draw.rectangle([0, 0, lbl_w, ypx(pad_y)], fill=PAD_COL)
    # Padding bottom
    draw.rectangle([0, ypx(LABEL_HEIGHT - pad_y), lbl_w, H], fill=PAD_COL)

    # Top text zone
    draw.rectangle([0, ypx(content_top), lbl_w, ypx(top_zone_bottom)], fill=TEXT_COL)

    # Gap top
    draw.rectangle([0, ypx(top_zone_bottom), lbl_w, ypx(barcode_area_top)], fill=GAP_COL)

    # Barcode zone
    if barcode_available > 0:
        # Center barcode in available space
        bc_mid = (barcode_area_top + barcode_area_bottom) / 2
        bc_top = bc_mid - barcode_h_actual / 2
        bc_bot = bc_mid + barcode_h_actual / 2
        draw.rectangle([xpx(4), ypx(bc_top), lbl_w - xpx(4), ypx(bc_bot)], fill=BC_COL)
        # Draw fake barcode lines
        for i in range(20):
            bx = xpx(10) + i * int((lbl_w - xpx(20)) / 20)
            bar_w = 2 if i % 3 else 3
            draw.rectangle([bx, ypx(bc_top) + 4, bx + bar_w, ypx(bc_bot) - 4], fill=(0, 0, 0))

    # Gap bottom
    draw.rectangle([0, ypx(barcode_area_bottom), lbl_w, ypx(bottom_zone_top)], fill=GAP_COL)

    # Bottom text zone
    draw.rectangle([0, ypx(bottom_zone_top), lbl_w, ypx(content_bottom)], fill=TEXT_COL)

    # ── Draw sample text ────────────────────────────────────────────

    sample_top = ["2004381848", "PO: 10", "Quantità: 4559057090"]
    for i in range(min(n_top_lines, len(sample_top))):
        ty = ypx(content_top + (i + 0.5) * ls) - 4
        txt = sample_top[i]
        bbox_w = draw.textlength(txt, font=font_sm)
        draw.text((lbl_w // 2 - int(bbox_w) // 2, ty), txt, fill=LABEL_TXT, font=font_sm)

    # Bottom text
    btxt = "155002 AAP 035"
    bbox_w = draw.textlength(btxt, font=font_sm)
    bty = ypx(bottom_zone_top + bottom_zone_h / 2) - 5
    draw.text((lbl_w // 2 - int(bbox_w) // 2, bty), btxt, fill=LABEL_TXT, font=font_sm)

    # ── Label border (dashed) ───────────────────────────────────────

    for i in range(0, lbl_w, 8):
        draw.line([(i, 0), (min(i + 4, lbl_w), 0)], fill=BORDER_COL, width=2)
        draw.line([(i, H - 1), (min(i + 4, lbl_w), H - 1)], fill=BORDER_COL, width=2)
    for i in range(0, H, 8):
        draw.line([(0, i), (0, min(i + 4, H))], fill=BORDER_COL, width=2)
        draw.line([(lbl_w - 1, i), (lbl_w - 1, min(i + 4, H))], fill=BORDER_COL, width=2)

    # ── Annotations on the right ────────────────────────────────────

    ax = lbl_w + 8  # annotation x start

    def annot(y1_mm, y2_mm, label, color=ANNOT_COL):
        y1 = ypx(y1_mm)
        y2 = ypx(y2_mm)
        mid = (y1 + y2) // 2
        draw.line([(ax, y1), (ax + 12, y1)], fill=color, width=1)
        draw.line([(ax, y2), (ax + 12, y2)], fill=color, width=1)
        draw.line([(ax + 6, y1), (ax + 6, y2)], fill=color, width=1)
        draw.text((ax + 16, mid - 5), label, fill=color, font=font_xs)

    annot(0, pad_y, f"pad {pad_y:.1f}")
    annot(content_top, top_zone_bottom, f"testo {top_text_h:.1f}")
    annot(top_zone_bottom, barcode_area_top, f"gap {gap_val:.1f}")
    if barcode_available > 0:
        annot(barcode_area_top, barcode_area_bottom, f"barcode {barcode_h_actual:.1f}")
    annot(barcode_area_bottom, bottom_zone_top, f"gap {gap_val:.1f}")
    annot(bottom_zone_top, content_bottom, f"testo {bottom_zone_h:.1f}")
    annot(content_bottom, LABEL_HEIGHT, f"pad {pad_y:.1f}")

    # Overflow warning
    if barcode_available <= 0:
        draw.text((10, H // 2 - 10), "⚠ OVERFLOW!", fill=(255, 0, 0), font=font_sm)

    return img
