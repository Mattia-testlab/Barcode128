import os
import io
import tempfile
import pandas as pd
import barcode
from barcode.writer import ImageWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from PIL import Image

# ==========================================
# PAGE & LABEL CONSTANTS
# ==========================================
PAGE_WIDTH, PAGE_HEIGHT = A4  # 210mm x 297mm
COLS = 3
ROWS = 8

LABEL_WIDTH = 70.0 * mm
LABEL_HEIGHT = 37.125 * mm  # exactly 297 / 8

# ==========================================
# DEFAULT LAYOUT PARAMETERS
# ==========================================
DEFAULT_LAYOUT = {
    # Spacing and margins
    "margin_y": 4.5 * mm,         # Top and bottom inner margin
    "margin_x": 2.0 * mm,         # Left and right inner margin
    "text_barcode_spacing": 0.5 * mm,
    # Barcode
    "barcode_max_width_pct": 0.92, # Barcode max width relative to label width
    "barcode_height": 19.0 * mm,   # Default height
    # Profile "COLLI" overrides
    "colli_font_size": 7,
    "colli_line_spacing": 2.2 * mm,
    # Profile "SKT" overrides
    "skt_font_size": 9,
    "skt_line_spacing": 2.5 * mm,
}

def read_excel_headers(file_path):
    """
    Reads only the headers of the uploaded Excel file.
    Returns a list of column names.
    """
    df = pd.read_excel(file_path, nrows=0)
    return list(df.columns)

def generate_barcode_image(data):
    """
    Generates a Code 128 barcode image (PIL Image) in memory.
    """
    code128 = barcode.get("code128", data, writer=ImageWriter())
    # Generate barcode without text
    options = {
        "write_text": False,
        "quiet_zone": 1.0,
        "module_width": 0.2, # thin lines for good resolution
        "module_height": 15.0,
        "dpi": 300
    }
    img_byte_arr = io.BytesIO()
    code128.write(img_byte_arr, options=options)
    img_byte_arr.seek(0)
    return Image.open(img_byte_arr)

def calculate_label_position(index, offset_x_mm=0.0, offset_y_mm=0.0):
    """
    Calculates the exact (x, y) coordinates of the bottom-left corner of a label
    based on its index (0 to 23), taking into account global hardware offsets.
    The grid starts from Top-Left, reading left-to-right, top-to-bottom.
    """
    col = index % COLS
    row = index // COLS
    
    # x is from left
    x = (col * LABEL_WIDTH) + (offset_x_mm * mm)
    
    # y is from bottom (ReportLab coordinates)
    # Row 0 is the top row, so its Y coordinate is top of page minus one label height.
    y = PAGE_HEIGHT - ((row + 1) * LABEL_HEIGHT) + (offset_y_mm * mm)
    
    return x, y

def process_data(df, profile, mapping):
    """
    Processes the DataFrame according to the selected profile and mapping.
    Handles 'N copies' for SKT.
    Returns a list of dictionaries with extracted string components.
    """
    labels_data = []
    
    for _, row in df.iterrows():
        # Get Barcode Data
        b_col = mapping.get("Codice Barcode")
        if not b_col or b_col == "(nessuna)" or pd.isna(row.get(b_col)):
            continue
        barcode_value = str(row[b_col]).strip()
        
        # Determine number of copies (Default 1)
        copies = 1
        if profile == "SKT":
            q_col = mapping.get("Numero Copie")
            if q_col and q_col != "(nessuna)" and not pd.isna(row.get(q_col)):
                try:
                    copies = int(row[q_col])
                except ValueError:
                    copies = 1
                    
        # Extract Texts
        testo_1_col = mapping.get("Testo Superiore 1")
        testo_2_col = mapping.get("Testo Superiore 2")
        testo_3_col = mapping.get("Testo Superiore 3")
        
        testo_1 = str(row[testo_1_col]).strip() if testo_1_col and testo_1_col != "(nessuna)" and not pd.isna(row.get(testo_1_col)) else ""
        testo_2 = str(row[testo_2_col]).strip() if testo_2_col and testo_2_col != "(nessuna)" and not pd.isna(row.get(testo_2_col)) else ""
        testo_3 = str(row[testo_3_col]).strip() if testo_3_col and testo_3_col != "(nessuna)" and not pd.isna(row.get(testo_3_col)) else ""
        
        # Format Text blocks based on Profile
        text_lines = []
        if profile == "COLLI":
            # Riga 1: {Numero Cartone}
            if testo_1: text_lines.append(testo_1)
            # Riga 2: PO: {Numero PO}
            if testo_2: text_lines.append(f"PO: {testo_2}")
            # Riga 3: Quantità: {Quantità}
            if testo_3: text_lines.append(f"Quantità: {testo_3}")
                
        elif profile == "SKT":
            # Riga 1: {Codice SKT}
            if testo_1: text_lines.append(testo_1)
            # Riga 2: PO: {Numero PO}
            if testo_2: text_lines.append(f"PO: {testo_2}")
            
        data_packet = {
            "barcode": barcode_value,
            "texts": text_lines
        }
        
        for _ in range(copies):
            labels_data.append(data_packet)
            
    return labels_data

def generate_pdf(df, profile, mapping, start_position=1, offset_x=0.0, offset_y=0.0, layout_overrides=None):
    """
    Main entry point for generating the final PDF.
    Start position is 1-indexed (1 to 24).
    """
    if layout_overrides is None:
        layout_overrides = {}
        
    layout = DEFAULT_LAYOUT.copy()
    
    # Merge overrides (they come in mm from UI, so multiply by mm if needed, but we assume UI sends exact float values in mm/pt)
    if "margin_y" in layout_overrides: layout["margin_y"] = layout_overrides["margin_y"] * mm
    if "margin_x" in layout_overrides: layout["margin_x"] = layout_overrides["margin_x"] * mm
    if "text_barcode_spacing" in layout_overrides: layout["text_barcode_spacing"] = layout_overrides["text_barcode_spacing"] * mm
    if "barcode_height" in layout_overrides: layout["barcode_height"] = layout_overrides["barcode_height"] * mm
    if "line_spacing" in layout_overrides:
        layout["colli_line_spacing"] = layout_overrides["line_spacing"] * mm
        layout["skt_line_spacing"] = layout_overrides["line_spacing"] * mm
    if "font_size" in layout_overrides:
        layout["colli_font_size"] = layout_overrides["font_size"]
        layout["skt_font_size"] = layout_overrides["font_size"]

    # Profile specific setup
    font_name = "Helvetica-Bold"
    if profile == "COLLI":
        font_size = layout["colli_font_size"]
        line_spacing = layout["colli_line_spacing"]
    else:  # SKT
        font_size = layout["skt_font_size"]
        line_spacing = layout["skt_line_spacing"]

    # Validate start position
    current_idx = max(0, start_position - 1)
    
    # Process data
    labels_to_print = process_data(df, profile, mapping)
    
    # Initialization
    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    c = canvas.Canvas(temp_pdf.name, pagesize=A4)
    
    for lbl in labels_to_print:
        # Check if we need a new page
        if current_idx >= 24:
            c.showPage()
            current_idx = 0
            
        # Get bottom-left corner of the label
        lb_x, lb_y = calculate_label_position(current_idx, offset_x, offset_y)
        
        # Define the drawable area inside the label padding
        draw_x_min = lb_x + layout["margin_x"]
        draw_x_max = lb_x + LABEL_WIDTH - layout["margin_x"]
        draw_width = draw_x_max - draw_x_min
        
        draw_y_top = lb_y + LABEL_HEIGHT - layout["margin_y"]
        draw_y_bottom = lb_y + layout["margin_y"]
        draw_height = draw_y_top - draw_y_bottom
        
        # 1. DRAW TEXT (Top-down)
        c.setFont(font_name, font_size)
        
        current_y = draw_y_top
        for text_line in lbl["texts"]:
            # Move down by font size to baseline
            current_y -= (font_size * 0.35)  # approx baseline adjustment
            # X Centering
            text_width = c.stringWidth(text_line, font_name, font_size)
            text_x = draw_x_min + (draw_width - text_width) / 2.0
            
            c.drawString(text_x, current_y, text_line)
            # Move down for next line
            current_y -= line_spacing
            
        # 2. DRAW BARCODE
        # Calculate remaining space
        # Account for text-barcode spacing
        barcode_top_y = current_y - layout["text_barcode_spacing"]
        remaining_height = barcode_top_y - draw_y_bottom
        
        if remaining_height > 0:
            barcode_img = generate_barcode_image(lbl["barcode"])
            img_w, img_h = barcode_img.size
            
            # Constrain dimensions
            bc_target_w = draw_width * layout["barcode_max_width_pct"]
            bc_target_h = min(layout["barcode_height"], remaining_height)
            
            # Barcode aspect ratio check - usually we want to stretch Code128 properly
            # but we just fit it in the space.
            # python-barcode ImageWriter creates an image with margins, we crop them conceptually or just resize.
            
            # Center vertically in the remaining space
            bc_y = draw_y_bottom + (remaining_height - bc_target_h) / 2.0
            
            # Center horizontally
            bc_x = draw_x_min + (draw_width - bc_target_w) / 2.0
            
            # Save tmp image for reportlab
            tmp_img_path = tempfile.mktemp(suffix=".png")
            barcode_img.save(tmp_img_path)
            
            c.drawImage(tmp_img_path, bc_x, bc_y, width=bc_target_w, height=bc_target_h)
            os.remove(tmp_img_path)
            
        current_idx += 1
        
    c.save()
    return temp_pdf.name
