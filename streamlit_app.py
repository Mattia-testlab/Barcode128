import streamlit as st
import os
import pandas as pd
from label_generator import (
    read_excel_headers,
    read_excel_data,
    generate_pdf,
    get_dummy_records,
    PROFILES,
    LABELS_PER_PAGE,
)
import tempfile
import base64

def main():
    st.set_page_config(page_title="Barcode Pro Web", page_icon="🏷️", layout="wide")

    st.title("🏷️ Barcode Pro Web")
    st.markdown("Generatore di etichette Code 128 per logistica e magazzino")

    # Sidebar for configuration
    st.sidebar.header("1. Carica File")
    uploaded_file = st.sidebar.file_uploader("Scegli un file Excel", type=["xlsx", "xls"])

    if uploaded_file is not None:
        # Save uploaded file to a temporary location to use with existing functions
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        try:
            headers = read_excel_headers(tmp_path)
            st.success(f"File caricato con successo: {len(headers)} colonne trovate.")

            col1, col2 = st.columns(2)

            with col1:
                st.header("2. Mapping Colonne")
                mapping = {}
                field_names = [
                    ("Codice Barcode", True),
                    ("Testo Superiore 1", False),
                    ("Testo Superiore 2", False),
                    ("Testo Superiore 3", False),
                    ("Testo Inferiore", False),
                    ("Numero Copie", False),
                ]

                options = ["(nessuna)"] + headers
                
                # Check for profile-based auto-mapping
                selected_profile = st.selectbox("Scegli profilo per auto-mapping", list(PROFILES.keys()))
                
                defaults = PROFILES[selected_profile].get("default_mapping", {})

                for field, required in field_names:
                    label = f"{field} {'*' if required else '(opzionale)'}"
                    default_index = 0
                    if field in defaults and defaults[field] in headers:
                        default_index = headers.index(defaults[field]) + 1
                    
                    val = st.selectbox(label, options, index=default_index, key=field)
                    if val != "(nessuna)":
                        mapping[field] = val

            with col2:
                st.header("3. Impostazioni e Stampa")
                profile = st.radio("Profilo di stampa finale", list(PROFILES.keys()), index=list(PROFILES.keys()).index(selected_profile))
                
                # ── Field info box ──────────────────────────────────────
                prof_data = PROFILES[profile]
                field_info = prof_data.get("field_info", {})
                
                if profile == "COLLI":
                    st.markdown("""
                    <div style="background: linear-gradient(135deg, #e8f4f8, #d4e8f0); border-radius: 12px; padding: 16px; margin: 8px 0; border-left: 4px solid #2196F3;">
                        <div style="font-weight: 700; font-size: 15px; color: #1565C0; margin-bottom: 10px;">📦 Profilo COLLI — Layout Etichetta</div>
                        <p style="font-size: 14px; color: #333; margin: 0; line-height: 1.6;">
                            <b>1.</b> In alto: <b>CARTONE</b><br>
                            <b>2.</b> Sotto: <b>PO:</b> + {codice PO}<br>
                            <b>3.</b> Sotto: <b>Quantità:</b> + {quantità}<br>
                            <b>4.</b> Al centro: <b>Barcode</b> ({Codice QVC})
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
                else:  # SKT
                    st.markdown("""
                    <div style="background: linear-gradient(135deg, #f3e8f8, #e4d4f0); border-radius: 12px; padding: 16px; margin: 8px 0; border-left: 4px solid #9C27B0;">
                        <div style="font-weight: 700; font-size: 15px; color: #7B1FA2; margin-bottom: 10px;">🏷️ Profilo SKT — Layout Etichetta</div>
                        <p style="font-size: 14px; color: #333; margin: 0; line-height: 1.6;">
                            <b>1.</b> In alto: <b>Codice SKT</b><br>
                            <b>2.</b> Sotto: <b>PO</b> + {codice PO}<br>
                            <b>3.</b> Al centro: <b>Barcode</b> ({Codice QVC})<br>
                            <b>4.</b> In basso: <b>Codice QVC</b>
                        </p>
                    </div>
                    """, unsafe_allow_html=True)

                start_pos = st.number_input(f"Posizione di inizio (1-{LABELS_PER_PAGE})", min_value=1, max_value=LABELS_PER_PAGE, value=1)
                
                off_x = st.number_input("Offset X (mm)", value=0.0, step=0.1)
                off_y = st.number_input("Offset Y (mm)", value=0.0, step=0.1)

                with st.expander("🎛️ Regolazioni Layout"):
                    lo_pad = st.slider("Margine interno (pad Y, mm) [Standard: 4.5]", 0.5, 6.0, 4.5, 0.5,
                                       help="Distanza dal bordo superiore/inferiore dell'etichetta (min 4.5 consigliato per i bordi di stampa)")
                    lo_gap = st.slider("Spazio testo-barcode (mm) [Standard: 0.5]", 0.1, 4.0, 0.5, 0.1,
                                       help="Distanza tra le righe di testo e il barcode")
                    lo_ls = st.slider("Spaziatura righe testo (mm) [Standard: 2.2]", 1.5, 5.0, 2.2, 0.1,
                                      help="Distanza tra le righe di testo superiore")
                    lo_fs = st.slider("Dimensione font (pt) [Standard: 7]", 5, 14, 7, 1,
                                      help="Dimensione del testo superiore e inferiore")
                    lo_bh = st.slider("Altezza barcode (mm) [Standard: 19.0]", 10.0, 35.0, 19.0, 0.5,
                                      help="Altezza del barcode (il barcode si restringe automaticamente se lo spazio verticale non è sufficiente)")



                layout_overrides = {
                    "pad_y_mm": lo_pad,
                    "gap_mm": lo_gap,
                    "line_spacing_mm": lo_ls,
                    "font_size_pt": lo_fs,
                    "barcode_height_mm": lo_bh,
                }

                # ── Generazione e Download ─────────────────────────────────
                missing_required = [f for f, req in field_names if req and f not in mapping]
                
                if missing_required:
                    st.warning(f"⚠️ Seleziona i campi obbligatori per abilitare il download: {', '.join(missing_required)}")
                else:
                    try:
                        records = read_excel_data(tmp_path)
                        
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as out_tmp:
                            output_path = out_tmp.name

                        result_path = generate_pdf(
                            records=records,
                            mapping=mapping,
                            profile=profile,
                            start_pos=start_pos,
                            offset_x=off_x,
                            offset_y=off_y,
                            output_path=output_path,
                            layout_overrides=layout_overrides,
                        )

                        with open(result_path, "rb") as f:
                            pdf_data = f.read()

                        st.success("✅ File PDF pronto! (Il file si aggiorna automaticamente se modifichi gli slider sopra)")
                        st.download_button(
                            label="📥 Scarica PDF Aggiornato",
                            data=pdf_data,
                            file_name=f"{os.path.splitext(uploaded_file.name)[0]}_etichette_{profile}.pdf",
                            mime="application/pdf"
                        )
                        
                    except Exception as e:
                        st.error(f"Errore durante la generazione del PDF: {e}")
        
        finally:
            # Note: We should ideally cleanup the temp file, but NamedTemporaryFile(delete=False) 
            # is used because generate_pdf might need it. 
            pass

    else:
        st.info("Carica un file Excel dalla barra laterale per iniziare.")

if __name__ == "__main__":
    main()
