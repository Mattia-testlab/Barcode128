import streamlit as st
import os
import pandas as pd
from label_generator import (
    read_excel_headers,
    read_excel_data,
    generate_pdf,
    PROFILES,
    LABELS_PER_PAGE,
)
import tempfile

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
                auto_map = st.button("🪄 Applica Auto-mapping")
                
                defaults = {}
                if auto_map:
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
                
                start_pos = st.number_input(f"Posizione di inizio (1-{LABELS_PER_PAGE})", min_value=1, max_value=LABELS_PER_PAGE, value=1)
                
                off_x = st.number_input("Offset X (mm)", value=0.0, step=0.1)
                off_y = st.number_input("Offset Y (mm)", value=0.0, step=0.1)

                if st.button("⚡ GENERA PDF"):
                    # Validate mapping
                    missing_required = [f for f, req in field_names if req and f not in mapping]
                    if missing_required:
                        st.error(f"Campi obbligatori mancanti: {', '.join(missing_required)}")
                    else:
                        with st.spinner("Generazione PDF in corso..."):
                            try:
                                records = read_excel_data(tmp_path)
                                
                                # Use temporary path for output
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
                                )

                                with open(result_path, "rb") as f:
                                    st.download_button(
                                        label="📥 Scarica PDF",
                                        data=f,
                                        file_name=f"{os.path.splitext(uploaded_file.name)[0]}_etichette.pdf",
                                        mime="application/pdf"
                                    )
                                st.success("PDF generato con successo!")
                                
                            except Exception as e:
                                st.error(f"Errore durante la generazione: {e}")
                            finally:
                                if 'result_path' in locals() and os.path.exists(result_path):
                                    pass # Keep for download button if needed, or cleanup safely later
        
        finally:
            # Note: We should ideally cleanup the temp file, but NamedTemporaryFile(delete=False) 
            # is used because generate_pdf might need it. 
            pass

    else:
        st.info("Carica un file Excel dalla barra laterale per iniziare.")

if __name__ == "__main__":
    main()
