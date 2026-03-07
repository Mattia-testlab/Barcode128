import streamlit as st
import pandas as pd
import io
import tempfile
import label_generator as lg # nostro backend

st.set_page_config(
    page_title="Barcode Pro Web",
    page_icon="🏷️",
    layout="wide"
)

st.title("🏷️ Barcode Pro Web")
st.markdown("Generatore di etichette barcode Code 128 (Formato A4 - Griglia 3x8).")

# ==========================================
# SIDEBAR - CARICAMENTO FILE
# ==========================================
with st.sidebar:
    st.header("📂 Carica Dati")
    uploaded_file = st.file_uploader("Carica file Excel (.xlsx)", type=["xlsx"])

# Se c'è un file caricato...
if uploaded_file is not None:
    # Salvataggio temporaneo per leggere gli header
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name
        
    # Lettura headers e auto-mapping
    try:
        headers = lg.read_excel_headers(tmp_path)
    except Exception as e:
        st.error(f"Errore nella lettura del file Excel: {e}")
        st.stop()
        
    df = pd.read_excel(tmp_path, dtype=str)
    # Pulisci i NaN e converte in stringhe
    df = df.fillna("")

    # === LOGICA AUTO-MAPPING ===
    # Tentiamo di associare i campi logici in base a parole chiave comuni
    def guess_header(options, current_headers):
        for opt in options:
            for h in current_headers:
                if opt.lower() in str(h).lower():
                    return h
        return "(nessuna)"

    default_mapping = {
        "Codice Barcode": guess_header(["barcode", "qvc", "codice", "sku"], headers),
        "Testo Superiore 1": guess_header(["cartone", "skt", "serial", "codice skt"], headers),
        "Testo Superiore 2": guess_header(["po", "acquisto", "order"], headers),
        "Testo Superiore 3": guess_header(["qta", "quantità", "quantita", "qty"], headers),
        "Numero Copie": guess_header(["copie", "ripetizioni", "n_copie", "qta"], headers)  # qta usato per n copie a volte
    }

    # Creazione delle 2 colonne principali
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("🔗 Mapping Colonne (Auto-completato)")
        st.info("Abbina le colonne del tuo file Excel ai campi dell'etichetta.")
        
        mapping = {}
        # Menù a tendina con l'opzione (nessuna) in cima
        header_options = ["(nessuna)"] + headers
        
        mapping["Codice Barcode"] = st.selectbox(
            "Codice Barcode (Obbligatorio)*", 
            header_options, 
            index=header_options.index(default_mapping["Codice Barcode"]) if default_mapping["Codice Barcode"] in header_options else 0
        )
        
        mapping["Testo Superiore 1"] = st.selectbox(
            "Testo Superiore 1 (Es. Cartone/SKT)", 
            header_options,
            index=header_options.index(default_mapping["Testo Superiore 1"]) if default_mapping["Testo Superiore 1"] in header_options else 0
        )
        
        mapping["Testo Superiore 2"] = st.selectbox(
            "Testo Superiore 2 (Es. PO)", 
            header_options,
            index=header_options.index(default_mapping["Testo Superiore 2"]) if default_mapping["Testo Superiore 2"] in header_options else 0
        )
        
        mapping["Testo Superiore 3"] = st.selectbox(
            "Testo Superiore 3 (Es. Quantità)", 
            header_options,
            index=header_options.index(default_mapping["Testo Superiore 3"]) if default_mapping["Testo Superiore 3"] in header_options else 0
        )
        
        mapping["Numero Copie"] = st.selectbox(
            "Numero Copie (Solo per SKT)", 
            header_options,
            index=header_options.index(default_mapping["Numero Copie"]) if default_mapping["Numero Copie"] in header_options else 0
        )
        
        # NOTE: Non richiediamo o mappiamo testi inferiori come da specifiche.
        st.caption("I campi non obbligatori possono essere lasciati su '(nessuna)'.")

    with col2:
        st.subheader("⚙️ Impostazioni Stampa")
        
        # Scelta profilo
        profile = st.radio("Scegli il Profilo Finale:", ["COLLI", "SKT"], horizontal=True)
        
        # Anteprima HTML colorata (Recap Grafico)
        st.markdown(f"**Preview Grafica ({profile})**")
        html_preview = ""
        if profile == "COLLI":
            html_preview = f"""
            <div style="border: 2px dashed #4b4b4b; padding: 15px; border-radius: 8px; width: 100%; text-align: center; background-color: #1e1e1e;">
                <p style="font-size: 14px; margin: 2px 0; color: #e0e0e0;"><b>{{{{ Testo Sup 1: Cartone }}}}</b></p>
                <p style="font-size: 14px; margin: 2px 0; color: #e0e0e0;"><b>PO: {{{{ Testo Sup 2: PO }}}}</b></p>
                <p style="font-size: 14px; margin: 2px 0; color: #e0e0e0;"><b>Quantità: {{{{ Testo Sup 3: Qty }}}}</b></p>
                <div style="margin-top: 10px; background-color: white; padding: 10px; border-radius: 4px;">
                    <b style="color: black; font-size: 18px;">|||| | ||||||| | |||| (BARCODE QVC)</b>
                </div>
            </div>
            """
        else: # SKT
            html_preview = f"""
            <div style="border: 2px dashed #4b4b4b; padding: 15px; border-radius: 8px; width: 100%; text-align: center; background-color: #1e1e1e;">
                <p style="font-size: 16px; margin: 2px 0; color: #e0e0e0;"><b>{{{{ Testo Sup 1: Codice SKT }}}}</b></p>
                <p style="font-size: 16px; margin: 2px 0; color: #e0e0e0;"><b>PO: {{{{ Testo Sup 2: PO }}}}</b></p>
                <div style="margin-top: 15px; background-color: white; padding: 10px; border-radius: 4px;">
                    <b style="color: black; font-size: 18px;">|||| | ||||||| | |||| (BARCODE QVC)</b>
                </div>
            </div>
            """
        st.markdown(html_preview, unsafe_allow_html=True)
        st.write("")
        
        # Posizione di inizio e offset
        row1, row2, row3 = st.columns(3)
        with row1:
            start_pos = st.number_input("Posizione Start (1-24)", min_value=1, max_value=24, value=1)
        with row2:
            offset_x = st.number_input("Offset X (mm)", value=0.0, step=0.1)
        with row3:
            offset_y = st.number_input("Offset Y (mm)", value=0.0, step=0.1)
            
        # Regolazioni di layout
        with st.expander("🎛️ Regolazioni Layout"):
            layout_overrides = {}
            layout_overrides["margin_y"] = st.slider("Margine interno Y (mm)", min_value=0.5, max_value=6.0, value=4.5, step=0.1)
            layout_overrides["text_barcode_spacing"] = st.slider("Spazio testo-barcode (mm)", min_value=0.1, max_value=4.0, value=0.5, step=0.1)
            layout_overrides["line_spacing"] = st.slider("Spaziatura righe testo (mm)", min_value=1.5, max_value=5.0, value=2.2 if profile == "COLLI" else 2.5, step=0.1)
            layout_overrides["font_size"] = st.slider("Dimensione font superiore (pt)", min_value=5, max_value=14, value=7 if profile == "COLLI" else 9, step=1)
            layout_overrides["barcode_height"] = st.slider("Altezza max barcode (mm)", min_value=15.0, max_value=35.0, value=19.0, step=0.5)

    # ==========================================
    # SEZIONE FINALE - GENERAZIONE PDF
    # ==========================================
    st.divider()
    
    # Controlli di validazione
    if mapping["Codice Barcode"] == "(nessuna)":
        st.warning("⚠️ Seleziona una colonna per il 'Codice Barcode' per poter generare il PDF.")
    else:
        # Generazione su pulsante per non bloccare la UI, ma se vogliamo real-time, 
        # possiamo generarlo e metterlo nel Download Button immediatamente.
        with st.spinner("Generazione PDF in corso..."):
            try:
                pdf_path = lg.generate_pdf(
                    df=df,
                    profile=profile,
                    mapping=mapping,
                    start_position=start_pos,
                    offset_x=offset_x,
                    offset_y=offset_y,
                    layout_overrides=layout_overrides
                )
                
                with open(pdf_path, "rb") as pdf_file:
                    pdf_bytes = pdf_file.read()
                    
                st.success("✅ Generazione completata con successo! Clicca il pulsante qui sotto per scaricare il PDF aggiornato.")
                
                # Real-time ready per il download
                st.download_button(
                    label="📥 Scarica PDF Aggiornato",
                    data=pdf_bytes,
                    file_name=f"etichette_{profile.lower()}_{start_pos}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Errore durante la generazione del PDF: {e}")
                
else:
    st.info("👈 Carica un file `.xlsx` dalla barra laterale per iniziare.")
