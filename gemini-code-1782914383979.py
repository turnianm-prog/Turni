import streamlit as st
import pdfplumber
import openpyxl
import io
import re
import math

st.set_page_config(page_title="Gestione Turni", page_icon="🔄")
st.title("🔄 Generatore Variazioni Turni")

def estrai_database_turni(pdf_file):
    database_turni = {}
    with pdfplumber.open(pdf_file) as pdf:
        for pagina in pdf.pages:
            testo = pagina.extract_text()
            if not testo: continue
            
            righe = testo.split('\n')
            for riga in righe:
                # Cerca matricola (cifre) seguita da altro testo
                # Questa regex è più generica per catturare la riga
                match = re.search(r'(\d+)\s+(.+)', riga)
                if match:
                    matr = match.group(1).lstrip("0")
                    # Salviamo l'intera riga trovata per sicurezza
                    database_turni[matr] = match.group(2).strip()
    return database_turni

uploaded_excel = st.file_uploader("Carica il file Excel delle Variazioni", type=["xlsx"])
uploaded_pdf = st.file_uploader("Carica il Tabellone Generale (PDF)", type=["pdf"])

if uploaded_excel and uploaded_pdf:
    if st.button("🚀 Elabora e Scarica"):
        with st.spinner("Elaborazione in corso..."):
            db = estrai_database_turni(uploaded_pdf)
            wb_orig = openpyxl.load_workbook(uploaded_excel)
            ws_orig = wb_orig.active
            
            dati = []
            for row in range(3, ws_orig.max_row + 1):
                # Cerchiamo matricola nelle colonne 1, 2, 3 o 5, 6, 7
                for col_idx in [1, 5]: 
                    val_cognome = ws_orig.cell(row=row, column=col_idx).value
                    if val_cognome:
                        m_match = re.search(r"(\d+)", str(val_cognome))
                        m_str = m_match.group(1).lstrip("0") if m_match else ""
                        
                        # Recupera il turno dal DB se la matricola esiste
                        turno_trovato = db.get(m_str, "N.D.")
                        
                        dati.append({"c": str(val_cognome), "t": turno_trovato})
            
            wb_new = openpyxl.Workbook()
            ws_new = wb_new.active
            
            max_r = math.ceil(len(dati) / 2) if dati else 1
            for i, d in enumerate(dati):
                lato = 0 if i < max_r else 1
                r = (i % max_r) + 3
                c = 1 if lato == 0 else 5
                ws_new.cell(row=r, column=c, value=d['c'])
                ws_new.cell(row=r, column=c+1, value=d['t'])
            
            buffer = io.BytesIO()
            wb_new.save(buffer)
            buffer.seek(0)
            
            st.success("Elaborazione completata!")
            st.download_button(
                label="📥 Scarica Excel",
                data=buffer.getvalue(),
                file_name="Variazioni_Finale.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
