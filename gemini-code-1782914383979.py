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
            # Estrazione testuale riga per riga per maggiore compatibilità
            testo = pagina.extract_text()
            if not testo: continue
            
            # Cerca pattern comuni: Matricola (numeri) - Linea/Turno - Orario
            # Questo è un approccio universale se la tabella non viene letta come tabella
            righe = testo.split('\n')
            for riga in righe:
                # Cerca una sequenza numerica (matricola) seguita da testo
                match = re.search(r'(\d+)\s+([A-Z0-9\-]+)\s+(\d{1,2}[:.]\d{2})\s+(\d{1,2}[:.]\d{2})', riga)
                if match:
                    matr = match.group(1).lstrip("0")
                    linea = match.group(2)
                    inizio = match.group(3).replace('.', ':')
                    durata = match.group(4).replace('.', ':')
                    database_turni[matr] = f"{linea} {inizio} {durata}"
    return database_turni

uploaded_excel = st.file_uploader("Carica il file Excel delle Variazioni", type=["xlsx"])
uploaded_pdf = st.file_uploader("Carica il Tabellone Generale (PDF)", type=["pdf"])

if uploaded_excel and uploaded_pdf:
    if st.button("🚀 Elabora e Scarica"):
        with st.spinner("Elaborazione in corso..."):
            db = estrai_database_turni(uploaded_pdf)
            
            if not db:
                st.warning("Non ho trovato dati nel PDF. Verifica che il PDF contenga testo selezionabile.")
            else:
                st.write(f"Trovati {len(db)} turni nel database.")
                
                wb_orig = openpyxl.load_workbook(uploaded_excel)
                ws_orig = wb_orig.active
                
                dati = []
                for row in range(3, ws_orig.max_row + 1):
                    for cols in [(1,2,3), (5,6,7)]:
                        val_cog = ws_orig.cell(row=row, column=cols[0]).value
                        if val_cog:
                            cog_str = str(val_cog).strip()
                            matr_match = re.search(r"(\d+)", cog_str)
                            m_str = matr_match.group(1).lstrip("0") if matr_match else ""
                            
                            turno = str(ws_orig.cell(row=row, column=cols[1]).value or "")
                            if m_str in db:
                                turno = db[m_str]
                            dati.append({"c": cog_str, "t": turno})
                
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
