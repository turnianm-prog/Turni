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
    testo_estratto = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for pagina in pdf.pages:
            testo = pagina.extract_text()
            if testo:
                testo_estratto.append(testo)
                righe = testo.split('\n')
                for riga in righe:
                    # Regex flessibile: cerca numeri (matricola) seguiti da testo e orari
                    match = re.search(r'(\d+)\s+([A-Za-z0-9\-]+)\s+(\d{1,2}[:.]\d{2})', riga)
                    if match:
                        matr = match.group(1).lstrip("0")
                        valore = match.group(0)
                        database_turni[matr] = valore
                        
    return database_turni, "\n".join(testo_estratto)

uploaded_excel = st.file_uploader("Carica il file Excel delle Variazioni", type=["xlsx"])
uploaded_pdf = st.file_uploader("Carica il Tabellone Generale (PDF)", type=["pdf"])

if uploaded_pdf:
    if st.button("🔍 Diagnostica PDF"):
        db, testo_debug = estrai_database_turni(uploaded_pdf)
        st.subheader("Anteprima testo letto:")
        st.text_area("Testo estratto", testo_debug[:1000], height=200)
        st.write(f"Turni trovati nel database: {len(db)}")
        st.write("Esempio dati trovati:", list(db.items())[:5])

if uploaded_excel and uploaded_pdf:
    if st.button("🚀 Elabora e Scarica"):
        with st.spinner("Elaborazione..."):
            db, _ = estrai_database_turni(uploaded_pdf)
            wb_orig = openpyxl.load_workbook(uploaded_excel)
            ws_orig = wb_orig.active
            
            dati = []
            for row in range(3, ws_orig.max_row + 1):
                for cols in [(1,2,3), (5,6,7)]:
                    val_cog = ws_orig.cell(row=row, column=cols[0]).value
                    if val_cog:
                        m_match = re.search(r"(\d+)", str(val_cog))
                        m_str = m_match.group(1).lstrip("0") if m_match else ""
                        turno = str(ws_orig.cell(row=row, column=cols[1]).value or "")
                        if m_str in db:
                            turno = db[m_str]
                        dati.append({"c": str(val_cog), "t": turno})
            
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
            
            st.download_button(
                label="📥 Scarica Excel",
                data=buffer.getvalue(),
                file_name="Variazioni_Finale.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
