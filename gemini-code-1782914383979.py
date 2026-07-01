import streamlit as st
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import re
import math

st.set_page_config(page_title="Gestione Turni", page_icon="🔄")
st.title("🔄 Generatore Variazioni Turni")

def estrai_database_turni(pdf_file):
    database_turni = {}
    with pdfplumber.open(pdf_file) as pdf:
        for pagina in pdf.pages:
            tabella = pagina.extract_table()
            if not tabella: continue
            for riga in tabella:
                if len(riga) >= 9 and riga[0]:
                    matr_i = str(riga[0]).split('\n')
                    tm_i = str(riga[5] if riga[5] else "").split('\n')
                    mon_i = str(riga[6] if riga[6] else "").split('\n')
                    dur_i = str(riga[8] if len(riga) > 8 and riga[8] else "").split('\n')
                    
                    for i in range(min(len(matr_i), len(tm_i))):
                        m = matr_i[i].strip().lstrip("0")
                        if not m: continue
                        linea = tm_i[i].split('-')[0].strip() if i < len(tm_i) else ""
                        mo = mon_i[i].strip().replace('.', ':') if i < len(mon_i) else ""
                        du = dur_i[i].strip().replace('.', ':') if i < len(dur_i) else ""
                        if linea and mo and du:
                            database_turni[m] = f"{linea} {mo} {du}"
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
                for cols in [(1,2,3), (5,6,7)]:
                    val_cog = ws_orig.cell(row=row, column=cols[0]).value
                    if val_cog:
                        cog_str = str(val_cog).strip()
                        matr = re.match(r"^(\d+)", cog_str)
                        m_str = matr.group(1).lstrip("0") if matr else ""
                        turno = str(ws_orig.cell(row=row, column=cols[1]).value or "")
                        if m_str in db: turno = db[m_str]
                        dati.append({"m": m_str, "c": cog_str, "t": turno})
            
            # Creazione nuovo file in memoria
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.column_dimensions['A'].width = 18.7
            ws.column_dimensions['E'].width = 18.7
            
            max_r = math.ceil(len(dati) / 2) if dati else 1
            for i, d in enumerate(dati):
                lato = 0 if i < max_r else 1
                r = (i % max_r) + 3
                c = 1 if lato == 0 else 5
                ws.cell(row=r, column=c, value=d['c'])
                ws.cell(row=r, column=c+1, value=d['t'])
                ws.row_dimensions[r].height = 21.0
            
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.success("Operazione completata!")
            st.download_button(
                label="📥 Scarica Excel Formattato",
                data=output,
                file_name="Variazioni_Finale.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
