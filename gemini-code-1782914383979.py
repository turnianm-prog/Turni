import streamlit as st
import pdfplumber
import openpyxl
from openpyxl.styles import Alignment, Border, Side
import io
import re
import math

# Configurazione pagina
st.set_page_config(page_title="Gestione Turni", page_icon="🔄")
st.title("🔄 Generatore Variazioni Turni")

# --- FUNZIONE DI ESTRAZIONE ---
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

# --- INTERFACCIA E LOGICA ---
uploaded_excel = st.file_uploader("Carica il file Excel delle Variazioni", type=["xlsx"])
uploaded_pdf = st.file_uploader("Carica il Tabellone Generale (PDF)", type=["pdf"])

if uploaded_excel and uploaded_pdf:
    if st.button("🚀 Elabora e Scarica"):
        with st.spinner("Elaborazione in corso..."):
            db = estrai_database_turni(uploaded_pdf)
            wb_orig = openpyxl.load_workbook(uploaded_excel)
            ws_orig = wb_orig.active
            
            dati = []
            # Scansione di tutte le righe nell'Excel originale
            for row in range(3, ws_orig.max_row + 1):
                for col_cognome, col_turno in [(1, 2), (5, 6)]:
                    val_cog = ws_orig.cell(row=row, column=col_cognome).value
                    if val_cog:
                        cog_str = str(val_cog).strip()
                        matr = re.match(r"^(\d+)", cog_str)
                        m_str = matr.group(1).lstrip("0") if matr else ""
                        
                        # Se la matricola è nel PDF prendiamo il turno, altrimenti scriviamo "-"
                        turno = db.get(m_str, "-") 
                        
                        dati.append({"c": cog_str, "t": turno})
            
            # Creazione nuovo file
            wb = openpyxl.Workbook()
            ws = wb.active
            
            bordo_sottile = Side(border_style="thin", color="000000")
            bordo = Border(left=bordo_sottile, right=bordo_sottile, top=bordo_sottile, bottom=bordo_sottile)
            allineamento = Alignment(horizontal="left", vertical="center")
            
            # Calcolo per dividere i dati in due colonne
            max_r = math.ceil(len(dati) / 2) if dati else 1
            for i, d in enumerate(dati):
                lato = 0 if i < max_r else 1
                r = (i % max_r) + 3
                c = 1 if lato == 0 else 5
                
                c1 = ws.cell(row=r, column=c, value=d['c'])
                c2 = ws.cell(row=r, column=c+1, value=d['t'])
                
                for cella in [c1, c2]:
                    cella.border = bordo
                    cella.alignment = allineamento
                ws.row_dimensions[r].height = 20
            
            # Impostazioni larghezza
            ws.column_dimensions['A'].width = 25
            ws.column_dimensions['B'].width = 20
            ws.column_dimensions['C'].width = 2
            ws.column_dimensions['E'].width = 25
            ws.column_dimensions['F'].width = 20
            
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.success("Operazione completata! Tutti i nomi sono stati inclusi.")
            st.download_button(
                label="📥 Scarica Excel Formattato",
                data=output,
                file_name="Variazioni_Finale.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
