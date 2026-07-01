import streamlit as st
import pdfplumber
import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
import io
import re

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
                    for i in range(min(len(matr_i), len(tm_i))):
                        m = matr_i[i].strip().lstrip("0")
                        if not m: continue
                        linea = tm_i[i].split('-')[0].strip()
                        mo = mon_i[i].strip().replace('.', ':') if i < len(mon_i) else ""
                        if linea and mo:
                            database_turni[m] = f"{linea} {mo}"
    return database_turni

uploaded_excel = st.file_uploader("Carica file Excel", type=["xlsx"])
uploaded_pdf = st.file_uploader("Carica file PDF", type=["pdf"])

if uploaded_excel and uploaded_pdf:
    if st.button("🚀 Elabora e Scarica"):
        db = estrai_database_turni(uploaded_pdf)
        wb_orig = openpyxl.load_workbook(uploaded_excel)
        ws_orig = wb_orig.active
        
        wb = openpyxl.Workbook()
        ws = wb.active
        
        # Stili
        bold = Font(bold=True)
        center = Alignment(horizontal="center", vertical="center")
        left = Alignment(horizontal="left", vertical="center")
        thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        gray = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

        # Intestazione e Titolo
        ws.merge_cells('A1:C1'); ws.merge_cells('E1:G1')
        ws['A1'] = ws['E1'] = "Variazioni servizio - Giorno: 2"
        for c in ['A1', 'E1']: ws[c].font = bold; ws[c].alignment = center
        
        headers = ["Cognome", "Turno", "Assegnato"]
        for col, h in enumerate(headers, 1):
            for start in [0, 4]:
                c = ws.cell(row=2, column=start + col, value=h)
                c.font = bold; c.fill = gray; c.border = thin; c.alignment = center

        # Scrittura riga per riga
        row_target = 3
        for row in range(3, ws_orig.max_row + 1):
            # Sinistra (col 1,2) e Destra (col 5,6)
            for i, cols in enumerate([(1, 2), (5, 6)]):
                nome = ws_orig.cell(row=row, column=cols[0]).value
                turno_orig = ws_orig.cell(row=row, column=cols[1]).value
                if nome:
                    m_match = re.match(r"^(\d+)", str(nome))
                    m_str = m_match.group(1).lstrip("0") if m_match else ""
                    turno = db.get(m_str, str(turno_orig or ""))
                    
                    start_c = 0 if i == 0 else 4
                    ws.cell(row=row_target, column=start_c + 1, value=str(nome)).border = thin
                    ws.cell(row=row_target, column=start_c + 2, value=turno).border = thin
                    ws.cell(row=row_target, column=start_c + 3, value="").border = thin
            row_target += 1
        
        for col in ['A', 'E']: ws.column_dimensions[col].width = 25
        for col in ['B', 'F', 'C', 'G']: ws.column_dimensions[col].width = 15
            
        output = io.BytesIO()
        wb.save(output); output.seek(0)
        st.download_button("📥 Scarica Excel", output, "Variazioni_Finale.xlsx")
