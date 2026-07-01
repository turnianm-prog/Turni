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
                        linea = tm_i[i].split('-')[0].strip
