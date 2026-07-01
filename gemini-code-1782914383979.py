# ... (tutto il codice prima rimane uguale)
            
            # Creazione nuovo file in memoria
            wb = openpyxl.Workbook()
            ws = wb.active
            # ... (manteniamo le righe di formattazione che avevi)
            
            # SALVATAGGIO OTTIMIZZATO
            buffer = io.BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            
            st.success("Operazione completata! Il file è pronto.")
            
            st.download_button(
                label="📥 Scarica Excel Formattato",
                data=buffer.getvalue(), # Passiamo il contenuto esplicito
                file_name="Variazioni_Finale.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
