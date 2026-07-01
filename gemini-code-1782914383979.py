# --- Creazione nuovo file con formattazione migliorata ---
            wb = openpyxl.Workbook()
            ws = wb.active
            
            # Stili per abbellire
            bordo_sottile = Side(border_style="thin", color="000000")
            bordo = Border(left=bordo_sottile, right=bordo_sottile, top=bordo_sottile, bottom=bordo_sottile)
            allineamento = Alignment(horizontal="left", vertical="center")

            max_r = math.ceil(len(dati) / 2) if dati else 1
            
            for i, d in enumerate(dati):
                lato = 0 if i < max_r else 1
                r = (i % max_r) + 3
                c = 1 if lato == 0 else 5
                
                # Scrittura celle
                c1 = ws.cell(row=r, column=c, value=d['c'])
                c2 = ws.cell(row=r, column=c+1, value=d['t'])
                
                # Applicazione stili
                for cella in [c1, c2]:
                    cella.border = bordo
                    cella.alignment = allineamento
                
                ws.row_dimensions[r].height = 20

            # Impostazione larghezza colonne (importante per non vedere i tagli)
            ws.column_dimensions['A'].width = 25  # Cognome/Matr sinistra
            ws.column_dimensions['B'].width = 20  # Turno sinistra
            ws.column_dimensions['C'].width = 2   # Spazio centrale
            ws.column_dimensions['E'].width = 25  # Cognome/Matr destra
            ws.column_dimensions['F'].width = 20  # Turno destra
            
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
