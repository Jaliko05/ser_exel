import random
import os
from openpyxl import load_workbook
from openpyxl.styles import Font

def create_report_excel(ruta_excel, posiciones, posiciones_estilos, posiciones_fin, datos_txt):
    libro = load_workbook(ruta_excel)
    wb_copy = libro

    for hoja, vars_posiciones in posiciones.items():
        if hoja in datos_txt:
            hoja_excel = wb_copy[hoja]
            fila_fin, columna_fin = map(int, posiciones_fin[hoja].split(','))
            print(fila_fin, columna_fin)
            

            for subclave, vars_valores in datos_txt[hoja].items():
                for var_key, valor in vars_valores.items():
                    if var_key in vars_posiciones:
                        fila_base, columna = map(int, vars_posiciones[var_key].split(','))
                        fila = fila_base + int(subclave) - 1 
                        hoja_excel.cell(row=fila, column=columna, value=valor)
                        

            # Volver a escribir ??FIN?? si había múltiples repeticiones
            if len(datos_txt[hoja]) > 1:
                if hoja != 'Principal':
                # Eliminar ??FIN?? si son rep
                    for row in hoja_excel.iter_rows():
                        for cell in row:
                            if cell.value == "??FIN??":
                                cell.value = None
                # Volver a escribir ??FIN?? en la ultima fila de la rep
                cell = hoja_excel.cell(row=fila, column=columna_fin, value="??FIN??")
                cell.font = Font(color="FFFFFF")

    randon = random.randint(0, 1000000)
    nameArchive = os.path.splitext(ruta_excel)[0]
    print("nameArchive: ",nameArchive)
    newName = nameArchive + '_' + str(randon) + '.xlsx'   
    print("newName: ",newName)
    wb_copy.save(newName)
    return newName