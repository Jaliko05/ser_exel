import random
import os
from openpyxl import load_workbook

def create_report_excel(ruta_excel, posiciones, posiciones_fin, datos_txt):
    libro = load_workbook(ruta_excel)
    wb_copy = libro

    for hoja, vars_posiciones in posiciones.items():
        if hoja in datos_txt:
            hoja_excel = wb_copy[hoja]
            fila_fin, columna_fin = map(int, posiciones_fin[hoja].split(','))
            print(fila_fin, columna_fin)
            
            # Verificar si hay múltiples repeticiones
            if hoja != 'Principal':
                # Eliminar ??FIN??
                for row in hoja_excel.iter_rows():
                    for cell in row:
                        if cell.value == "??FIN??":
                            cell.value = None

            for subclave, vars_valores in datos_txt[hoja].items():
                for var_key, valor in vars_valores.items():
                    if var_key in vars_posiciones:
                        fila_base, columna = map(int, vars_posiciones[var_key].split(','))
                        fila = fila_base + int(subclave) - 1 
                        hoja_excel.cell(row=fila, column=columna, value=valor)

            # Volver a escribir ??FIN?? si había múltiples repeticiones
            # if len(datos_txt[hoja]) > 1:
            #     hoja_excel.cell(row=fila + 1, column=columna_fin, value="??FIN??")

    randon = random.randint(0, 1000000)
    nameArchive = os.path.splitext(ruta_excel)[0]
    print("nameArchive: ",nameArchive)
    newName = nameArchive + '_' + str(randon) + '.xlsx'   
    print("newName: ",newName)
    wb_copy.save(newName)