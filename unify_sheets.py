import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy

# Cargar el archivo de origen
input_file = '"C:\\Users\\javier.puentes\\ser_excel\\SIIF_IDEA\\PLANTILLAS\\FINE088_834914.xlsx"'
wb = openpyxl.load_workbook(input_file)

# Seleccionar la primera hoja para unificar las demás en ella
unified_sheet = wb.active

current_row = unified_sheet.max_row + 1

# Función para copiar el contenido y estilos de una celda a otra
def copy_cell(source_cell, target_cell):
    target_cell.value = source_cell.value
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)

# Iterar sobre todas las hojas del archivo excepto la primera
for sheet_name in wb.sheetnames[1:]:
    sheet = wb[sheet_name]
    for row in sheet.iter_rows():
        for source_cell in row:
            target_cell = unified_sheet.cell(row=current_row, column=source_cell.column)
            copy_cell(source_cell, target_cell)
        current_row += 1

# Guardar el archivo unificado
wb.save(input_file)
