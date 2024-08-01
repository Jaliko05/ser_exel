def get_data_report_txt(ruta_archivo):
    resultado = {}
    contador_hojas = {}

    with open(ruta_archivo, 'r', encoding='utf-8') as archivo:
        lineas = archivo.readlines()

        # Omitir la primera línea
        lineas = lineas[1:]

        for linea in lineas:
            partes = linea.split('')
            if not partes or len(partes) < 2:
                continue

            hoja = partes[0]
            datos = partes[1:]

            if hoja not in resultado:
                resultado[hoja] = {}

            if hoja not in contador_hojas:
                contador_hojas[hoja] = 1
            else:
                contador_hojas[hoja] += 1

            subclave = str(contador_hojas[hoja]).zfill(3)

            resultado[hoja][subclave] = {}
            for i, valor in enumerate(datos, start=1):
                var_key = f"<VAR{i:03d}>"
                resultado[hoja][subclave][var_key] = valor.strip()

    return resultado

# # Ejemplo de uso
# ruta_archivo = "C:\\Users\\javier.puentes\\ser_excel\\SIIF_IDEA\\AYUDAS\\PAGINAS\\reportes\\000015656.txt"
# datos = get_data_report_txt(ruta_archivo)
# print(datos)
