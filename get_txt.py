def get_routs(routAplication, pos):
    rout_rutas = routAplication + '\\rutas.txt'
    pos = int(pos)
    with open(rout_rutas, 'r') as archivo:
        for index, line in enumerate(archivo, start=1):
            if index == pos:
                return line.strip()  # Devuelve la línea en la posición solicitada sin espacios adicionales
        return None


def get_name_template(ruta_archivo,separator):
    with open(ruta_archivo, 'r') as archivo:
        primera_linea = archivo.readline().strip()  
        datos = primera_linea.split(separator)  
        return datos[0]