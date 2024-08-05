from log import log
from create_report_excel import create_report_excel
from get_data_report_txt import get_data_report_txt
from get_data_template_excel import get_data_template_excel
import sys
import os
from get_txt import get_routs, get_name_template
from pathlib import Path
from xls2xlsx import XLS2XLSX
from convert_xls_to_xlsx import convert_xls_to_xlsx
from unify_sheets import unify_sheets

def main():
    #obtener datos de la linea de comandos
    params = sys.argv


    #Vatiables generales
    separator = '' 
    # name_report_txt = params[1]
    # number_session = params[2]
    name_report_txt = '000007453'

    rout_aplication = str(Path(__file__).parent.absolute())# ruta SIIFNET
    print(rout_aplication)

    #Armar rutas
    # rout_environment = os.path.dirname(rout_aplication) #ruta del ambiente IDEA
    # print(rout_environment)
    rout_environment = "C:\\Users\\javier.puentes\\ser_excel"
    
    rout_log = rout_environment + "\\" + get_routs(rout_aplication,11).strip()
    print(rout_log)
    

    #Variables iniciales para log
    nameAplication = "ser_excel"
    message = "Iniciando aplicación ser_excel \n"
 

    # if params.count > 1:
    rout_fiel_txt = rout_environment + "\\" + get_routs(rout_aplication,24).strip() + name_report_txt + '.txt' #ruta del reporte txt
    message = message + "Ruta del archivo de reporte: " + rout_fiel_txt + "\n"
    print(rout_fiel_txt)

    if os.path.exists(rout_fiel_txt):
        name_template = get_name_template(rout_fiel_txt,separator)
        print(name_template)
        rout_template_excel = rout_environment + "\\" + get_routs(rout_aplication,4).strip() + name_template + '.xlsx' #ruta del template excel
        print(rout_template_excel)
        if not os.path.exists(rout_template_excel):
            rout_template_excel_xls = rout_environment + "\\" + get_routs(rout_aplication,4).strip() + name_template + '.xls' 
            print(rout_template_excel_xls)
            print("template xls, convertir a xlsx")
            message = message + " Plantilla con extension xls, se convertirá a xlsx \n"
            try:
                convert_xls_to_xlsx(rout_template_excel_xls, rout_template_excel)
            except Exception as e:
                print(e)
                message = message + "Error al convertir la plantilla xls a xlsx" + "\n"
        
        if os.path.exists(rout_template_excel):
            message = message + "Ruta del archivo de plantilla: " + rout_template_excel + "\n"
            datos_txt = get_data_report_txt(rout_fiel_txt)
            posiciones_excel, posiciones_estilos,  posiciones_fin = get_data_template_excel(rout_template_excel)
            name_report_excel = create_report_excel(rout_template_excel, posiciones_excel, posiciones_estilos, posiciones_fin, datos_txt)
            #unificacion de las hojas
            unify_sheets(name_report_excel)
        else:
            print("No existe el archivo de la plantilla")
            message = message + "No existe el archivo de la plantilla" + "\n"

    else:
        print("No existe el archivo de reporte")
        message = message + "No existe el archivo de reporte" + "\n"

    log(rout_log, nameAplication, message)
    
main()
