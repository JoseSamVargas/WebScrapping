from datetime import datetime
import glob
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import math
from bs4 import BeautifulSoup
from openpyxl import load_workbook

ruta_grabacion = r'C:\Users\ws-samlap\.spyder-py3'
nombre_carpeta= 'RUCS'
ruta_excel_RUCS = r'C:\Users\ws-samlap\.spyder-py3\RUCS'
nombre_excel_RUCS = 'RUC.xlsx'
ruta_chromedriver = r'C:\Users\ws-samlap\.spyder-py3\Chronium\92'

path = ruta_grabacion + '\\' + nombre_carpeta
path_excel_RUCS = ruta_excel_RUCS + '\\' + nombre_excel_RUCS

driver = webdriver.Chrome(executable_path = ruta_chromedriver + '\\' + 'chromedriver.exe')
driver.get('http://serviciosweb.digemid.minsa.gob.pe/Consultas/Establecimientos')

criterio = Select(driver.find_element_by_id('param1'))
criterio.select_by_value('2')

tiempo_inicio = datetime.now()
mes = tiempo_inicio.strftime('%m')
año = tiempo_inicio.strftime('%Y')

lista_años = glob.glob(path +'\\' + '2' + '*')

if lista_años:
    path_ultimo_año = max(lista_años, key=os.path.getctime)
    ultimo_año = os.path.basename(os.path.normpath(path_ultimo_año))
    lista_RUCS_mes_vez = glob.glob(path_ultimo_año +'\\' + 'RUCS' + '*')
    path_ultimo_mes_vez = max(lista_RUCS_mes_vez, key=os.path.getctime)
    ultima_fecha = os.path.getctime(path_ultimo_mes_vez)
    ultima_cadena = os.path.basename(os.path.normpath(path_ultimo_mes_vez))
    ultimo_sufijo = ultima_cadena[7:]
    ultimo_mes = ultima_cadena[7:9]
   
    if (mes != ultimo_mes) | (año != ultimo_año):
        os.makedirs(path + '\\' + año + '\\' + 'RUCS' + ' - ' + mes)

    elif len(ultima_cadena) == 9:
        os.makedirs(path + '\\' + año + '\\' + 'RUCS' + ' - ' + mes + ' - ' + str(2))

    elif len(ultima_cadena) >= 13:
        i = 1
        while True:
            if len(ultima_cadena) == i+12:
                    vez = os.path.basename(os.path.normpath(path_ultimo_mes_vez))[-i:]
                    os.makedirs(path + '\\' + año + '\\' + 'RUCS' + ' - ' + mes + ' - ' + str(int(vez) + 1))
                    break
            i = i + 1

else: 
    os.makedirs(path + '\\' + año + '\\' + 'RUCS' + ' - ' + mes)

lista_nueva = glob.glob(path + '\\' + año + '\\' + 'RUCS' + ' - ' + mes + '*')
path_nuevo = max(lista_nueva, key=os.path.getctime)
cadena_nueva = os.path.basename(os.path.normpath(path_nuevo))
sufijo_nuevo = cadena_nueva[7:]

excelRUCS = pd.read_excel(path_excel_RUCS)
excelRUCS = excelRUCS.iloc[:,-2:]
excelRUCS = excelRUCS.dropna()
excelRUCS.iloc[:,0] = excelRUCS.iloc[:,0].astype('int64')

cadena_excel_anteriores = []
lista_excel_anteriores_limpia = []
compacto_anterior = pd.DataFrame()
lista_empresas = []
tabla_compacto = pd.DataFrame()
tabla_nuevos = pd.DataFrame()

encabezados = ['Detalle', 'Item', 'NºRegistro', 'Cat.', 'Nombre Comercial', 'Razón Social', 'R.U.C', 'Dirección', 'Ubigeo', 'Situación', 'Empadronado']
dictado_encabezados = {}
for i in range(len(encabezados)):
    dictado_encabezados[i] = encabezados[i]

if lista_años:
    path_compacto_anterior = path_ultimo_mes_vez + '\\' + 'Compacto' + ' - ' + ultimo_sufijo + '.xlsx'
    compacto_anterior = pd.read_excel(path_compacto_anterior, index_col = 0, dtype = str)

for i in range(len(excelRUCS)):
    RUC = str(excelRUCS.iloc[i,0])
    empresa = excelRUCS.iloc[i,1]

    lista_empresas.append(empresa)
    cuenta = lista_empresas.count(empresa)
    info = []
    hoja_actual = 1   
    tabla = pd.DataFrame()

    while True:
        driver.find_element_by_id('param2').clear()
        driver.find_element_by_id('param2').send_keys(RUC)
        driver.find_element_by_id('btn_consultar2').click()

        try:
            WebDriverWait(driver, 30).until(EC.text_to_be_present_in_element((By.XPATH, '//*[@id="tresultados"]/tbody/tr[1]/td[7]'), RUC))
            break
        
        except:
            pass

    registros = int(driver.find_element_by_xpath('//*[@id="tresultados"]/thead[1]/tr[2]/td/a/b').text)
    filas = len(driver.find_elements_by_xpath('//*[@id="tresultados"]/tbody/tr'))
    hojas = math.ceil(registros/filas)
    
    for i in range(hojas):
        html = driver.page_source
        soup = BeautifulSoup(html,'html.parser')
        div = soup.select_one("table#tresultados")
        tabla_hoja = pd.read_html(str(div), converters = {2 : str, 6 : str})[0]
        tabla_hoja.shape[1]
        tabla_hoja.columns = range(tabla_hoja.shape[1])
        tabla = tabla.append(tabla_hoja)
        
        i = i + 1

        if i < hojas:
            while True:
                if hoja_actual > 1:
                    driver.find_element_by_xpath('//*[@id="tresultados"]/thead[2]/tr/td/a['+str(hojas+1)+']').click()

                else:
                    driver.find_element_by_xpath('//*[@id="tresultados"]/thead[2]/tr/td/a[1]').click()
                    
                try:                    
                    WebDriverWait(driver, 30).until(EC.text_to_be_present_in_element((By.XPATH, '//*[@id="tresultados"]/tbody/tr[1]/td[2]'), str(25*(hoja_actual) + 1)))
                    break
                
                except:
                    pass

            hoja_actual = hoja_actual + 1

    tabla.rename(columns = dictado_encabezados, inplace = True)
    tabla = tabla.drop(['Detalle'], axis = 1)
    tabla = tabla.set_index('Item')

    if cuenta == 1:
        tabla.to_excel(path_nuevo + '\\' + empresa + ' - ' + sufijo_nuevo + '.xlsx', sheet_name = 'Hoja1')

    else:
        path_excel = path_nuevo + '\\' + empresa + ' - ' + sufijo_nuevo + '.xlsx'
        book = load_workbook(path_excel)
        writer = pd.ExcelWriter(path_excel, engine = 'openpyxl')
        writer.book = book

        tabla.to_excel(writer, sheet_name = 'Hoja' + str(cuenta))
        writer.save()
        writer.close()

    tabla = tabla.copy()
    tabla['Marca'] = empresa
    tabla['Hoja'] = 'Hoja' + str(cuenta)
    tabla_compacto = tabla_compacto.append(tabla)

    if lista_años:
        compacto_anterior_espefico = compacto_anterior[compacto_anterior['R.U.C'] == RUC]
        compacto_anterior_espefico = compacto_anterior_espefico.copy()
        compacto_anterior_espefico['Observaciones'] = 'Duplicado'
    
        if len(compacto_anterior_espefico) != 0:
            tabla_doble = tabla.append(compacto_anterior_espefico)
            nuevos = tabla_doble.drop_duplicates(subset=['Dirección'], keep = False)
            tabla_nuevos = tabla_nuevos.append(nuevos)
            compacto_anterior_espefico = compacto_anterior_espefico.drop(columns=['Observaciones'])
    
tabla_compacto.to_excel(path_nuevo + '\\' + 'Compacto' + ' - ' + sufijo_nuevo + '.xlsx', sheet_name = 'Hoja1')

if lista_años:
    writer = pd.ExcelWriter(path_nuevo + '\\' + 'Nuevos - ' + sufijo_nuevo + '.xlsx', engine = 'xlsxwriter')
    workbook  = writer.book
    tabla_nuevos.to_excel(writer, sheet_name = 'Hoja1')

    worksheet = writer.sheets['Hoja1']
    formato = workbook.add_format({'bold': True})

    if len(tabla_nuevos) == 0:
        i = 3
        worksheet.write(1, 0, 'No hay nuevos locales')

    else:
        i = 2

    worksheet.write(len(tabla_nuevos) + i, 0, 'Inicio:', formato)
    worksheet.write(len(tabla_nuevos) + i, 1, datetime.fromtimestamp(ultima_fecha).strftime('%d/%m/%Y'))
    worksheet.write(len(tabla_nuevos) + (i+1), 0, 'Fin:', formato)
    worksheet.write(len(tabla_nuevos) + (i+1), 1, tiempo_inicio.strftime('%d/%m/%Y'))

    writer.save()
    
    if (mes != ultimo_mes) | (año != ultimo_año):
        path_consolidados = path + '\\' + '2*' + '\\' + 'Consolidados'
        lista_consolidados = glob.glob(path_consolidados + '\\' + 'Consolidado' + '*')
        
        if lista_consolidados:
            ultimo_consolidado = max(lista_consolidados, key=os.path.getctime)
            fecha_cierre_anterior = pd.read_excel(ultimo_consolidado).iloc[-1,1]
            fecha_apertura = datetime.strptime(fecha_cierre_anterior, '%d/%m/%Y')
            
            if fecha_apertura.month == int(ultimo_mes):
                apertura = path + '\\' + str(fecha_apertura.year) + '\\' + 'RUCS' + ' - ' + ultimo_mes + '\\' + 'Compacto' + ' - ' + ultimo_mes + '.xlsx'
            
            else:
                lista_apertura = glob.glob(path + '\\' + str(fecha_apertura.year) + '\\' + 'RUCS' + ' - ' + '{:02d}'.format(fecha_apertura.month) + '*')
                carpeta_apertura = max(lista_apertura, key=os.path.getctime)
                cadena_apertura = os.path.basename(os.path.normpath(carpeta_apertura))
                sufijo_apertura = cadena_apertura[7:]
                apertura = carpeta_apertura + '\\' + 'Compacto' + ' - ' + sufijo_apertura  + '.xlsx'
             
        else:         
            path_primera_carpeta = path_ultimo_año + '\\' + 'RUCS' + ' - ' + ultimo_mes
            fecha_apertura = datetime.fromtimestamp(os.path.getctime(path_primera_carpeta))
            apertura = path_primera_carpeta + '\\' + 'Compacto' + ' - ' + ultimo_mes + '.xlsx'
            
        compacto_apertura = pd.read_excel(apertura, index_col = 0, dtype = str)
        compacto_apertura = compacto_apertura.copy()
        compacto_apertura['Observaciones'] = 'Duplicado'
        
        cierre_cero = datetime(int(año), int(mes), 1)
        cierre_post = tiempo_inicio
        cierre_pre = datetime.fromtimestamp(os.path.getctime(path_ultimo_mes_vez))
        
        minimo = min(cierre_pre - cierre_cero, cierre_post - cierre_cero, key = abs)
        
        if minimo == cierre_pre - cierre_cero:
            cierre = path_ultimo_mes_vez + '\\' + 'Compacto' + ' - ' + ultimo_sufijo + '.xlsx'
            fecha_cierre = cierre_pre
        
            if apertura == cierre:
                cierre = path_nuevo + '\\' + 'Compacto' + ' - ' + sufijo_nuevo + '.xlsx'
                fecha_cierre = cierre_post
        
        else:
            cierre = path_nuevo + '\\' + 'Compacto' + ' - ' + sufijo_nuevo + '.xlsx'
            fecha_cierre = cierre_post
                
        compacto_cierre = pd.read_excel(cierre, index_col = 0, dtype = str)

        tabla_compacto_doble = compacto_apertura.append(compacto_cierre)
        tabla_consolidados = tabla_compacto_doble.drop_duplicates(subset=['Dirección'], keep = False)
        compacto_apertura = compacto_apertura.drop(columns=['Observaciones'])

        año_consolidados = str(fecha_cierre.year)
        
        if fecha_cierre.year != fecha_apertura.year:
            corte_año = datetime(fecha_apertura.year + 1, 1, 1)
            maximo = max(corte_año - fecha_apertura, fecha_cierre - corte_año, key = abs)
                
            if maximo == corte_año - fecha_apertura:
                año_consolidados = str(fecha_apertura.year)
        
        carpeta_nuevo_consolidado = path + '\\' + año_consolidados + '\\' + 'Consolidados' 

        if not os.path.exists(carpeta_nuevo_consolidado):
            os.makedirs(carpeta_nuevo_consolidado)
        
        writer = pd.ExcelWriter(carpeta_nuevo_consolidado + '\\' + 'Consolidado' + ' ' + fecha_cierre.strftime('%Y-%m-%d') + '.xlsx', engine = 'xlsxwriter')
        workbook  = writer.book
        tabla_consolidados.to_excel(writer, sheet_name = 'Hoja1')  
        
        formato = workbook.add_format({'bold': True})
        worksheet = writer.sheets['Hoja1']

        if len(tabla_consolidados) == 0:
            i = 3
            worksheet.write(1, 0, 'No hay nuevos locales')

        else:
            i = 2

        worksheet.write(len(tabla_consolidados) + i, 0, 'Inicio:', formato)
        worksheet.write(len(tabla_consolidados) + i, 1, fecha_apertura.strftime('%d/%m/%Y'))
        worksheet.write(len(tabla_consolidados) + (i+1), 0, 'Fin:', formato)
        worksheet.write(len(tabla_consolidados) + (i+1), 1, fecha_cierre.strftime('%d/%m/%Y'))

        writer.save()

driver.close()

tiempo_fin = datetime.now()
print('Tiempo de ejecución: {}'.format(tiempo_fin - tiempo_inicio))