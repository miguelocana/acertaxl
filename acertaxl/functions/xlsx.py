import xlrd, datetime
import pandas as pd
from time import sleep
import numpy as np

# Comprueba si la primera hoja es la de Datos Solicitud
def check_primera_hoja(archivo):
    xl = pd.ExcelFile(archivo)
    res = xl.sheet_names 
    # Devolverá True si Datos está en la primera hoja
    if 'Datos' in res[0]:
        return (True,len(res))
    else:
        return False

def extraccion_datos(indice_hoja,workbook):
    sheet = workbook.sheet_by_index(indice_hoja)
    archivo = 'xlsx'
    if sheet.cell_value(4,4) != '':
        razon = sheet.cell_value(4,4)
    else: 
        razon = np.nan
    if sheet.cell_value(6,4) != '':
        nif = sheet.cell_value(6,4)
    else: 
        nif = np.nan
    if sheet.cell_value(8,4) != '':
        titulo = sheet.cell_value(8,4)
    else: 
        titulo = np.nan
    if sheet.cell_value(11,5) != '':
        a1 = sheet.cell_value(11,5)
        f_inicio = datetime.datetime(*xlrd.xldate_as_tuple(a1,workbook.datemode))
    else: 
        f_inicio = np.nan
    if sheet.cell_value(4,4) != '':
        a2 = sheet.cell_value(11,7)
        f_fin = datetime.datetime(*xlrd.xldate_as_tuple(a2,workbook.datemode))
    else: 
        f_fin = np.nan
    if sheet.cell_value(13,4) != '':
        año = int(sheet.cell_value(13,4))
    else: 
        año = np.nan
    if sheet.cell_value(13,7) != '':
        acronimo = sheet.cell_value(13,7)
    else: 
        acronimo = np.nan
    if sheet.cell_value(15,4) != '':
        expediente = sheet.cell_value(15,4)
    else: 
        expediente = np.nan

    hoja = [archivo,razon,nif,titulo,f_inicio,f_fin,año,acronimo,expediente]
    # Devuelve un dataframe
    return pd.DataFrame([hoja],
                       columns=['Archivo','Razón_Social','NIF','Título','F_Inicio','F_Fin','Año_Inicio','Acrónimo','Expediente'])

# Donde empieza el cuadro de la hoja
def inicio(sheet):
    w = 'Código'
    columna = []
    for i in range(0,100):
        try:
            columna.append(sheet.cell_value(i,1))
        except IndexError:
            break
    for i in columna:
        try:
            if w in i:
                start = (columna.index(i),1)
                break
            else:
                pass
        except TypeError:
            pass
    return start

# Comprueba cuantas filas hay en la hoja
def rows(sheet,start):
    r = start[0] + 2
    codes = []
    i = False
    while i == False:
        if sheet.cell_value(r,1) != '':
            codes.append(sheet.cell_value(r,1))
            r += 1
        else:
            break
    rows = len(codes)
    # Devuelve el número total de filas en la hoja
    return rows

# Comprueba cuantos años hay
def years(sheet,start):
    s = start
    columnas = []
    years = []
    for i in range(0,100):
        try:
            columnas.append(sheet.cell_value(s[0],s[1]+i))
        except IndexError:
            break
    for i in columnas:
        if type(i) == float or type(i) == int:
            years.append((i,columnas.index(i)+1))
    col_recuento = (years[1][1]-years[0][1]) * len(years)
    # Devuelve tuplas de año/posicion y numero de columnas
    return (years,col_recuento)

# Extrae los datos del anexo en un dataframe
def extraccion_gastos(ruta):
    workbook = xlrd.open_workbook(ruta)
    hoja = workbook.sheet_by_index(1)
    a2 = inicio(hoja)
    a3 = rows(hoja,a2)
    a4 = years(hoja,a2)
    personal = []
    inicio_row = a2[0]+2
    for i in range(0,a3):
        x = 0
        y = 5
        for l in range(0,len(a4[0])):
            año = int(a4[0][l][0])      
            persona = []
            for j in range(1,5):
                persona.append(hoja.cell_value(inicio_row,j))
            coste = []
            x += 5
            y += 5    
            for k in range(x,y):
                coste.append(hoja.cell_value(inicio_row,k))        
            personal.append([persona,coste,año])  
        inicio_row += 1

    personalCLEAN = []
    for i in personal:
        person = i[0]+i[1]+[i[2]]
        personalCLEAN.append(person)
    
    personal_df = personal_df = pd.DataFrame(personalCLEAN,columns=['CODIGO','NOMBRE','TITULACION','I+D','HORAS_I+D','COSTE_I+D','HORAS_I','COSTE_I','TOTAL','AÑO'])
    
    return personal_df