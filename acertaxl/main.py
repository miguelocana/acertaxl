from functions.xlsx import check_primera_hoja, years, extraccion_datos, inicio, rows, extraccion_gastos
import xlrd, datetime
import pandas as pd
from time import sleep
import numpy as np
from sqlalchemy import create_engine
engine = create_engine('sqlite://', echo=False)

print('ACERTAxl: extracción y validación de datos del ANEXO II.')
print('-'*40)

file = input('Introduce la ruta: ')
i = False
while i == False:
    try:
        workbook = xlrd.open_workbook(file)
        a1 = check_primera_hoja(file)
        if a1[0] == True:
            for i in '.'*3:
                print(i)
                sleep(0.3)
            print('Archivo subido correctamente.')
            sleep(2)
            print('Evaluando...')
            for i in '.'*3:
                print(i)
                sleep(0.3)
            print('Portada: OK')
            for i in '.'*3:
                print(i)
                sleep(0.3)
            print('Gastos: OK')
            sleep(2)
            print('Extrayendo datos de solicitud...')
            for i in '.'*3:
                print(i)
                sleep(0.3)
            datos_solicitud = extraccion_datos(0,workbook)
            print('Completado!')
            sleep(2)
            
            print('Extrayendo gastos...')
            for i in '.'*3:
                print(i)
                sleep(0.3)
            gastos_solicitud = extraccion_gastos(file)
            print('Completado!')
            
            sleep(1)
            print('-'*20)
            print('¿En qué formato quieres la extracción?')
            print('\t1 Excel')
            print('\t2 CSV')
            print('\t3 SQL')
            o = int(input('Elige una opción: '))
            
            if o == 1:
                print('Extracción completada! (.xlsx)')
                datos_solicitud.to_excel('datos_solicitud.xlsx',index=False)
                gastos_solicitud.to_excel('gastos_anexo.xlsx',index=False)
            elif o == 2:
                print('Extracción completada! (.csv)')
                datos_solicitud.to_csv('datos_solicitud',index=False)
                gastos_solicitud.to_csv('gastos_anexo',index=False)
            elif o == 3:
                print('Extracción completada! (.sql)')
                datos_solicitud.to_sql('datos_solicitud',con=engine)
                gastos_solicitud.to_sql('gastos_anexo',con=engine)
            else:
                pass
            
    except OSError:
        print('Ruta no válida.')
        break  