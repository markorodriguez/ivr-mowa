# Import the required libraries
import pyttsx3
import pandas as pd
import numpy as np
import os
import xlsxwriter
import time
import uuid
from classes.Speech import Speech

# Setting properties for Pyttsx3 library
engine = pyttsx3.init()

engine.setProperty('rate', 165)
engine.setProperty('volume', 1.5)

engine.setProperty(
    'voice', 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_ES-MX_SABINA_11.0')

#print('Mueva el archivo a la raíz del proyecto')
#campaign = input('Ingrese el nombre o ID de la campaña: ')
#source = input('Ingrese el nombre del archivo a cargar: ')
#sheet = input('Ingrese el nombre de la hoja con los datos: ')
#report = input('Ingrese el nombre del reporte a generar: ')

def generate_speech(source, sheet, report, campaign):
    # Values that will be used to generate the audio files
    start = time.perf_counter()
    users = []
    print('Reading the excel file...')
    excel = pd.ExcelFile(f'./sources/{source}.xlsx').parse(f'{sheet}')

    # Create the excel report and adding headers
    print('Creating the excel report...')
    workbook = xlsxwriter.Workbook(f'./reportes/{campaign}/{report}.xlsx')
    worksheet = workbook.add_worksheet('REPORTE')
    worksheet.write('A1', 'ID')
    worksheet.write('B1', 'DNI')
    worksheet.write('C1', 'NÚMERO_TELEFÓNICO')
    worksheet.write('D1', 'SPEECH')
    worksheet.write('E1', 'LONGITUD')
    worksheet.write('F1', 'REDONDEO')
    worksheet.write('G1', 'RESPUESTA')


    # Create the excel report for the errors
    os.makedirs(f'./reportes/{campaign}', exist_ok=True)
    print('Creating the excel report for the errors...')
    workbookError = xlsxwriter.Workbook(f'./reportes/{campaign}/errores_{report}.xlsx')
    worksheetErrores = workbookError.add_worksheet('ERRORES')
    worksheetErrores.write('A1', 'ID')
    worksheetErrores.write('B1', 'DNI')
    worksheetErrores.write('C1', 'NÚMERO_TELEFÓNICO')
    worksheetErrores.write('D1', 'NOMBRE_COMPLETO')
    worksheetErrores.write('E1', 'MOTIVO')

    # Extract the names from the excel file and generate the complete speech
    validIdx = -1
    errorsIdx = -1
    for i in range(0, len(excel[:1300])):
        print('Procesando usuario', i+1, 'de', len(excel[:1000]), '...')
        uuid_id = str(uuid.uuid4())
        phone = excel['TELEFONO_ALL'][i]
        # We need a uniform way for this field since it's really important
        dni = excel['DOCUMENTO'][i]
        name = excel['NOMBRE_B'][i]
        name_parsed = name.split(' ')
        # Check errors
        if phone == 0 or name.__contains__('No se encontró') or len(name_parsed) == 0 or len(name_parsed) == 1:
            errorsIdx += 1
            worksheetErrores.write(errorsIdx+1, 0, uuid_id)
            worksheetErrores.write(errorsIdx+1, 1, dni)
            worksheetErrores.write(errorsIdx+1, 2, phone)
            worksheetErrores.write(errorsIdx+1, 3, name)
            # Specify the error type
            if phone == 0:

                worksheetErrores.write(errorsIdx+1, 4, 'Número de teléfono no válido')
            elif name.__contains__('No se encontró'):

                worksheetErrores.write(errorsIdx+1, 4, 'No se encontró el nombre')
            elif len(name_parsed) == 0 or len(name_parsed) == 1:

                worksheetErrores.write(errorsIdx+1, 4, 'El nombre no es válido')
        else:
            validIdx += 1
            valid_name = name_parsed[len(name_parsed)-1]
            if len(str(dni)) == 7:
                dni_fixed = '0' + str(dni)
                users.append(Speech(uuid_id, phone, dni_fixed, valid_name, engine, idx=validIdx, worksheet=worksheet, campaign=campaign))
            else:
                users.append(Speech( uuid_id, phone, dni, valid_name, engine, idx=validIdx, worksheet=worksheet, campaign=campaign))

    for user in users:
        user.create_audio()
        user.add_silence()
        user.generate_details()
        user.generate_report()
        
    workbook.close()
    workbookError.close()

    print('Tiempo de ejecución:', round((time.perf_counter() - start)/60, 2), 'minutos')
    print('Se trabajó correctamente con:', len(users), 'usuarios')
    print('Se encontraron', len(excel[:1300]) - len(users), 'errores')
    print('Done')


#generate_speech(source, sheet, report, campaign)
#generate_speech('mestrias', 'HOJA_FINAL2', 'ReporteMestrias', 'Mestrias')
#generate_speech('esan', 'HOJA_FINAL', 'ReporteESAN', 'Esan')
generate_speech('custom1101', 'HOJA_FINAL2', 'Reporte_Prueba_11012023', 'R_11012023')
def find_unique():
    values = []
    excel = pd.ExcelFile(f'./sources/mestrias.xlsx').parse(f'HOJA_FINAL2')
    for i in range(0, len(excel[:100000])):
        id_audio = excel['TELEFONO_ALL'][i]
        name = excel['NOMBRE_B'][i]
        dni = excel['DOCUMENTO'][i]
        if id_audio != 0 and not name.__contains__('No se encontró') and len(name.split(' ')) > 1:
            values.append(str(id_audio) + str(dni))
    
    print(len(values), 'TOTAL REGISTROS VALIDOS')
    print(len(set(values)), 'TOTAL REGISTROS UNICOS')
    print(len(values) - len(set(values)), 'TOTAL REGISTROS DUPLICADOS')

#find_unique()

def desired_quantity(source, sheet, report, campaign, desired_quantity):
    # Values that will be used to generate the audio files
    start = time.perf_counter()
    users = []
    print('Reading the excel file...')
    excel = pd.ExcelFile(f'./sources/{source}.xlsx').parse(f'{sheet}')

    # Create the excel report and adding headers
    print('Creating the excel report...')
    workbook = xlsxwriter.Workbook(f'./reportes/{campaign}/{report}.xlsx')
    worksheet = workbook.add_worksheet('REPORTE')
    worksheet.write('A1', 'ID')
    worksheet.write('B1', 'DNI')
    worksheet.write('C1', 'NÚMERO_TELEFÓNICO')
    worksheet.write('D1', 'SPEECH')
    worksheet.write('E1', 'LONGITUD')
    worksheet.write('F1', 'REDONDEO')
    worksheet.write('G1', 'RESPUESTA')


    # Create the excel report for the errors
    os.makedirs(f'./reportes/{campaign}', exist_ok=True)
    print('Creating the excel report for the errors...')
    workbookError = xlsxwriter.Workbook(f'./reportes/{campaign}/errores_{report}.xlsx')
    worksheetErrores = workbookError.add_worksheet('ERRORES')
    worksheetErrores.write('A1', 'ID')
    worksheetErrores.write('B1', 'DNI')
    worksheetErrores.write('C1', 'NÚMERO_TELEFÓNICO')
    worksheetErrores.write('D1', 'NOMBRE_COMPLETO')
    worksheetErrores.write('E1', 'MOTIVO')

    # Extract the names from the excel file and generate the complete speech
    validIdx = -1
    errorsIdx = -1
    for i in range(0, len(excel[:1300])):
        print('Procesando usuario', i+1, 'de', len(excel[:1000]), '...')
        uuid_id = str(uuid.uuid4())
        phone = excel['TELEFONO_ALL'][i]
        # We need a uniform way for this field since it's really important
        dni = excel['DOCUMENTO'][i]
        name = excel['NOMBRE_B'][i]
        name_parsed = name.split(' ')
        # Check errors
        if phone == 0 or name.__contains__('No se encontró') or len(name_parsed) == 0 or len(name_parsed) == 1:
            errorsIdx += 1
            worksheetErrores.write(errorsIdx+1, 0, uuid_id)
            worksheetErrores.write(errorsIdx+1, 1, dni)
            worksheetErrores.write(errorsIdx+1, 2, phone)
            worksheetErrores.write(errorsIdx+1, 3, name)
            # Specify the error type
            if phone == 0:

                worksheetErrores.write(errorsIdx+1, 4, 'Número de teléfono no válido')
            elif name.__contains__('No se encontró'):

                worksheetErrores.write(errorsIdx+1, 4, 'No se encontró el nombre')
            elif len(name_parsed) == 0 or len(name_parsed) == 1:

                worksheetErrores.write(errorsIdx+1, 4, 'El nombre no es válido')
        else:
            validIdx += 1
            valid_name = name_parsed[len(name_parsed)-1]
            if len(str(dni)) == 7:
                dni_fixed = '0' + str(dni)
                users.append(Speech(uuid_id, phone, dni_fixed, valid_name, engine, idx=validIdx, worksheet=worksheet, campaign=campaign))
            else:
                users.append(Speech( uuid_id, phone, dni, valid_name, engine, idx=validIdx, worksheet=worksheet, campaign=campaign))

    for user in users:
        user.create_audio()
        user.add_silence()
        user.generate_details()
        user.generate_report()
        
    workbook.close()
    workbookError.close()

    print('Tiempo de ejecución:', round((time.perf_counter() - start)/60, 2), 'minutos')
    print('Se trabajó correctamente con:', len(users), 'usuarios')
    print('Se encontraron', len(excel[:1000]) - len(users), 'errores')
    print('Done')

def write_errors():
    print('errors')

def write_success():
    print('success')