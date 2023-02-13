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

workbooks = []

def generate_report(idx):
    
    # Create the excel report and adding headers
    #print('Creating the excel report...')
    workbook = xlsxwriter.Workbook(f'./SULLANA_TEST/reportes/{idx}.xlsx')
    
    return [workbook]

def generate_speech(source, sheet, report, campaign):
    # Values that will be used to generate the audio files
    
    users = []
    #print('Reading the excel file...')
    excel = pd.ExcelFile(f'./bd/{source}.xlsx').parse(f'{sheet}')

    #print(excel)

    # Create the excel report and adding headers

    # Extract the names from the excel file and generate the complete speech
    validIdx = -1
    #errorsIdx = -1
    iteration_group = 0

    # get result from the function generate_report()

    
    #global wsh_success
    wsh_successBase = generate_report(iteration_group)[0]
    
    #global wsh_wbok
    #wsh_wbok = generate_report(iteration_group)[1]

    #global wsh_errors
    #wsh_errors = generate_report(iteration_group)[1]

    headers = (excel.columns.values)
    phone_headers = []
    for i in range(0, len(headers)):
        if 'TELEFONO' in headers[i]:
            phone_headers.append(headers[i])

    print(phone_headers)
    counter = 0

    for i in range(0, len(excel)):
        
        #print(f'Audio {counter} de {len(excel)} generado')
        
        worsheet = wsh_successBase.add_worksheet()
        
        counter += 1

        validIdx += 1 

        wsh_success = generate_report(iteration_group)[0]

        if i % 10000 == 0 and i != 0 or i == len(excel) - 1:
            
            # Declarar excels
            iteration_group += 1
            #print('entró')
            wsh_success = generate_report(iteration_group)[0]
            #wsh_wbok = generate_report(iteration_group)[1]
            
            worsheet = wsh_success.add_worksheet()
            worsheet.write('A1', 'ID')
            worsheet.write('B1', 'DNI')
            worsheet.write('C1', 'NÚMERO_TELEFÓNICO')
            worsheet.write('D1', 'SPEECH')
            worsheet.write('E1', 'LONGITUD')
            worsheet.write('F1', 'REDONDEO')
            worsheet.write('G1', 'RESPUESTA')
            
            det_array = []

            for j in range(0, len(users)):
                users[j].create_audio()
                users[j].add_silence()
                users[j].generate_details()
                details = users[j].generate_report() 
                det_array.append(details)
            
            #print('Cerrando el workbook')
            
            #print(det_array)

            for k in range(0, len(det_array)):
                try:
                    worsheet.write(f'A{k+2}', det_array[k][0])
                    worsheet.write(f'B{k+2}', det_array[k][1])
                    worsheet.write(f'C{k+2}', det_array[k][2])
                    worsheet.write(f'D{k+2}', det_array[k][3])
                    worsheet.write(f'E{k+2}', det_array[k][4])
                    worsheet.write(f'F{k+2}', det_array[k][5])
                    worsheet.write(f'G{k+2}', det_array[k][6])
                except:
                    pass
                
            wsh_success.close()
            users = []
            validIdx = 0
            print('---------------')

        #print('acá')    

        for j in range(0, len(phone_headers)):
            uuid_id = str(uuid.uuid4())
            phone = excel[phone_headers[j]][i]
            dni = excel['DNI'][i]
            name = excel['NOMBRE'][i]
            if not name == '-' and len(str(phone)[0:9])>0:
                users.append(Speech( uuid_id, phone, dni, name, engine, idx=validIdx, worksheet=worsheet, campaign=campaign, iteration=iteration_group))

        """
        uuid_id = str(uuid.uuid4())
        phone = excel['TELEFONO 1'][i]
        dni = excel['DNI'][i]
        name = excel['NOMBRE'][i]

        if not name == '-':
            users.append(Speech( uuid_id, phone, dni, name, engine, idx=validIdx, worksheet=worsheet, campaign=campaign, iteration=iteration_group))
        """

    for workbook in workbooks:
        workbook.close()

    print('Done')



#generate_speech(source, sheet, report, campaign)
#generate_speech('mestrias', 'HOJA_FINAL2', 'ReporteMestrias', 'Mestrias')
#generate_speech('esan', 'HOJA_FINAL', 'ReporteESAN', 'Esan')
generate_speech('sullana', 'Hoja1', 'Reporte_Sullana', 'sullana')



