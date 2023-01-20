import threading
import pyttsx3
import pandas as pd
import numpy as np
from pydub import AudioSegment
import os
import xlsxwriter
import time
import uuid

# Setting properties for Pyttsx3 library
engine = pyttsx3.init()

engine.setProperty('rate', 150)
engine.setProperty('volume', 1)

engine.setProperty(
    'voice', 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_ES-MX_SABINA_11.0')

# Values that will be used to generate the audio files
speeches = []

# Array that will contain the details of the audio files
details = []

threads = []

chunks = []

print('Reading the excel file...')
excel = pd.ExcelFile('./sources/esan.xlsx', engine='openpyxl').parse('HOJA_FINAL')

validIdx = -1
errorsIdx = -1

print('Creating the excel report...')
workbook = xlsxwriter.Workbook(f'./reportes/hilo_success.xlsx')
worksheet = workbook.add_worksheet('REPORTE')
worksheet.write('A1', 'ID')
worksheet.write('B1', 'DNI')
worksheet.write('C1', 'NÚMERO_TELEFÓNICO')
worksheet.write('D1', 'SPEECH')
worksheet.write('E1', 'LONGITUD')
worksheet.write('F1', 'REDONDEO')
worksheet.write('G1', 'RESPUESTA')


print('Creating the excel report for the errors...')
workbookError = xlsxwriter.Workbook(f'./reportes/hilo_error.xlsx')
worksheetErrores = workbookError.add_worksheet('ERRORES')
worksheetErrores.write('A1', 'ID')
worksheetErrores.write('B1', 'DNI')
worksheetErrores.write('C1', 'NÚMERO_TELEFÓNICO')
worksheetErrores.write('D1', 'NOMBRE_COMPLETO')
worksheetErrores.write('E1', 'MOTIVO')

# Generate chunks of 10000 rows, asign them to a thread and extract name for every row in the chun


def generate_speech(chunk):
    global validIdx
    global errorsIdx
    global details

    uui_array = []

    # CREATE AUDIOS
    for i in range(0, len(chunk)):
        uuid_id = str(uuid.uuid4())

        phone = chunk[i][2]
        # We need a uniform way for this field since it's really important
        dni = chunk[i][1]
        name = chunk[i][3]
        name_parsed = name.split(' ')
        # Check errors
        if phone == 0 or name.__contains__('No se encontró') or len(name_parsed) == 0 or len(name_parsed) == 1:
            ++ errorsIdx
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
            valid_name = name_parsed[len(name_parsed)-1]
            ++ validIdx
            if len(str(dni)) == 7:
                dni_fixed = '0' + str(dni)
                speech = f'Hola, {valid_name}'
                engine.save_to_file(speech, f'./hilos/tmp/{uuid_id}.wav')
                engine.runAndWait()
                """
                time.sleep(0.5)
                audio = AudioSegment.from_wav(f'./hilos/tmp/{uuid_id}.wav')
                audio_silence = audio + audio.silent(duration=5000)
                audio_silence.export(f'./hilos/audios/{uuid_id}.wav', format='wav')
                #os.remove(f'./hilos/tmp/{uuid_id}.wav')
                audio_lenght = round(audio_silence.duration_seconds, 2)
                audio_rounded = round(audio_silence.duration_seconds)

                details.append([validIdx ,uuid_id, dni_fixed, phone, speech, audio_lenght, audio_rounded, ''])
                """
                speeches.append(speech)
                #print(speech)
            else:
                speech = f'Hola, {valid_name}'
                engine.save_to_file(speech, f'./hilos/tmp/{uuid_id}.wav')
                engine.runAndWait()
                """
                time.sleep(0.5)
                audio = AudioSegment.from_wav(f'./hilos/tmp/{uuid_id}.wav')
                audio_silence = audio + audio.silent(duration=5000)	
                audio_silence.export(f'./hilos/audios/{uuid_id}.wav', format='wav')
                #os.remove(f'./hilos/tmp/{uuid_id}.wav')
                audio_lenght = round(audio_silence.duration_seconds, 2)
                audio_rounded = round(audio_silence.duration_seconds)

                details.append([validIdx, uuid_id, dni, phone, speech, audio_lenght, audio_rounded, ''])
                """
                #print(speech)
                speeches.append(speech)
    #"""
    # ADD SILENCE

    for details in details:
        worksheet.write(details[0]+1, 0, details[1])
        worksheet.write(details[0]+1, 1, details[2])
        worksheet.write(details[0]+1, 2, details[3])
        worksheet.write(details[0]+1, 3, details[4])
        worksheet.write(details[0]+1, 4, details[5])
        worksheet.write(details[0]+1, 5, details[6])
        worksheet.write(details[0]+1, 6, details[7])
    #"""
        


# Split excel in chunks of 10000 rows and create new arrays
for i in range(0, len(excel), 10000):
    excel_chunk = excel[i:i+10000]
    # convert the dataframe to a numpy array
    print(excel_chunk)
    print(type(excel_chunk))
    array = excel_chunk.to_numpy()

    chunks.append(array)

# Create a thread for each chunk
for chunk in chunks:
    t = threading.Thread(target=generate_speech, args=(chunk,))
    t.start()
    threads.append(t)

# Wait for all threads to finish
for thread in threads:
    thread.join()

# Save the excel file
print('Saving excel file...')
print(len(speeches))
workbook.close()
workbookError.close()






"""
# For each row in the excel file, extract the name and generate the complete speech
def generate_speech(idx):
    name = excel['SPEECH'][idx].split(',')[0]
    speeches.append(f'Hola, ¿Me comunico con {name}?')

threads = []

# Split excel in chunks of 10000 rows and create a thread for each chunk
for i in range(0, len(excel), 10000):
    threads.append(threading.Thread(target=generate_speech, args=(i,)))
    
for _ in range(10):
    # [10K]
    # Recorrerer el excel
    t = threading.Thread(target=generate_speech, args=())
    t.start()
    threads.append(t)

for thread in threads:
    thread.join()

# Dividir el longitud en bloques de 10K, sobrar (los itero)
# 10 k a 1 hilo

for i in range(0, len(excel), 10):
    name = excel['SPEECH'][i].split(',')[0]
    speeches.append(f'Hola, ¿Me comunico con {name}?')
    
"""
