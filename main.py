# Import the required libraries
import pyttsx3;
import pandas as pd;
import numpy as np;
from pydub import AudioSegment;
import os
import xlsxwriter
import time

# Setting properties for Pyttsx3 library
engine = pyttsx3.init()

engine.setProperty('rate', 150);
engine.setProperty('volume', 0.9);

engine.setProperty('voice', 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_ES-MX_SABINA_11.0')

# Values that will be used to generate the audio files
speeches = []

excel = pd.ExcelFile('./data300k.xlsx').parse('MENSAJES');

# Extract the names from the excel file and generate the complete speech
for i in range(0, len(excel)):
    name = excel['SPEECH'][i].split(',')[0]
    speeches.append(f'Hola, ¿Me comunico con {name}?')

# Array that will contain the details of the audio files
details = []

# Create the audios and the report
start = time.perf_counter()

for idx, speech in enumerate(speeches[:5000]):
    # Creating the base audios
    print(f'Processing {idx} of {len(speeches[:5000])}')
    engine.save_to_file(speech, f'./tmp/audio_{idx}.wav')
    engine.runAndWait()
    #  Adding silence to the audios
    audio = AudioSegment.from_wav(f'./tmp/audio_{idx}.wav')
    audio_silent = audio.append(audio.silent(duration=5000))
    # Required details in the details array
    audio_length = round(audio_silent.duration_seconds, 2)
    audio_rounded = round(audio_silent.duration_seconds)
    audio_silent.export(f'./audios/audio_{idx}.wav', format='wav')
    # Remove the temporal audio
    os.remove(f'./tmp/audio_{idx}.wav')
    details.append([idx, speech, audio_length, audio_rounded])

# Create the excel report and adding headers
workbook = xlsxwriter.Workbook('./excels/report.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'ID')
worksheet.write('B1', 'SPEECH')
worksheet.write('C1', 'LENGTH')
worksheet.write('D1', 'ROUNDED')

# Fill the excel report with the array
for idx, detail in enumerate(details):
    worksheet.write(idx+1, 0, detail[0])
    worksheet.write(idx+1, 1, detail[1])
    worksheet.write(idx+1, 2, detail[2])
    worksheet.write(idx+1, 3, detail[3])

workbook.close()

end = time.perf_counter()
print(f'Finished in {round(end-start, 2)} seconds')



