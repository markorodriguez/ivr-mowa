import speech_recognition as sr
import os
from pydub import AudioSegment
from pydub.playback import play
import pyttsx3
import time
import pywhisper as whisper
import xlsxwriter
import pandas as pd

engine = pyttsx3.init()

engine.setProperty('rate', 165)
engine.setProperty('volume', 1)

engine.setProperty(
    'voice', 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_ES-MX_SABINA_11.0')

recognizer = sr.Recognizer()
# 400 is the default energy threshold. Try adjusting this value based on your background noise
recognizer.energy_threshold = 500

model = whisper.load_model('small')
model_base = whisper.load_model('base')
model_medium = whisper.load_model('medium')


def improve_audio():
    # audio_start = path.join(path.dirname(path.realpath(__file__)), 'audios/Mestrias/0bfee943-70b1-4910-b0f7-dc853fa73408.wav')
    # audio_clean = recognizer.adjust_for_ambient_noise(audio_start, duration=1)
    # AudioSegment.export(audio_clean, './audios/rpta/lidiaclean.wav', format='wav')
    base_audio = AudioSegment.from_wav('./audios/rpta/lidia3.wav')
    audio_improved = base_audio + 30
    audio_improved.export('./audios/rpta/lidia3.wav', format='wav')
    transpile_audio()


def transpile_audio():
    with sr.AudioFile('./audios/rpta/lidia3.wav') as source:
        recognizer.adjust_for_ambient_noise(source, duration=1)
        audio = recognizer.record(source)
        try:
            text = recognizer.recognize_vosk(audio, language='es-MX')
            # text = r.recognize_google(audio, language='es-ES')
            print(text)
        except:
            print('Error')


def get_audios_file():
    # GET ALL THE AUDIOS FROM IVR/296320 wITH THE SAME NAME
    audios = os.listdir('./IVR/voice')

    workbook = xlsxwriter.Workbook('./comparacion.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.write('A1', 'Audio ID')
    worksheet.write('B1', 'Google')
    worksheet.write('C1', 'Tiempo')
    worksheet.write('D1', 'Whisper BASE')
    worksheet.write('E1', 'Tiempo')
    worksheet.write('F1', 'Whisper SMALL')
    worksheet.write('G1', 'Tiempo')
    worksheet.write('H1', 'Whisper MEDIUM')
    worksheet.write('I1', 'Tiempo')

    idx = 0

    for i in range(len(audios)):
        with sr.AudioFile(f'./IVR/voice/{audios[i]}') as source:
            recognizer.adjust_for_ambient_noise(source, duration=0.5)
            audio = recognizer.record(source)
            print('audio', audios[i])
            try:
                google_start = time.perf_counter()
                text_google = recognizer.recognize_google(
                    audio, language='es-ES')
                google_end = time.perf_counter()

                base_start = time.perf_counter()
                result_base = model_base.transcribe(
                    f'./IVR/voice/{audios[i]}', language='es')
                base_end = time.perf_counter()

                small_start = time.perf_counter()
                result = model.transcribe(
                    f'./IVR/voice/{audios[i]}', language='es')
                small_end = time.perf_counter()

                medium_start = time.perf_counter()
                result_medium = model_medium.transcribe(
                    f'./IVR/voice/{audios[i]}', language='es')
                medium_end = time.perf_counter()

                idx += 1
                print(result['text'])
                print(result_base['text'])
                worksheet.write(idx, 0, audios[i])
                worksheet.write(idx, 1, text_google)
                worksheet.write(idx, 2, str(
                    round(google_end - google_start, 2)))
                worksheet.write(idx, 3, result_base['text'])
                worksheet.write(idx, 4, str(
                    round(base_end - base_start, 2)))
                worksheet.write(idx, 5, result['text'])
                worksheet.write(idx, 6, str(
                    round(small_end - small_start, 2)))
                worksheet.write(idx, 7, result_medium['text'])
                worksheet.write(idx, 8, str(
                    round(medium_end - medium_start , 2)))

            except:
                print('error')
    workbook.close()


get_audios_file()
