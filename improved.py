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
    audios = os.listdir('./IVR/296320')
    
    workbook = xlsxwriter.Workbook('./comparacion.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.write('A1', 'Audio ID')
    worksheet.write('B1', 'Google')
    worksheet.write('C1', 'Whisper BASE')
    worksheet.write('D1', 'Whisper SMALL')

    for idx, audio in enumerate(audios):
        compare_models(audio, idx+1, worksheet)
    workbook.close()


def compare_models(audio_path, idx, worksheet):
    with sr.AudioFile(f'./IVR/296320/{audio_path}') as source:
        recognizer.adjust_for_ambient_noise(source, duration=0.5)
        audio = recognizer.record(source)
        print(audio_path)
        try:
            text_google = recognizer.recognize_google(audio, language='es-ES')
            result = model.transcribe(f'./IVR/296320/{audio_path}', language='es')
            result_base = model_base.transcribe(f'./IVR/296320/{audio_path}', language='es')
            print(result['text'])
            print(result_base['text'])
            worksheet.write(idx, 0, audio_path)
            worksheet.write(idx, 1, text_google)
            worksheet.write(idx, 2, result_base['text'])
            worksheet.write(idx, 3, result['text'])
        except:
            print('error')

    print('---------------------')


get_audios_file()
