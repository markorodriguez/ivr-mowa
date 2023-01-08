# Import the required libraries
import pandas as pd;
from pydub import AudioSegment;
import os

class Speech:
    def __init__(self, id_audio, dni, name, engine, idx, worksheet, campaign):
        self.id_audio = id_audio
        self.name = name
        self.dni = dni
        self.speech = f'Hola, ¿Me comunico con {self.name}?'
        self.engine = engine
        self.details = []
        self.idx = idx
        self.worksheet = worksheet
        self.audio_silent = " "
        self.campaign = campaign

    def create_audio(self):
        os.makedirs(f'./tmp/{self.campaign}', exist_ok=True)
        self.engine.save_to_file(self.speech, f'./tmp/{self.campaign}/{self.id_audio}_{self.dni}.wav')
        self.engine.runAndWait()

    def generate_details(self):
        audio_length = round(self.audio_silent.duration_seconds, 2)
        audio_rounded = round(self.audio_silent.duration_seconds)
        self.details = [self.id_audio, self.dni, self.speech, audio_length, audio_rounded, ' ']

    def generate_report(self):
        self.worksheet.write(self.idx+1, 0, str(self.details[0]))
        self.worksheet.write(self.idx+1, 1, self.details[1])
        self.worksheet.write(self.idx+1, 2, self.details[2])
        self.worksheet.write(self.idx+1, 3, self.details[3])
        self.worksheet.write(self.idx+1, 4, self.details[4])
        self.worksheet.write(self.idx+1, 5, self.details[5])

    def add_silence(self):
        
        audio = AudioSegment.from_wav(f'./tmp/{self.campaign}/{self.id_audio}_{self.dni}.wav')
        self.audio_silent = audio.append(audio.silent(duration=5000))
        os.makedirs(f'./audios/{self.campaign}', exist_ok=True)
        self.audio_silent.export(f'./audios/{self.campaign}/{self.id_audio}_{self.dni}.wav', format='wav')
        os.remove(f'./tmp/{self.campaign}/{self.id_audio}_{self.dni}.wav')
    
    def set_error(self, reason):
        details_error = [self.id_audio, self.dni, self.name, reason]
        self.worksheet.write(self.idx+1, 0, details_error[0])
        self.worksheet.write(self.idx+1, 1, details_error[1])
        self.worksheet.write(self.idx+1, 2, details_error[2])
        self.worksheet.write(self.idx+1, 3, details_error[3])
        
    
    
