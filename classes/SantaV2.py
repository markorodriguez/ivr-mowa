# Import the required libraries
import pandas as pd;
from pydub import AudioSegment;
import os

class Speech:
    def __init__(self, id_audio, phone, dni, name, engine, idx, worksheet, campaign, iteration):
        self.id_audio = id_audio
        self.phone = phone
        self.name = name
        self.dni = dni
        self.speech = f'Hola {self.name}'
        self.engine = engine
        self.details = []
        self.idx = idx
        self.worksheet = worksheet
        self.audio_silent = " "
        self.campaign = campaign
        self.iteration = iteration
        

    def create_audio(self):
        os.makedirs(f'./tmp/{self.campaign}', exist_ok=True)
        self.engine.save_to_file(self.speech, f'./tmp/{self.campaign}/{self.id_audio}.wav')
        self.engine.runAndWait()

    def generate_details(self):
        audio_length = round(self.audio_silent.duration_seconds, 2)
        audio_rounded = round(self.audio_silent.duration_seconds)
        self.details = [self.id_audio, self.dni, self.phone, self.speech, audio_length, audio_rounded, ' ']

    def generate_report(self):

        """
        self.worksheet.write(self.idx+1, 0, self.details[0])
        self.worksheet.write(self.idx+1, 1, self.details[1])
        self.worksheet.write(self.idx+1, 2, self.details[2])
        self.worksheet.write(self.idx+1, 3, self.details[3])
        self.worksheet.write(self.idx+1, 4, self.details[4])
        self.worksheet.write(self.idx+1, 5, self.details[5])
        self.worksheet.write(self.idx+1, 6, self.details[6])
        """

        return self.details

        
        #write_row

    # TODO: CAMBIAR EL NOMBRE DE LA CARPETA DE AUDIOS TODO LO QUE DIGA SANTA-V2
    def add_silence(self):
        audio = AudioSegment.from_wav(f'./tmp/{self.campaign}/{self.id_audio}.wav')
        self.audio_silent = audio.append(audio.silent(duration=5000))

        #os.makedirs(f'./audios/{self.campaign}', exist_ok=True)
        os.makedirs(f'./SANTA-V2/audios/{self.iteration}', exist_ok=True)
        self.audio_silent.export(f'./SANTA-V2/audios/{self.iteration}/{self.id_audio}.wav', format='wav')
        os.remove(f'./tmp/{self.campaign}/{self.id_audio}.wav')
    
    def set_error(self, reason):
        details_error = [self.id_audio, self.dni, self.name, reason]
        self.worksheet.write(self.idx+1, 0, details_error[0])
        self.worksheet.write(self.idx+1, 1, details_error[1])
        self.worksheet.write(self.idx+1, 2, details_error[2])
        self.worksheet.write(self.idx+1, 3, details_error[3])
    
    # Generate error report
    
    
