import pywhisper as whisper

model = whisper.load_model('small')
result = model.transcribe('./lidia.wav', language='es', fp16=False)
print(result)