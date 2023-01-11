import speech_recognition as sr
import os.path
import psycopg2
from datetime import datetime
import wave
import contextlib
import logging
import time
import glob
import unicodedata
from sys import exit
from datetime import date

# Whisper dependencies
import pywhisper as whisper
model = whisper.load_model('base')

def getCurrentTime():
	localtime = time.localtime()
	currentTime = time.strftime("%d-%m-%Y %I:%M:%S %p", localtime)
	return currentTime

def getConnection():
	print(" Obteniendo conexion BD: " + getCurrentTime())
	try:
		connection = psycopg2.connect(
			user = "postgres",
			password = "admin",
			host = "localhost",
			port = "5432",
			database = "localDB"
		)
		print("Conexion Ok")
		return connection
	except (psycopg2.Error) as error :
		print ("Error obteniendo conexion PostgreSQL ", error)

def createFileTxt(_strToSave, _campaign):
	try:
		#d2 = today.strftime("%Y_%m_%d") 
		d2 = str(_campaign)
		pathVoice = '/home/centos/speech-text/'
		fileNamex = pathVoice + 'text_file_'+ d2 + ".txt"
		f = open(fileNamex, "a") #open with append in file
		f.write('{}'.format(_strToSave))
		f.close()

	except Exception as error:
		print('error to create file', error)


def getCampaignsOfDay():
	print(" Obteniendo campanhas del dia, " + getCurrentTime())
	records = []
	try:
		querySQL = """ SELECT id, corporate_id FROM usrsms.sendings_voice WHERE (scheduled_start_date::date = CURRENT_DATE) AND process_status = 7 
		AND id in (SELECT campaign_id FROM usrsms.cdr_info WHERE creation_date::date = CURRENT_DATE GROUP BY campaign_id)"""
		connection = getConnection()
		cursor = connection.cursor()
		cursor.execute(querySQL)
		campaigns = cursor.fetchall()
		print("Campanhas obtenidas: " + str(len(campaigns)))
		for column in campaigns:
			print("Campanha: " + str(column[0]) + ", usuario: " + column[1])
			records.append(column[0])
	except Exception as error:
		print('Error obteniendo campanhas: ', error)
	finally:
		if connection:
			connection.close()
			cursor.close()
			print("Connection Closed")

	return records

def getCdrInfoByCampaign(_id):
	print(" Obteniendo datos de cdr_info, " + getCurrentTime())
	records = []
	try:
		querySQL = " SELECT unique_id, status_text FROM usrsms.cdr_info WHERE campaign_id = " + str(_id)
		connection = getConnection()
		cursor = connection.cursor()
		cursor.execute(querySQL)
		records = cursor.fetchall()
		print("Registros obtenidos: " + str(len(records)))
	except Exception as error:
		print('Error obteniendo cdrInfo: ', error)
	finally:
		if connection:
			connection.close()
			cursor.close()
			print("Connection Closed")

	return records


localDir = '/var/spool/asterisk/monitor/Grabaciones/' #'/var/spool/asterisk/monitor/Grabaciones/2020/11/04/'
#localDir = '/Diego/IVRTEST/294793/'
r = sr.Recognizer()
today = date.today()
d1 = today.strftime("%Y/%m/%d/")
#d1 = 'file'+ str(184887) + '/' #agregado para campanhas de innova 24/06
localPath = localDir + d1
print("directory: ", localPath)

#lista de archivos carpeta
#listOfFiles = [f for f in os.listdir(localPath)]

def speechToText(_dictionary, _campaign):
	listOfFiles = [os.path.basename(x) for x in glob.glob(localPath + "*--" + str(_campaign) +"-*.wav")]
	#print(listOfFiles)
	i_0 = 0
	duration = 0.999
	_str = ''
	keytUniqueIdList = _dictionary.keys()
	#print(keytUniqueIdList)
	allDataToUpdate = []
	try:
		for f_index in listOfFiles:
			file_name = f_index
			subName = f_index.split('-')

			msisdn = subName[0] #f_index[:i_1]
			date_x = subName[4] #f_index[(i_1+2):(i_2-1)] #-4 -> .wav
			uniq_id = subName[3]
			temp_file = localPath + f_index

			if uniq_id in keytUniqueIdList:
				speechText = '-'
				statusText = _dictionary[uniq_id]
				_size = os.path.getsize(temp_file)
				fname = temp_file

				with contextlib.closing(wave.open(fname,'r')) as f:
					frames = f.getnframes()
					rate = f.getframerate()
					duration = frames / float(rate)

				if _size < 50:
					_str = _str + msisdn + '|' + date_x + '|' + '-|'+ str(duration) + '|' + str(_size) + '|' + temp_file
					if  statusText == 'RECHAZADO_O_CONT_ESCUCHADO_TODO':
						statusText = 'RECHAZADO_O_CONT_ESCUCHADO_TODO - RECHAZADO'
				else:
					#
					audiosr = sr.AudioFile(temp_file)	
					with audiosr as source:
						audio = r.record(source)

					try:
						#text = r.recognize_google(audio, language="es-ES",show_all=False)
						# Whisper recognition
						text = model.transcribe(f'{localPath}/{f_index}', language='es')
						
						_str = _str + msisdn + '|' + date_x + '|' + '{}'.format(text)  +'|'+ str(duration) + '|' + str(_size) + '|' + temp_file
						#speechText = text
						# Whisper speech
						whisper_speech = text

						if statusText == 'RECHAZADO_O_CONT_ESCUCHADO_TODO':
							if  'mensaje' in '{}'.format(text) or  'grabe' in '{}'.format(text) or 'tono' in '{}'.format(text) or 'buzon' in '{}'.format(text) or 'buzón' in '{}'.format(text):
								statusText = 'RECHAZADO_O_CONT_ESCUCHADO_TODO - RECHAZADO'
							else:
								if(duration < 0.3 and _size < 4000):
									statusText = 'RECHAZADO_O_CONT_ESCUCHADO_TODO - RECHAZADO'
								else:
									statusText = 'RECHAZADO_O_CONT_ESCUCHADO_TODO - CONTESTADO'
						
					except Exception as ex:
						print('Error recognize')
						print(ex)
						_str = _str + msisdn + '|' + date_x + '|' + '-x-|'+ str(duration) + '|' + str(_size) + '|' + temp_file
						whisper_speech = whisper_speech if len(whisper_speech) > 1 else '-x-'

				items = [uniq_id, statusText, whisper_speech]			
				allDataToUpdate.append(items)

				_str += '\n'	
				i_0 = i_0 + 1
				if (i_0%50) == 0:
					now = datetime.now()
					print('now = ', now, ' audio: ', i_0)
					#print("now = " + now + ', audio' + str(i_0))

		createFileTxt(_str, _campaign)
	except Exception as error:
		print('error to read files: ', error)

	return allDataToUpdate

def updateCdrInfo(_dataToUpdate, _id):
        try:
                print("Intentando actualizar: " + str(_id) + ", con: " + str(len(_dataToUpdate)) + " datos  ")
                tableTmp = ""
                idx = 0
                updateIdx = 0
                roundTime = 1
                connection = getConnection()
                cursor = connection.cursor()

                for data in _dataToUpdate:
                        if idx == 0:
                                tableTmp = "SELECT '" + data[0] + "' as uniqueId, '" + data[1] + "' as statusText, '" + data[2] + "' as speechText"
                        else:
                                tableTmp = tableTmp + " UNION SELECT '" + data[0] + "', '" + data[1] + "', '" + data[2] + "' "
                        idx += 1
                        if idx % 1000 == 0 or (updateIdx + idx) == len(_dataToUpdate):
                                updateIdx += idx
                                idx = 0
                                querySQL = """ UPDATE usrsms.cdr_info cdr SET status_text = tb.statusText, speech_text = tb.speechText
                                                        FROM ( """ + tableTmp + """ ) as tb (uniqueId, statusText, speechText)
                                                        WHERE campaign_id = """ + str(_id) + """ AND unique_id = tb.uniqueId """
                                cursor.execute(querySQL)
                                connection.commit()
                                count = cursor.rowcount
                                print(count, " records Updated successfully, round: ", roundTime)
                                roundTime += 1
        except Exception as error:
                logging.error("Error actualizando cdr_info ", _id, " en round: ", roundTime)
                print('Error actualizando cdr_info', error)
        finally:
                if connection:
                        connection.close()
                        cursor.close()
                        print("Connection Closed")

def speechToTextByCampaign(_id):
	myDictionary = {}
	cdrInfoRecords = getCdrInfoByCampaign(_id)
	if len(cdrInfoRecords) > 0:
		for cdrInfo in cdrInfoRecords:
			myDictionary[str(cdrInfo[0])] = cdrInfo[1]
		dataToUpdate = speechToText(myDictionary, _id)
		print("Data for update: " + str(len(dataToUpdate)))
		if len(dataToUpdate) > 0:
			updateCdrInfo(dataToUpdate, _id) 
		else:
			print("Sin datos para actualizar para la campanha: " + str(_id))

#obtengo campanhas del dia
campaignIds = getCampaignsOfDay()
print(campaignIds)
if len(campaignIds) > 0:
	for ids in campaignIds:
		print("Iniciando proceso para: " + str(ids) + ", " + getCurrentTime())
		speechToTextByCampaign(ids)
else:
	print("No se obtuvieron Campanhas para procesar")

