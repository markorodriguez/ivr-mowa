// required dependencies
const fs = require("fs");
const util = require("util");
const xlsx = require("xlsx");

const aws = require('aws-sdk')
const { TranscribeStreamingClient, StartMedicalStreamTranscriptionCommand } = require('@aws-sdk/client-transcribe-streaming')

// IBM required dependencies

const IBM = require("ibm-watson/speech-to-text/v1");
const { IamAuthenticator } = require("ibm-watson/auth");

// AZURE required dependencies
const AZURE = require("microsoft-cognitiveservices-speech-sdk");

// IMPORT AUDIOS FROM SAMPLE

const audios = fs.readdirSync("./audios/masivo");
const limitedAudios = audios.slice(0, 50);

const ibmSTT = new IBM({
  authenticator: new IamAuthenticator({
    apikey: "rsHwjvv-Do805tyu7UtQXxnGO4jujegREiszyWDx5dIx",
  }),
  serviceUrl: "https://api.us-east.speech-to-text.watson.cloud.ibm.com",
});

// AWS SETUP
// https://docs.aws.amazon.com/transcribe/latest/dg/getting-started-sdk.html

const AWSClient = new TranscribeStreamingClient({
  region: "us-east-1"
});

// AZURE SETUP

const AzureSpeech = AZURE.SpeechConfig.fromSubscription(
  "41e7b5ec7c0e4a65b90d530804127a0c",
  "eastus"
);

AzureSpeech.speechRecognitionLanguage = "es-ES";

const processDataAzure = async () => {
  // AWS LOGIC
  /*
  const params = {};
  const command = new StartMedicalStreamTranscriptionCommand(params);
  try {
    const data = await AWSClient.send(command);
    console.log(data);
  } catch (error) {
    console.log(error);
  }
*/
  // AZURE LOGIC

  const azureConfig = AZURE.AudioConfig.fromWavFileInput(
    fs.readFileSync("./audios/lidia.wav")
  ); // fs.readFileSync(Audio.wav)
  const recognizerAzure = new AZURE.SpeechRecognizer(AzureSpeech, azureConfig);

  recognizerAzure.recognizeOnceAsync((result) => {
    console.log("Recognized text", result.text, typeOf(result.text));
    recognizerAzure.close();
  });
};

const processDataIBM = async () => {
  // IBM LOGIC

  const recognizeParams = {
    audio: fs.readFileSync("./audios/audio2.wav"),
    model: "es-ES_NarrowbandModel",
    contentType: "audio/wav",
    wordAlternativesThreshold: 0.9,
    keywords: ["sí", "si", "no"],
    keywordsThreshold: 0.5,
  };

  const transcriptIBM = await ibmSTT.recognize(recognizeParams);
  console.log(util.inspect(transcriptIBM, { showHidden: false, depth: null }));
};

//  processDataAzure();
// processDataIBM();

let azErrors = 0
let ibmErrors = 0

const processExcel = async () => {
  const comparison = xlsx.utils.book_new();
  const data = [];

  await Promise.all(
    limitedAudios.map(async (audio, index) => {
      console.log(`audio ${index + 1} de ${limitedAudios.length}`)
      let azureTranscript = "-";

      console.log('procesando azure')
      const azureConfig = AZURE.AudioConfig.fromWavFileInput(
        fs.readFileSync(`./audios/masivo/${audio}`)
      ); // fs.readFileSync(Audio.wav)
      const recognizerAzure = new AZURE.SpeechRecognizer(
        AzureSpeech,
        azureConfig
      );

      recognizerAzure.recognizeOnceAsync((result, err) => {
        if (result.text.length > 0 && result.text != "undefined") {
          azureTranscript = result.text;
          azErrors++
        }
        recognizerAzure.close();
      });

      let ibmResponse;

      try {
        console.log('procesando IBM')
        const recognizeParams = {
          audio: fs.readFileSync(`./audios/masivo/${audio}`),
          model: "es-ES_NarrowbandModel",
          contentType: "audio/wav",
          wordAlternativesThreshold: 0.9,
          keywords: ["sí", "si", "no"],
          keywordsThreshold: 0.5,
        };

        const transcriptIBM = await ibmSTT.recognize(recognizeParams);
        ibmResponse =
          transcriptIBM.result.results[0].alternatives[0].transcript;
      } catch (error) {
        ibmResponse = " ";
        ibmErrors++
        console.log(error)
      }

      try {
        
      } catch (error) {
        
      }

      data.push({
        audio,
        azure: azureTranscript,
        ibm: ibmResponse,
      });
    })
  );

  console.log('Se terminó de procesar...')
  console.log(`Errores de azure: ${limitedAudios.length - azErrors}`)
  console.log(`Errores de IBM: ${ibmErrors}`)
  console.log('Creando excel...')

  const worksheet = xlsx.utils.json_to_sheet(data);
  xlsx.utils.book_append_sheet(comparison, worksheet, "HOJA_1");
  xlsx.writeFile(comparison, "comparison.xlsx");
};

processExcel();
