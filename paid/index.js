// AWS required dependencies
import aws from "aws-sdk";

import {
  TranscribeStreamingClient,
  StartMedicalStreamTranscriptionCommand,
} from "@aws-sdk/client-transcribe-streaming";

// IBM required dependencies
import IBM from "ibm-watson/speech-to-text/v1";
import AuthIBM from "ibm-watson/auth";

// AZURE required dependencies
import AZURE from "microsoft-cognitiveservices-speech-sdk";


// IMPORT AUDIOS FROM SAMPLE

const ibmSTT = new IBM({
  authenticator: new AuthIBM({
    apikey: "YOUR_API_KEY",
  }),
  serviceUrl: "https://api.us-east.speech-to-text.watson.cloud.ibm.com",
});

// AWS SETUP
// https://docs.aws.amazon.com/transcribe/latest/dg/getting-started-sdk.html

const AWSClient = new TranscribeStreamingClient({
  region: "us-east-1",
});

// AZURE SETUP

const AzureSpeech = AZURE.SpeechConfig.fromSubscription("YOUR_API_KEY", "us-east-1");

const processData = async () => {

  // IBM LOGIC
  const transcriptIBM = ibmSTT.recognize();
  console.log(transcriptIBM);

  // AWS LOGIC
  const params = {};
  const command = new StartMedicalStreamTranscriptionCommand(params);
  try {
    const data = await AWSClient.send(command);
    console.log(data);
  } catch (error) {
    console.log(error);
  }

  // AZURE LOGIC

  const azureConfig = AZURE.AudioConfig.fromWavFileInput("sample.wav"); // fs.readFileSync(Audio.wav)
  const recognizerAzure = new AZURE.SpeechRecognizer(AzureSpeech, azureConfig)

  recognizerAzure.recognizeOnceAsync(result => {
    console.log('Recognized text', result.text)
    recognizerAzure.close()
  })


};

processData();
