using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using mmisharp;
using Microsoft.Speech.Recognition;

namespace speechModality
{
    public class SpeechMod
    {
        Boolean closeConfirmation = false;
        Boolean confidenceConfirmation = false;
        Boolean removeSlideConfirmation = false;
        string jsonTmp;
        SemanticValue resultSemanticsTmp;

        private SpeechRecognitionEngine sre;
        private Grammar gr;
        public event EventHandler<SpeechEventArg> Recognized;
        protected virtual void onRecognized(SpeechEventArg msg)
        {
            EventHandler<SpeechEventArg> handler = Recognized;
            if (handler != null)
            {
                handler(this, msg);
            }
        }

        private LifeCycleEvents lce;
        private MmiCommunication mmic;

        public SpeechMod()
        {
            //init LifeCycleEvents..
            lce = new LifeCycleEvents("ASR", "FUSION","speech-1", "acoustic", "command"); // LifeCycleEvents(string source, string target, string id, string medium, string mode)
            //mmic = new MmiCommunication("localhost",9876,"User1", "ASR");  //PORT TO FUSION - uncomment this line to work with fusion later
            mmic = new MmiCommunication("localhost", 8000, "User1", "ASR"); // MmiCommunication(string IMhost, int portIM, string UserOD, string thisModalityName)

            mmic.Send(lce.NewContextRequest());

            //load pt recognizer
            sre = new SpeechRecognitionEngine(new System.Globalization.CultureInfo("pt-PT"));
            gr = new Grammar(Environment.CurrentDirectory + "\\grammarInitial.grxml", "rootRule");
            sre.LoadGrammar(gr);

            
            sre.SetInputToDefaultAudioDevice();
            sre.RecognizeAsync(RecognizeMode.Multiple);
            sre.SpeechRecognized += Sre_SpeechRecognized;
            sre.SpeechHypothesized += Sre_SpeechHypothesized;

        }

        private void Sre_SpeechHypothesized(object sender, SpeechHypothesizedEventArgs e)
        {
            onRecognized(new SpeechEventArg() { Text = e.Result.Text, Confidence = e.Result.Confidence, Final = false });
        }

        private void Sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            onRecognized(new SpeechEventArg(){Text = e.Result.Text, Confidence = e.Result.Confidence, Final = true});

            Tts t = new Tts();

            // ignore low confidance levels
            if (e.Result.Confidence < 0.2) {
                return;
            }

            // if confidence is between 20% and 50%
            if (e.Result.Confidence <= 0.5) {
                t.Speak("Desculpe, não consegui entender. Pode repetir, por favor...");
                return;
            }

            //SEND
            string json = "{ \"recognized\": [";
            foreach (var resultSemantic in e.Result.Semantics)
            {
                json += "\"" + resultSemantic.Value.Value +"\", ";
            }
            json = json.Substring(0, json.Length - 2);
            json += "]";

            Console.WriteLine(json);

            json += ", \"keys\": [";
            foreach (var resultSemantic in e.Result.Semantics) {
                json += "\"" + resultSemantic.Key + "\", ";
            }
            json = json.Substring(0, json.Length - 2);
            json += "] }";

            Console.WriteLine(json);
            
            // if confidence is between 50% and 65%
            if (e.Result.Confidence <= 0.65) {
                t.Speak("Não tenho a certeza do que disse. Disse " + e.Result.Text + "?");
                foreach (var resultSemantic in e.Result.Semantics) {
                    if (!resultSemantic.Value.Value.Equals("YES")) {
                        jsonTmp = json;
                        resultSemanticsTmp = e.Result.Semantics;
                    }
                }
                confidenceConfirmation = true;
                return;
            }

            var exNot = lce.ExtensionNotification(e.Result.Audio.StartTime + "", e.Result.Audio.StartTime.Add(e.Result.Audio.Duration) + "", e.Result.Confidence, json);

            foreach (var resultSemantic in e.Result.Semantics) {
                if (resultSemantic.Value.Value.Equals("YES")) {
                    if (closeConfirmation) {
                        gr = new Grammar(Environment.CurrentDirectory + "\\grammarInitial.grxml", "rootRule");
                        sre.UnloadAllGrammars();
                        sre.LoadGrammar(gr);
                        closeConfirmation = false;
                    }
                    if (confidenceConfirmation) {
                        cleanAllConfirmations();
                        foreach (var resultSemanticTmp in resultSemanticsTmp) {
                            chooseCommand(t, resultSemanticTmp, jsonTmp);
                        }
                        confidenceConfirmation = false;
                    }
                    if (removeSlideConfirmation) {
                        removeSlideConfirmation = false;
                    }
                    exNot = lce.ExtensionNotification(e.Result.Audio.StartTime + "", e.Result.Audio.StartTime.Add(e.Result.Audio.Duration) + "", e.Result.Confidence, jsonTmp);
                } else if (resultSemantic.Value.Value.Equals("NO")) {
                    t.Speak("Devo ter percebido mal!");
                    cleanAllConfirmations();
                } else {
                    cleanAllConfirmations();
                    chooseCommand(t, resultSemantic, json);
                }
            }

            mmic.Send(exNot);
        }
        void cleanAllConfirmations() {
            closeConfirmation = false;
            confidenceConfirmation = false;
        }

        void chooseCommand(Tts t, KeyValuePair<String, SemanticValue> resultSemantic, string json) {
            if (resultSemantic.Value.Value.Equals("OPEN_POWERPOINT")) {
                gr = new Grammar(Environment.CurrentDirectory + "\\grammarEdition.grxml", "rootRule");
                sre.UnloadAllGrammars();
                sre.LoadGrammar(gr);
            } else if (resultSemantic.Value.Value.Equals("CLOSE")) {
                t.Speak("Deseja fechar o programa?");
                jsonTmp = json;
                closeConfirmation = true;
            } else if (resultSemantic.Value.Value.Equals("START")) {
                gr = new Grammar(Environment.CurrentDirectory + "\\grammarPresentation.grxml", "rootRule");
                sre.UnloadAllGrammars();
                sre.LoadGrammar(gr);
            } else if (resultSemantic.Value.Value.Equals("STOP_PRESENTATION")) {
                gr = new Grammar(Environment.CurrentDirectory + "\\grammarEdition.grxml", "rootRule");
                sre.UnloadAllGrammars();
                sre.LoadGrammar(gr);
            } else if (resultSemantic.Value.Value.Equals("REMOVE_SLIDE")) {
                t.Speak("Deseja remover o slide?");
                removeSlideConfirmation = true;
            }
        }
    }
}
