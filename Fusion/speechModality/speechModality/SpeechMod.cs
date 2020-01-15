using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using mmisharp;
using Microsoft.Speech.Recognition;
using System.Timers;
using System.Windows.Shapes;
using System.Drawing;
using System.Windows.Threading;
using System.Windows.Media;
using System.Net.Sockets;
using System.Xml.Linq;
using Newtonsoft.Json;

namespace speechModality
{
    public class SpeechMod
    {
        
        Boolean closeConfirmation = false;
        Boolean confidenceConfirmation = false;
        Boolean removeSlideConfirmation = false;
        string jsonTmp;
        SemanticValue resultSemanticsTmp;

        private Tts tts;
        private SpeechRecognitionEngine sre;
        private Grammar gr;
        public event EventHandler<SpeechEventArg> Recognized;

        private int choice = 0;
        
        Timer speakingTimer;
        private Boolean assistantSpeaking = false;
        private bool assistantSpeakingFlag;

        protected virtual void onRecognized(SpeechEventArg msg)
        {
            EventHandler<SpeechEventArg> handler = Recognized;
            if (handler != null)
            {
                handler(this, msg);
            }
        }

        private Ellipse circle;

        public Dispatcher Dispatcher { get; }

        private LifeCycleEvents lce;
        private MmiCommunication mmic;
        private MmiCommunication mmiC;
        private bool presentationMode = false;

        public SpeechMod(System.Windows.Shapes.Ellipse circle, System.Windows.Threading.Dispatcher dispatcher)
        {
            //init LifeCycleEvents..
            this.circle = circle;
            this.Dispatcher = dispatcher;

           
        
            // CHANGED FOR FUSION ---------------------------------------

            lce = new LifeCycleEvents("ASR", "FUSION","speech-1", "acoustic", "command");
            mmic = new MmiCommunication("localhost",9876,"User1", "ASR");
            mmic.Start();
            // END CHANGED FOR FUSION------------------------------------
          
            mmic.Send(lce.NewContextRequest());

            mmiC = new MmiCommunication("localhost", 8000, "User2", "GUI");
            mmiC.Message += MmiC_Message;
            mmiC.Start();

            //load pt recognizer
            sre = new SpeechRecognitionEngine(new System.Globalization.CultureInfo("pt-PT"));
            gr = new Grammar(Environment.CurrentDirectory + "\\grammarInitial.grxml", "rootRule");
            sre.LoadGrammar(gr);

            
            sre.SetInputToDefaultAudioDevice();
            sre.RecognizeAsync(RecognizeMode.Multiple);
            sre.SpeechRecognized += Sre_SpeechRecognized;
            sre.SpeechHypothesized += Sre_SpeechHypothesized;

            tts = new Tts();
            // introduce assistant
            Speak("Olá, eu sou o seu assistente do PowerPoint, em que lhe posso ser util?", 5);

        }

        private void MmiC_Message(object sender, MmiEventArgs e)
        {
            Console.WriteLine(e.Message);
            var doc = XDocument.Parse(e.Message);
            var com = doc.Descendants("command").FirstOrDefault().Value;
            dynamic json = JsonConvert.DeserializeObject(com);
            Console.WriteLine(com);
            Console.WriteLine(json);
            Console.WriteLine(json.text_to_speak.ToString());
            switch ((string)json.action.ToString())
            {
                case "speak":
                    Speak((string)json.text_to_speak.ToString(), 5);
                    break;

                case "presentation":
                    gr = new Grammar(Environment.CurrentDirectory + "\\grammarPresentation.grxml", "rootRule");
                    sre.UnloadAllGrammars();
                    sre.LoadGrammar(gr);
                    break;

                case "stop_presentation":
                    gr = new Grammar(Environment.CurrentDirectory + "\\grammarEdition.grxml", "rootRule");
                    sre.UnloadAllGrammars();
                    sre.LoadGrammar(gr);
                    break;
            }
            
                    
            
        }

        //TTS
        private void Speak(String text, int seconds)
        {
            string str = "<speak version=\"1.0\"";
            str += " xmlns:ssml=\"http://www.w3.org/2001/10/synthesis\"";
            str += " xml:lang=\"pt-PT\">";
            str += text;
            str += "</speak>";

            tts.Speak(str, 0);

            // enable talking flag
            assistantSpeaking = true;
            assistantSpeakingFlag = true;

            Dispatcher.Invoke(() =>
            {
                circle.Fill = Brushes.Red;
            });

            Console.WriteLine("Assistant speaking.");

            
            speakingTimer = new Timer(seconds * 2000);
            speakingTimer.Elapsed += OnSpeakingEnded;
            speakingTimer.AutoReset = false;
            speakingTimer.Enabled = true;
        }

        private void OnSpeakingEnded(Object source, ElapsedEventArgs e)
        {
            Console.WriteLine("Assistant stopped speaking.");
            assistantSpeaking = false;
            assistantSpeakingFlag = false;
           
            Dispatcher.Invoke(() =>
            {
                circle.Fill = Brushes.Green;
            });
        }

        private void RandomSpeak(String[] choices, int seconds)
        {
            Speak(choices[choice++ % choices.Length], seconds);
        }


        private void Sre_SpeechHypothesized(object sender, SpeechHypothesizedEventArgs e)
        {
            onRecognized(new SpeechEventArg() { Text = e.Result.Text, Confidence = e.Result.Confidence, Final = false, AssistantSpeaking = assistantSpeaking });
        }

        private void Sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            onRecognized(new SpeechEventArg(){Text = e.Result.Text, Confidence = e.Result.Confidence, Final = true, AssistantSpeaking = assistantSpeaking });

            // ignore while the assistant is speaking
            if (assistantSpeaking)
            {
                return;
            }
            

            // ignore low confidance levels
            else if (e.Result.Confidence < 0.5)
            {
                return;
            }

            // if confidence is between 30% and 60%
            else if (e.Result.Confidence <= 0.6)
            {
                Speak("Desculpe, não consegui entender. Pode repetir, por favor.", 3);
                return;
            }

            
            // CHANGED FOR FUSION ---------------------------------------
            //SEND
            string json = "{ \"recognized\": [";
            foreach (var resultSemantic in e.Result.Semantics)
            {
                json+= "\""+resultSemantic.Key + "\",\"" + resultSemantic.Value.Value +"\", ";
            }
            json = json.Substring(0, json.Length - 2);
            json += "] }";
            Console.WriteLine(json);
            // END CHANGED FOR FUSION ---------------------------------------
    
            
            // if confidence is between 60% and 65%
            if (e.Result.Confidence <= 0.70)
            {
                Speak("Não tenho a certeza do que disse. Disse " + e.Result.Text + "?", 4);
                
                foreach (var resultSemantic in e.Result.Semantics)
                {
                    if (!resultSemantic.Value.Value.Equals("YES"))
                    {
                        jsonTmp = json;
                        resultSemanticsTmp = e.Result.Semantics;
                    }
                }
                confidenceConfirmation = true;
                return;
            }
            
            var exNot = lce.ExtensionNotification(e.Result.Audio.StartTime + "", e.Result.Audio.StartTime.Add(e.Result.Audio.Duration) + "", e.Result.Confidence, json);


            foreach (var resultSemantic in e.Result.Semantics)
            {
                if (resultSemantic.Value.Value.Equals("YES"))
                {
                    if (closeConfirmation)
                    {
                        gr = new Grammar(Environment.CurrentDirectory + "\\grammarInitial.grxml", "rootRule");
                        sre.UnloadAllGrammars();
                        sre.LoadGrammar(gr);
                        closeConfirmation = false;
                    }
                    if (confidenceConfirmation)
                    {
                        cleanAllConfirmations();
                        foreach (var resultSemanticTmp in resultSemanticsTmp)
                        {
                            chooseCommand(resultSemanticTmp, jsonTmp);
                        }
                        confidenceConfirmation = false;
                    }
                    
                   
                    exNot = lce.ExtensionNotification(e.Result.Audio.StartTime + "", e.Result.Audio.StartTime.Add(e.Result.Audio.Duration) + "", e.Result.Confidence, jsonTmp);
                }
                else if (resultSemantic.Value.Value.Equals("NO"))
                {
                    Speak("Devo ter percebido mal!", 2);
                    cleanAllConfirmations();
                }
                else
                {
                    cleanAllConfirmations();
                    chooseCommand(resultSemantic, json);
                }
            }


            mmic.Send(exNot);
        }
        void cleanAllConfirmations()
        {
            closeConfirmation = false;
            confidenceConfirmation = false;
        }

        void chooseCommand(KeyValuePair<String, SemanticValue> resultSemantic, string json)
        {
            if (resultSemantic.Value.Value.Equals("OPEN_POWERPOINT"))
            {
                gr = new Grammar(Environment.CurrentDirectory + "\\grammarEdition.grxml", "rootRule");
                sre.UnloadAllGrammars();
                sre.LoadGrammar(gr);
                Speak("Power Point foi aberto!", 2);
            }
            else if (resultSemantic.Value.Value.Equals("CLOSE"))
            {
                Speak("Deseja fechar o programa?", 2);
                jsonTmp = json;
                closeConfirmation = true;

            }
            else if (resultSemantic.Value.Value.Equals("START"))
            {

                gr = new Grammar(Environment.CurrentDirectory + "\\grammarPresentation.grxml", "rootRule");
                sre.UnloadAllGrammars();
                sre.LoadGrammar(gr);
            }
            else if (resultSemantic.Value.Value.Equals("STOP_PRESENTATION"))
            {
                gr = new Grammar(Environment.CurrentDirectory + "\\grammarEdition.grxml", "rootRule");
                sre.UnloadAllGrammars();
                sre.LoadGrammar(gr);
            }
            
        }
    }
}
