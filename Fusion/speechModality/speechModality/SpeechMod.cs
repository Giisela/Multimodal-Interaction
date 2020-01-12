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

namespace speechModality
{
    public class SpeechMod
    {
        TcpClient tcpClient = new TcpClient();

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

        public SpeechMod(System.Windows.Shapes.Ellipse circle, System.Windows.Threading.Dispatcher dispatcher)
        {
            //init LifeCycleEvents..
            this.circle = circle;
            this.Dispatcher = dispatcher;

            // CHANGED FOR FUSION ---------------------------------------

            lce = new LifeCycleEvents("ASR", "FUSION","speech-1", "acoustic", "command");
            mmic = new MmiCommunication("localhost",9876,"User1", "ASR");

            // END CHANGED FOR FUSION------------------------------------

            mmic.Send(lce.NewContextRequest());

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
            Speak("Olá, eu sou o seu assistente do PowerPoint, está pronto para mais um dia?", 12);

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

            speakingTimer = new Timer(seconds * 1000);
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
            if (e.Result.Confidence < 0.2)
            {
                return;
            }

            // if confidence is between 20% and 50%
            if (e.Result.Confidence <= 0.5)
            {
                Speak("Desculpe, não consegui entender. Pode repetir, por favor...", 2);
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
    
            
            // if confidence is between 50% and 65%
            if (e.Result.Confidence <= 0.65)
            {
                Speak("Não tenho a certeza do que disse. Disse " + e.Result.Text + "?", 2);
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
                        Speak("Adeus, até uma próxima!", 2);
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

                else if (resultSemantic.Key.Equals("read"))
                {
                    Socket socket = tcpClient.Client;
                    byte[] buffer = new byte[1000000];  // length of the text "Hello world!"
                    try
                    { // receive data with timeout 10s
                        Receive(socket, buffer, 0, buffer.Length, 10000);
                        string str = Encoding.UTF8.GetString(buffer, 0, buffer.Length);
                        Speak(str, 3);
                    }
                    catch (Exception ex) { /* ... */ }

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

            }
            else if (resultSemantic.Value.Value.Equals("CLOSE"))
            {
                Speak("Deseja fechar o programa?", 2);
                jsonTmp = json;
                closeConfirmation = true;
            }
            
        }

        public static void Receive(Socket socket, byte[] buffer, int offset, int size, int timeout)
        {
            int startTickCount = Environment.TickCount;
            int received = 0;  // how many bytes is already received
            do
            {
                if (Environment.TickCount > startTickCount + timeout)
                    throw new Exception("Timeout.");
                try
                {
                    received += socket.Receive(buffer, offset + received, size - received, SocketFlags.None);
                }
                catch (SocketException ex)
                {
                    if (ex.SocketErrorCode == SocketError.WouldBlock ||
                        ex.SocketErrorCode == SocketError.IOPending ||
                        ex.SocketErrorCode == SocketError.NoBufferSpaceAvailable)
                    {
                        // socket buffer is probably empty, wait and try again
                        System.Threading.Thread.Sleep(30);
                    }
                    else
                        throw ex;  // any serious error occurr
                }
            } while (received < size);
        }
    }
}
