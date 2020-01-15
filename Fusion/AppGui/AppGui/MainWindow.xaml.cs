using System;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Xml.Linq;
using mmisharp;
using Newtonsoft.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;
using System.Drawing;
using Microsoft.Win32;
using System.Net.Sockets;
using System.Threading;
using System.Text;

namespace AppGui
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
       
        private PowerPoint._Application oPowerPoint;
        private PowerPoint._Presentation oPresentation;
        private PowerPoint._Slide oSlide;
        private PowerPoint.Shape tShape;

        private bool openpowerpoint = false;
        private bool presentationMode = false;
        float imgWidth;
        float imgHeight;

        string startupPath = System.IO.Directory.GetCurrentDirectory();

        private MmiCommunication mmiC;
        private LifeCycleEvents lce_speechMod;
        private MmiCommunication mmic_speechMod;
        public MainWindow()
        {
            InitializeComponent();


            lce_speechMod = new LifeCycleEvents("ASR", "FUSION", "speech-2", "acoustic", "command");
            mmic_speechMod = new MmiCommunication("localhost", 8000, "User2", "ASR");
            mmic_speechMod.Send(lce_speechMod.NewContextRequest());

            mmiC = new MmiCommunication("localhost",8000, "User1", "GUI");
            mmiC.Message += MmiC_Message;
            mmiC.Start();



        }

        private void SendMsg_Tts(string message, string type) 
        {
            string json = "{\"action\":\"" + type + "\",\"text_to_speak\":\"" + message + "\"}";
            var exNot = lce_speechMod.ExtensionNotification("", "", 0, json);
            mmic_speechMod.Send(exNot);
        }

        private void MmiC_Message(object sender, MmiEventArgs e)
        {
            Console.WriteLine(e.Message);
            var doc = XDocument.Parse(e.Message);
            var com = doc.Descendants("command").FirstOrDefault().Value;
            dynamic json = JsonConvert.DeserializeObject(com);

            //json.recognized
            Console.WriteLine(json);
            Console.WriteLine("0: ");
            Console.WriteLine(json.recognized[0].ToString());
            Console.WriteLine("1: ");
            Console.WriteLine(json.recognized[1].ToString());



            //TODO: See where should have the method cleanAllConfirmations()
            //https://docs.microsoft.com/en-us/office/vba/api/powerpoint.slide
            switch ((string)json.recognized[0].ToString())
            {
                //TODO: pôr tudo o que está para baixo aqui dentro
                case "openPowerPoint":

                    oPowerPoint = new PowerPoint.Application();
                    oPresentation = oPowerPoint.Presentations.Add();
                    examplePresentation();
                    openpowerpoint = true;
                    presentationMode = false;
                    break;

                case "slide":
                    switch ((string)json.recognized[1].ToString())
                    {
                        
                        case "NEXT":
                            oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex + 1].Select();
                            break;

                        case "NEXT_PRESENTATION":
                            oPresentation.SlideShowWindow.View.Next();
                            break;

                        case "PREVIOUS":
                            oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex - 1].Select();
                            break;

                        case "PREVIOUS_PRESENTATION":
                            oPresentation.SlideShowWindow.View.Previous();
                            break;

                        case "JUMP_TO":
                            oPresentation.Slides[Int32.Parse(json.recognized[3].ToString())].Select();
                            break;

                        case "JUMP_TO_SLIDE_PRESENTATION":
                            oPresentation.SlideShowWindow.View.GotoSlide(Int32.Parse(json.recognized[3].ToString()));
                            break;


                    }
                    break;

                case "read":
                    switch ((string)json.recognized[1].ToString())
                    {
                        case "TITLE":
                            var title = oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex].Shapes.Title.TextFrame.TextRange.Text;
                            SendMsg_Tts(title, "speak");
                            break;

                        case "TEXT":
                            var text = oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex].Shapes[2].TextFrame.TextRange.Text;
                            SendMsg_Tts(text, "speak");
                            break;

                        case "NOTE":
                            var notas = oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex].NotesPage.Shapes[2].TextFrame.TextRange.Text;
                            SendMsg_Tts(notas, "speak");
                            break;

                        case "TITLE_PRESENTATION":
                            var title_pres = oPresentation.SlideShowWindow.View.Slide.Shapes.Title.TextFrame.TextRange.Text;
                            SendMsg_Tts(title_pres, "speak");
                            break;

                        case "TEXT_PRESENTATION":
                            var text_pres = oPresentation.SlideShowWindow.View.Slide.Shapes[2].TextFrame.TextRange.Text;
                            SendMsg_Tts(text_pres, "speak");
                            break;

                        case "NOTE_PRESENTATION":
                            var note_pres = oPresentation.SlideShowWindow.View.Slide.NotesPage.Shapes[2].TextFrame.TextRange.Text;
                            SendMsg_Tts(note_pres, "speak");
                            break;

                    }
                    break;

                case "1":
                    switch ((string)json.recognized[1].ToString())
                    {
                        case "CropI":
                            Console.WriteLine("DO CROP IN!");
                            SendMsg_Tts("Crop In encontra-se ativo.", "speak");
                            tShape.PictureFormat.CropLeft = imgWidth * 20 / 100;
                            tShape.PictureFormat.CropRight = imgWidth * 20 / 100;
                            tShape.PictureFormat.CropBottom = imgHeight * 20 / 100;
                            tShape.PictureFormat.CropTop = imgHeight * 20 / 100;

                            break;
                    }
                    break;

                case "2":
                    switch ((string)json.recognized[1].ToString())
                    {
                        case "CropO":
                            Console.WriteLine("DO CROP OUT!");
                            //crop Picture
                            SendMsg_Tts("Crop Out encontra-se ativo.", "speak");
                            tShape.PictureFormat.CropLeft = imgWidth * (20 / 100);
                            tShape.PictureFormat.CropRight = imgWidth * (20 / 100);
                            tShape.PictureFormat.CropBottom = imgHeight * (20 / 100);
                            tShape.PictureFormat.CropTop = imgHeight * (20 / 100);

                            break;
                    }
                    break;

                case "3":
                    switch ((string)json.recognized[1].ToString())
                    {
                        case "NextR":
                            if (presentationMode == true)
                            {
                                oPresentation.SlideShowWindow.View.Next();
                            }
                            else
                            {
                                oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex + 1].Select();
                            }
                            break;
                    }
                    break;

                case "5":
                    switch ((string)json.recognized[1].ToString())
                    {
                        case "PreviouL":
                            Console.WriteLine("DO PREVIOUS!");
                            if (presentationMode == true)
                            {
                                oPresentation.SlideShowWindow.View.Previous();
                            }
                            else
                            {
                                oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex - 1].Select();

                            }
                            break;
                    }
                    break;

                case "6":
                    switch ((string)json.recognized[1].ToString())
                    {
                        case "ThemaR":
                            switch ((string)json.recognized[3].ToString())
                            {

                                case "1":
                                    SendMsg_Tts("Tema 1 encontra-se ativo.", "speak");

                                    string dir = @"C:\Program Files (x86)\Microsoft Office\";
                                    if (Directory.Exists(dir))
                                    {
                                        oPresentation.ApplyTheme(@"C:\Program Files (x86)\Microsoft Office\root\Document Themes 16\Facet.thmx");
                                    }
                                    else
                                    {
                                        oPresentation.ApplyTheme(@"C:\Program Files\Microsoft Office\root\Document Themes 16\Facet.thmx");
                                    }


                                    break;

                                case "2":
                                    SendMsg_Tts("Tema 2 encontra-se ativo.", "speak");

                                    string dir1 = @"C:\Program Files (x86)\Microsoft Office\";
                                    if (Directory.Exists(dir1))
                                    {
                                        oPresentation.ApplyTheme(@"C:\Program Files (x86)\Microsoft Office\root\Document Themes 16\Gallery.thmx");
                                    }
                                    else
                                    {
                                        oPresentation.ApplyTheme(@"C:\Program Files\Microsoft Office\root\Document Themes 16\Gallery.thmx");
                                    }
                                    break;

                                case "3":
                                    SendMsg_Tts("Tema 3 encontra-se ativo.", "speak");

                                    string dir2 = @"C:\Program Files (x86)\Microsoft Office\";
                                    if (Directory.Exists(dir2))
                                    {
                                        oPresentation.ApplyTheme(@"C:\Program Files (x86)\Microsoft Office\root\Document Themes 16\Ion.thmx");
                                    }
                                    else
                                    {
                                        oPresentation.ApplyTheme(@"C:\Program Files\Microsoft Office\root\Document Themes 16\Ion.thmx");
                                    }
                                    break;
                            }
                            break;

                    }
                    break;

                case "7":
                    switch ((string)json.recognized[1].ToString())
                    {
                        case "ZoomI":
                            Console.WriteLine("DO ZOOM IN!");
                            SendMsg_Tts("O modo Zoom In encontra-se ativo.", "speak");
                            tShape.ScaleHeight(1.2f, Microsoft.Office.Core.MsoTriState.msoFalse);
                            tShape.ScaleWidth(1.2f, Microsoft.Office.Core.MsoTriState.msoFalse);


                            break;
                    }
                    break;


                case "8":
                    switch ((string)json.recognized[1].ToString())
                    {
                        case "ZoomO":
                            SendMsg_Tts("O modo Zoom Out encontra-se ativo.", "speak");
                            Console.WriteLine("DO ZOOM OUT!");
                            tShape.ScaleHeight(0.8f, Microsoft.Office.Core.MsoTriState.msoFalse);
                            tShape.ScaleWidth(0.8f, Microsoft.Office.Core.MsoTriState.msoFalse);

                            break;
                    }
                    break;


                case "4":
                    switch ((string)json.recognized[1].ToString())
                    {
                        case "Open":
                            
                            Console.WriteLine("OPEN Presentation Mode!");
                            oPresentation.SlideShowSettings.Run();
                            presentationMode = true;
                            SendMsg_Tts("open presentation grammer", "presentation");
                            SendMsg_Tts("O modo apresentação encontra-se ativo.", "speak");
                            break;
                    }
                    break;

                case "0":
                    switch ((string)json.recognized[1].ToString())
                    {
                        case "Close":
                            Console.WriteLine("DO CLOSE!");
                            oPresentation.SlideShowWindow.View.Exit();
                            presentationMode = false;
                            SendMsg_Tts("change to edition grammer", "stop_presentation");
                            SendMsg_Tts("O modo apresentação foi desativado.", "speak");
                            break;
                    }
                    break;

                case "presentation":
                    switch ((string)json.recognized[0].ToString())
                    {
                        case "START":
                            SendMsg_Tts("O modo apresentação foi ativado.", "speak");
                            oPresentation.SlideShowSettings.Run();
                            break;
                        case "STOP_PRESENTATION":
                            SendMsg_Tts("O modo apresentação foi ativado.", "speak");
                            oPresentation.SlideShowWindow.View.Exit();
                            break;
                    }
                    break;


                case "close":
                    oPowerPoint.Quit();
                    System.Diagnostics.Process[] pros = System.Diagnostics.Process.GetProcesses();
                    for (int i = 0; i < pros.Count(); i++)
                    {
                        if (pros[i].ProcessName.ToLower().Contains("powerpnt"))
                        {
                            pros[i].Kill();
                        }
                    }
                    presentationMode = false;
                    SendMsg_Tts("Power Point foi fechado", "speak");

                    break;
            }

        }

        private void examplePresentation()
        {

            String presentationTitle = "Proposta de Trabalho 3";

            //Add a new slide with Title Layout
            oSlide = oPresentation.Slides.Add(oPresentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutTitle);
            tShape = oSlide.Shapes.Title;
            tShape.TextFrame.TextRange.Text = presentationTitle;
            tShape = oSlide.Shapes[2];
            tShape.TextFrame.TextRange.Text = "Carlos Ribeiro 71945\nGisela Pinto 76397";
            oSlide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "Falta a minha apresentação, eu sou o Salvador, o assistente pessoal do power point";


            oSlide = oPresentation.Slides.Add(oPresentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutText);
            tShape = oSlide.Shapes.Title;
            tShape.TextFrame.TextRange.Text = "PowerPoint";
            tShape = oSlide.Shapes[2];
            tShape.TextFrame.TextRange.Text = "Plataforma para apresentações mais profissionais.\n" +
                "Serve de guideline numa apresentação e também para partilhar informação acerca de um tema.\n" +
                "O objetivo do trabalho é facilitar ainda mais a sua utilização.";
            oSlide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "O Power Point é fixe para apresentar trabalhos!";

            oSlide = oPresentation.Slides.Add(oPresentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutText);
            tShape = oSlide.Shapes.Title;
            tShape.TextFrame.TextRange.Text = "Comandos Single";
            tShape = oSlide.Shapes[2];
            tShape.TextFrame.TextRange.Text = "Para estes escolhemos os comandos que só estavam representados ou em gestos ou em voz. São os seguintes:.\n" +
                "Abrir/Fechar Power Point (Speech)\n" +
                "Salta de/para (Speech)\n"+
                "Zoom In/Out (Gestos)"+
                "Crop In/Out (Gestos)"+
                "Ler Titulo, Texto ou Notas (Speech)";
            oSlide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "O Power Point é fixe para apresentar trabalhos!";


            oSlide = oPresentation.Slides.Add(oPresentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutText);
            tShape = oSlide.Shapes.Title;
            tShape.TextFrame.TextRange.Text = "Comandos Redundancy";
            tShape = oSlide.Shapes[2];
            tShape.TextFrame.TextRange.Text = "Avançar Slide\n" +
                "Recuar Slide.\n" +
                "Abrir Modo apresentação.\n" +
                "Fechar Modo Apresentação.\n" ;
            oSlide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "O Power Point é fixe para apresentar trabalhos!";


            oSlide = oPresentation.Slides.Add(oPresentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutText);
            tShape = oSlide.Shapes.Title;
            tShape.TextFrame.TextRange.Text = "Comandos Complementary";
            tShape = oSlide.Shapes[2];
            tShape.TextFrame.TextRange.Text = "Escolhemos para complementariedade o Mudar tema. Inicialmente dizemos tema e o respetivo número e de seguida executamos o gesto correspondente\n";
            oSlide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "O Power Point é fixe para apresentar trabalhos!";


            oSlide = oPresentation.Slides.Add(oPresentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutText);
            tShape = oSlide.Shapes.Title;
            tShape.TextFrame.TextRange.Text = "Imagem";
            tShape = oSlide.Shapes[2];
            oSlide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "Kitty cat power!";


            //Resize image
            OpenFileDialog open = new OpenFileDialog();

            string workingDirectory = Environment.CurrentDirectory;
            string parentDirectory = Directory.GetParent(workingDirectory).Parent.FullName;
            open.FileName = parentDirectory + @"\kitty_cat.jpg";
            Console.WriteLine(open.FileName);
            FileInfo file = new FileInfo(open.FileName);
            var sizeInBytes = file.Length;

            Bitmap img = new Bitmap(open.FileName);

            var imageHeight = img.Height;
            var imageWidth = img.Width;


            //to move image just modify left top from the function AddPicture
            tShape = oSlide.Shapes.AddPicture(open.FileName, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, imageWidth, imageHeight);

            imgWidth = tShape.Width;
            imgHeight = tShape.Height;
        }

    }
}
