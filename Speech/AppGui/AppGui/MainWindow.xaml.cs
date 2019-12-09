using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Xml.Linq;
using mmisharp;
using Newtonsoft.Json;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

using System.Diagnostics;
using System.IO;
using static System.Environment;
using System.Reflection;
using System.Security.Principal;
using System.Security.AccessControl;

namespace AppGui
{
    
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private PowerPoint._Application  oPowerPoint;
        private PowerPoint._Presentation oPresentation;
        private PowerPoint._Slide oSlide;
        private PowerPoint.Shape tShape;

        private MmiCommunication mmiC;
        public MainWindow()
        {
            InitializeComponent();


            mmiC = new MmiCommunication("localhost",8000, "User1", "GUI");
            mmiC.Message += MmiC_Message;
            mmiC.Start();

        }

        private void MmiC_Message(object sender, MmiEventArgs e)
        {
            Console.WriteLine(e.Message);
            var doc = XDocument.Parse(e.Message);
            var com = doc.Descendants("command").FirstOrDefault().Value;
            dynamic json = JsonConvert.DeserializeObject(com);

            Console.WriteLine(json);
            Console.WriteLine((string)json.recognized[0].ToString());
            Console.WriteLine((string)json.keys[0].ToString());

            Tts t = new Tts();

            //TODO: See where should have the method cleanAllConfirmations()
            //https://docs.microsoft.com/en-us/office/vba/api/powerpoint.slide
            switch ((string)json.keys[0].ToString()) {
                //TODO: pôr tudo o que está para baixo aqui dentro
                case "openPowerPoint":
                    oPowerPoint = new PowerPoint.Application();
                    oPresentation = oPowerPoint.Presentations.Add();
                    break;

                case "slide":
                    switch ((string)json.recognized[0].ToString()) {
                        case "NEXT_PRESENTATION":
                            oPresentation.SlideShowWindow.View.Next();
                            break;
                        case "PREVIOUS_PRESENTATION":
                            oPresentation.SlideShowWindow.View.Previous();
                            break;
                        case "NEXT":
                            oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex + 1].Select();
                            break;
                        case "PREVIOUS":
                            oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex - 1].Select();
                            break;

                        case "JUMP_TO":
                            oPresentation.Slides[Int32.Parse(json.recognized[1].ToString())].Select();
                            break;

                        case "JUMP_TO_SLIDE_PRESENTATION":
                            oPresentation.SlideShowWindow.View.GotoSlide(Int32.Parse(json.recognized[1].ToString()));
                            break;

                        case "NEW_SLIDE":
                            if (oPresentation.Slides.Count == 0)
                            {
                                oPresentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitle).Select();
                            }
                            else
                            {
                                oPresentation.Slides.Add(oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex + 1, PowerPoint.PpSlideLayout.ppLayoutTitle).Select(); ;
                            }
                            break;


                        case "REMOVE_SLIDE":
                            if (oPresentation.Slides.Count > 0)
                            {
                                oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex].Delete();
                                t.Speak("Slide removido!");
                            }
                            else
                            {
                                t.Speak("Não existe nenhum slide.");
                            }
                            break;
                       
                    }
                    break;

                case "read":
                    switch ((string)json.recognized[0].ToString()) {
                        case "TITLE":
                            var title = oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex].Shapes.Title.TextFrame.TextRange.Text;
                            t.Speak(title);
                            break;
                        case "TEXT":
                            var text = oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex].Shapes[2].TextFrame.TextRange.Text;
                            t.Speak(text);
                            break;
                        case "NOTE":
                            var notas = oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex].NotesPage.Shapes[2].TextFrame.TextRange.Text;
                            t.Speak(notas);
                            break;
                        case "TITLE_PRESENTATION":
                            var title_pres = oPresentation.SlideShowWindow.View.Slide.Shapes.Title.TextFrame.TextRange.Text;
                            t.Speak(title_pres);
                            break;
                        case "TEXT_PRESENTATION":
                            var text_pres = oPresentation.SlideShowWindow.View.Slide.Shapes[2].TextFrame.TextRange.Text;
                            t.Speak(text_pres);
                            break;
                        case "NOTE_PRESENTATION":
                            var note_pres = oPresentation.SlideShowWindow.View.Slide.NotesPage.Shapes[2].TextFrame.TextRange.Text;
                            t.Speak(note_pres);
                            break;
                    }
                    break;

               

                case "theme":
                    switch ((string)json.recognized[0].ToString()) {
                        case "1":

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

                case "save":
                    oPresentation.Save();
                    break;

                case "color":
                    Slide activeSlide = oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex]; ;
                    TextRange textRange = activeSlide.Shapes.Title.TextFrame.TextRange; ;
                    if (json.recognized.Count > 1) {
                        switch ((string)json.recognized[0].ToString()) {
                            /*case "TITLE":
                                activeSlide = oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex];
                                textRange = activeSlide.Shapes.Title.TextFrame.TextRange;
                                break;*/
                            case "TEXT":
                                activeSlide = oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex];
                                textRange = activeSlide.Shapes[2].TextFrame.TextRange;
                                break;
                        }
                    }
                    switch ((string)json.recognized[1].ToString()) {
                        case "YELLOW":
                            textRange.Font.Color.RGB = 379903;
                            t.Speak("Mudado para Amarelo");
                            break;
                        case "RED":
                            textRange.Font.Color.RGB = 255;
                            t.Speak("Mudado para Vermelho");
                            break;
                        case "BLUE":
                            textRange.Font.Color.RGB = 16711680;
                            t.Speak("Mudado para Azul");
                            break;
                        case "GREEN":
                            textRange.Font.Color.RGB = 2540123;
                            t.Speak("Mudado para Verde");
                            break;
                        case "BLACK":
                            textRange.Font.Color.RGB = 0;
                            t.Speak("Mudado para Preto");
                            break;
                    }
                    break;

                case "example":
                    String presentationTitle = "Proposta de Trabalho 2";

                    //Save the file
                    //oPresentation.SaveAs(presentationTitle, PowerPoint.PpSaveAsFileType.ppSaveAsPresentation);

                    //Add a new slide with Title Layout
                    oSlide = oPresentation.Slides.Add(oPresentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutTitle);
                    tShape = oSlide.Shapes.Title;
                    tShape.TextFrame.TextRange.Text = presentationTitle;
                    tShape = oSlide.Shapes[2];
                    tShape.TextFrame.TextRange.Text = "Carlos Ribeiro\nGisela Pinto";

                    oSlide = oPresentation.Slides.Add(oPresentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutText);
                    tShape = oSlide.Shapes.Title;
                    tShape.TextFrame.TextRange.Text = "Tema";
                    tShape = oSlide.Shapes[2];
                    tShape.TextFrame.TextRange.Text = "Interação por voz do Powerpoint";

                    //Add Image
                    //tShape = oSlide.Shapes.AddPicture("imagePowerPoint.png", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0);

                    oSlide = oPresentation.Slides.Add(oPresentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutText);
                    tShape = oSlide.Shapes.Title;
                    tShape.TextFrame.TextRange.Text = "Features para utilizar durante uma apresentação";
                    tShape = oSlide.Shapes[2];
                    tShape.TextFrame.TextRange.Text = "Avançar slide.\n" +
                        "Recuar slide.\n" +
                        "Saltar slides, por exemplo mudar do slide 2 para o 5.\n" +
                        "Ler texto de um slide.\n" +
                        "Ler notas de um slide fazendo assim a apresentação completa.\n" +
                        "Terminar a apresentação.\n" +
                        "Controlar um video que esteja integrado no slide(Iniciar/ parar).";

                    oSlide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "Eu, Salvador, estou a ler notas do slide 3, beijinhos e abraços";

                    //oSlide = oPresentation.Slides.Add(oPresentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutText);

                    //tShape = oSlide.Shapes.AddMediaObject2(@"C:\Users\Gisela Pinto\Documents\IM\Trabalho1\im_2019_2020\Basis4Assignment2\ppt.mp4", MsoTriState.msoTrue, MsoTriState.msoTrue, 8, 8, 530, 530);

                    oSlide = oPresentation.Slides.Add(oPresentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutText);
                    tShape = oSlide.Shapes.Title;
                    tShape.TextFrame.TextRange.Text = "Features para construção de uma apresentação";
                    tShape = oSlide.Shapes[2];
                    tShape.TextFrame.TextRange.Text = "Iniciar a criação da apresentação com um dos temas sugeridos.\n" +
                        "Acrescentar um novo slide em branco/ duplicado.\n" +
                        "Remover determinado slide.\n" +
                        "Escrever o que o utilizador ditar.\n" +
                        "Guardar alterações.\n" +
                        "Mudar cor do texto(algumas cores mais usadas).";
                    oPresentation.Slides[oSlide.SlideIndex].Select();
                    break;

                case "presentation":
                    switch ((string)json.recognized[0].ToString())
                    {
                        case "START":
                            oPresentation.SlideShowSettings.Run();
                            break;
                        case "STOP_PRESENTATION":
                            oPresentation.SlideShowWindow.View.Exit();
                            break;
                    }
                    break;

                case "close":
                    oPowerPoint.Quit();
                    System.Diagnostics.Process[] pros = System.Diagnostics.Process.GetProcesses();
                    for (int i = 0; i < pros.Count(); i++) {
                        if (pros[i].ProcessName.ToLower().Contains("powerpnt")) {
                            pros[i].Kill();
                        }
                    }
                    break;
            }
            /*


            //Edition
                        

                        case "COLOR_TITLE":
                            var activeSlide = oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex];
                            rangeTitle = activeSlide.Shapes.Title.TextFrame.TextRange;
                            if (json.recognized.Count > 1) {
                                switch ((string)json.recognized[1].ToString()) {
                                    case "YELLOW":
                                        rangeTitle.Font.Color.RGB = 379903;
                                        t.Speak("Mudado para Amarelo");
                                        break;
                                    case "RED":
                                        rangeTitle.Font.Color.RGB = 255;
                                        t.Speak("Mudado para Vermelho");
                                        break;
                                    case "BLUE":
                                        rangeTitle.Font.Color.RGB = 16711680;
                                        t.Speak("Mudado para Azul");
                                        break;
                                    case "GREEN":
                                        rangeTitle.Font.Color.RGB = 2540123;
                                        t.Speak("Mudado para Verde");
                                        break;
                                    case "BLACK":
                                        rangeTitle.Font.Color.RGB = 0;
                                        t.Speak("Mudado para Preto");
                                        break;
                                }
                            }else {
                                t.Speak("Deseja mudar para que cor?");
                                selectColorTitle = true;
                            }
                            break;

                        case "COLOR_TEXT":
                            t.Speak("Deseja mudar para que cor?");
                            var activeSlide2 = oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex];
                            rangeShape = activeSlide2.Shapes[2].TextFrame.TextRange;
                            if (json.recognized.Count > 1) {
                                switch ((string)json.recognized[1].ToString()) {
                                    case "YELLOW":
                                        rangeTitle.Font.Color.RGB = 379903;
                                        t.Speak("Mudado para Amarelo");
                                        break;
                                    case "RED":
                                        rangeTitle.Font.Color.RGB = 255;
                                        t.Speak("Mudado para Vermelho");
                                        break;
                                    case "BLUE":
                                        rangeTitle.Font.Color.RGB = 16711680;
                                        t.Speak("Mudado para Azul");
                                        break;
                                    case "GREEN":
                                        rangeTitle.Font.Color.RGB = 2540123;
                                        t.Speak("Mudado para Verde");
                                        break;
                                    case "BLACK":
                                        rangeTitle.Font.Color.RGB = 0;
                                        t.Speak("Mudado para Preto");
                                        break;
                                }
                            } else {
                                t.Speak("Deseja mudar para que cor?");
                                selectColorText = true;
                            }
                            break;

                        case "YELLOW":
                            if (selectColorTitle == true) {
                                selectColorTitle = false;
                                rangeTitle.Font.Color.RGB = 379903;
                                t.Speak("Mudado para Amarelo");
                            } else if (selectColorText == true) {
                                selectColorText = false;
                                rangeShape.Font.Color.RGB = 379903;
                                t.Speak("Mudado para Amarelo");
                            } else {
                                t.Speak("Devo ter percebido mal.");

                            }
                            break;

                        case "RED":
                            if (selectColorTitle == true) {
                                selectColorTitle = false;
                                rangeTitle.Font.Color.RGB = 255;
                                t.Speak("Mudado para Vermelho");
                            } else if (selectColorText == true) {
                                selectColorText = false;
                                rangeShape.Font.Color.RGB = 255;
                                t.Speak("Mudado para Vermelho");
                            } else {
                                t.Speak("Devo ter percebido mal.");
                            }
                            break;

                        case "BLUE":
                            if (selectColorTitle == true) {
                                selectColorTitle = false;
                                rangeTitle.Font.Color.RGB = 16711680;
                                t.Speak("Mudado para Azul");
                            } else if (selectColorText == true) {
                                selectColorText = false;
                                rangeShape.Font.Color.RGB = 16711680;
                                t.Speak("Mudado para Azul");
                            } else {

                                t.Speak("Devo ter percebido mal.");

                            }
                            break;

                        case "GREEN":
                            if (selectColorTitle == true) {
                                selectColorTitle = false;
                                rangeTitle.Font.Color.RGB = 2540123;
                                t.Speak("Mudado para Verde");
                            } else if (selectColorText == true) {
                                selectColorText = false;
                                rangeShape.Font.Color.RGB = 2540123;
                                t.Speak("Mudado para Verde");
                            } else {
                                t.Speak("Devo ter percebido mal.");
                            }
                            break;

                        case "BLACK":
                            if (selectColorTitle == true) {
                                selectColorTitle = false;
                                rangeTitle.Font.Color.RGB = 0;
                                t.Speak("Mudado para Preto");
                            } else if (selectColorText == true) {
                                selectColorTitle = false;
                                rangeShape.Font.Color.RGB = 0;
                                t.Speak("Mudado para Preto");
                            } else {
                                t.Speak("Devo ter percebido mal.");
                            }
                            break;

                        */
            }

        private void foundDirectory(string dir)
        {
            throw new NotImplementedException();
        }
    }
}
