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
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using Microsoft.Win32;
using System.Drawing;

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

        public MainWindow()
        {
            InitializeComponent();


            mmiC = new MmiCommunication("localhost", 8000, "User1", "GUI");
            mmiC.Message += MmiC_Message;
            mmiC.Start();
            oPowerPoint = new PowerPoint.Application();
            oPresentation = oPowerPoint.Presentations.Add();
            examplePresentation();
            openpowerpoint = true;
            presentationMode = false;

        }

        private void MmiC_Message(object sender, MmiEventArgs e)
        {
            Console.WriteLine(e.Message);
            var doc = XDocument.Parse(e.Message);
            var com = doc.Descendants("command").FirstOrDefault().Value;
            dynamic json = JsonConvert.DeserializeObject(com);

            Console.WriteLine(json);
            Console.WriteLine("Recognize: " + (string)json.recognized[1].ToString());
            Console.WriteLine("OPEN Power Point!");


            switch ((string)json.recognized[1].ToString())
            {
                case "CropI":
                    Console.WriteLine("DO CROP IN!");
                   
                    tShape.PictureFormat.CropLeft = imgWidth*20/100;
                    tShape.PictureFormat.CropRight = imgWidth * 20 / 100;
                    tShape.PictureFormat.CropBottom = imgHeight * 20 / 100;
                    tShape.PictureFormat.CropTop = imgHeight * 20 / 100;

                    break;

                case "CropO":
                    Console.WriteLine("DO CROP OUT!");
                    //crop Picture
                    
                    tShape.PictureFormat.CropLeft = imgWidth * (20/100);
                    tShape.PictureFormat.CropRight = imgWidth * (20 / 100);
                    tShape.PictureFormat.CropBottom = imgHeight * (20 / 100);
                    tShape.PictureFormat.CropTop = imgHeight * (20 / 100);
                    break;

                case "ZoomI":
                    Console.WriteLine("DO ZOOM IN!");

                    tShape.ScaleHeight(1.2f, Microsoft.Office.Core.MsoTriState.msoFalse);
                    tShape.ScaleWidth(1.2f, Microsoft.Office.Core.MsoTriState.msoFalse);


                    break;

                case "ZoomO":
                    Console.WriteLine("DO ZOOM OUT!");
                    tShape.ScaleHeight(0.8f, Microsoft.Office.Core.MsoTriState.msoFalse);
                    tShape.ScaleWidth(0.8f, Microsoft.Office.Core.MsoTriState.msoFalse);

                    break;

                case "ThemaR":
                    Console.WriteLine("DO ADD THEME!");
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

                case "Open":
                    Console.WriteLine("OPEN Presentation Mode!");
                    oPresentation.SlideShowSettings.Run();
                    presentationMode = true;
                    break;

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

                case "NextR":
                    Console.WriteLine("DO NEXT!");

                    if (presentationMode == true)
                    {
                        oPresentation.SlideShowWindow.View.Next();
                    }
                    else
                    {
                        oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex + 1].Select();
                    }

                    break;

                case "Close":
                    Console.WriteLine("DO CLOSE!");
                    oPresentation.SlideShowWindow.View.Exit();
                    presentationMode = false;
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

            oSlide = oPresentation.Slides.Add(oPresentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutText);
            tShape = oSlide.Shapes.Title;
            tShape.TextFrame.TextRange.Text = "PowerPoint";
            tShape = oSlide.Shapes[2];
            tShape.TextFrame.TextRange.Text = "Plataforma para apresentações mais profissionais.\n" +
                "Serve de guideline numa apresentação e também para partilhar informação acerca de um tema.\n" +
                "O objetivo do trabalho é facilitar ainda mais a sua utilização.";


            oSlide = oPresentation.Slides.Add(oPresentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutText);
            tShape = oSlide.Shapes.Title;
            tShape.TextFrame.TextRange.Text = "Cenário";
            tShape = oSlide.Shapes[2];
            tShape.TextFrame.TextRange.Text = "O Carlos e a Gisela estão no ínicio de uma apresentação numa conferência internacional e esqueceram-se do ponteiro para apresentar os slides.\n" +
                "A Gisela coloca-se em frente a um dispositivo e começa a sua apresentação de forma interativa.\n" +
                "O Carlos vai apresentando e mudando os slides.\n" +
                "A Gisela encontra-se a apresentar uma figura, e faz zoom para que o público consiga observar melhor.\n" +
                "No fim fecham a apresentação e tudo correu bem.";

            oSlide = oPresentation.Slides.Add(oPresentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutText);
            tShape = oSlide.Shapes.Title;
            tShape.TextFrame.TextRange.Text = "Gestos escolhidos";
            tShape = oSlide.Shapes[2];
            tShape.TextFrame.TextRange.Text = "Avançar slide.\n" +
                "Recuar slide.\n" +
                "Crop da imagem.\n" +
                "Zoom de uma imagem.\n" +
                "Adicionar tema.\n" +
                "Abrir modo apresentação.\n" +
                "Fechar modo apresentação.";

            oSlide = oPresentation.Slides.Add(oPresentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutText);
            tShape = oSlide.Shapes.Title;
            tShape.TextFrame.TextRange.Text = "Imagem";
            tShape = oSlide.Shapes[2];

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
