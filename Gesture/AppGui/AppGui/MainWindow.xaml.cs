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
            tShape.TextFrame.TextRange.Text = "Carlos Ribeiro\nGisela Pinto";

            oSlide = oPresentation.Slides.Add(oPresentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutText);
            tShape = oSlide.Shapes.Title;
            tShape.TextFrame.TextRange.Text = "Tema";
            tShape = oSlide.Shapes[2];
            tShape.TextFrame.TextRange.Text = "Interação com gestos no Powerpoint";


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
                "Fechar modo apresentação. \n";


            oSlide = oPresentation.Slides.Add(oPresentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutText);
            tShape = oSlide.Shapes.Title;
            tShape.TextFrame.TextRange.Text = "Imagem";
            tShape = oSlide.Shapes[2];

            //Resize image
            string startupPath = System.IO.Directory.GetCurrentDirectory();
            //string startupPath = Environment.CurrentDirectory;
            Console.WriteLine(startupPath);

            OpenFileDialog open = new OpenFileDialog();
            open.FileName = startupPath + @"\kitty_cat.jpg";
            FileInfo file = new FileInfo(open.FileName);
            var sizeInBytes = file.Length;

            Bitmap img = new Bitmap(open.FileName);

            var imageHeight = img.Height;
            var imageWidth = img.Width;


            //to move image just modify left top from the function AddPicture
            tShape = oSlide.Shapes.AddPicture(@"kitty_cat.jpg", Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, imageWidth, imageHeight);

            imgWidth = tShape.Width;
            imgHeight = tShape.Height;
        }
    }
}
