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
        TcpClient tcpClient = new TcpClient();
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

                            if (presentationMode == true)
                            {
                                oPresentation.SlideShowWindow.View.Next();
                            }
                            else
                            {
                                oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex + 1].Select();
                            }
                            break;
                        case "PREVIOUS":
                            if (presentationMode == true)
                            {
                                oPresentation.SlideShowWindow.View.Previous();
                            }
                            else
                            {
                                oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex - 1].Select();

                            }
                            break;

                        case "JUMP_TO":
                            if (presentationMode == true)
                            {
                                oPresentation.SlideShowWindow.View.GotoSlide(Int32.Parse(json.recognized[3].ToString()));
                            }
                            else
                            {
                                oPresentation.Slides[Int32.Parse(json.recognized[3].ToString())].Select();
                            }
                            break;

                       
                    }
                    break;

                case "read":
                    switch ((string)json.recognized[1].ToString())
                    {
                        case "TITLE":
                            if (presentationMode == true)
                            {
                                var title_pres = oPresentation.SlideShowWindow.View.Slide.Shapes.Title.TextFrame.TextRange.Text;
                                Socket socket = tcpClient.Client;

                                try
                                { // sends the text with timeout 10s
                                    Send(socket, Encoding.UTF8.GetBytes(title_pres), 0, title_pres.Length, 10000);
                                }
                                catch (Exception ex) { /* ... */ }
                                //t.Speak(title_pres);
                            }
                            else
                            {

                                var title = oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex].Shapes.Title.TextFrame.TextRange.Text;
                                Socket socket = tcpClient.Client;

                                try
                                { // sends the text with timeout 10s
                                    Send(socket, Encoding.UTF8.GetBytes(title), 0, title.Length, 10000);
                                }
                                catch (Exception ex) { /* ... */ }
                                //t.Speak(title);
                            }
                            break;
                        case "TEXT":
                            if (presentationMode == true)
                            {
                                var text_pres = oPresentation.SlideShowWindow.View.Slide.Shapes[2].TextFrame.TextRange.Text;
                                Socket socket = tcpClient.Client;

                                try
                                { // sends the text with timeout 10s
                                    Send(socket, Encoding.UTF8.GetBytes(text_pres), 0, text_pres.Length, 10000);
                                }
                                catch (Exception ex) { /* ... */ }
                                //t.Speak(text_pres);
                            }
                            else
                            {
                                var text = oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex].Shapes[2].TextFrame.TextRange.Text;
                                Socket socket = tcpClient.Client;

                                try
                                { // sends the text with timeout 10s
                                    Send(socket, Encoding.UTF8.GetBytes(text), 0, text.Length, 10000);
                                }
                                catch (Exception ex) { /* ... */ }
                                //t.Speak(text);
                            }
                            break;
                        case "NOTE":
                            if (presentationMode == true)
                            {
                                var note_pres = oPresentation.SlideShowWindow.View.Slide.NotesPage.Shapes[2].TextFrame.TextRange.Text;
                                Socket socket = tcpClient.Client;

                                try
                                { // sends the text with timeout 10s
                                    Send(socket, Encoding.UTF8.GetBytes(note_pres), 0, note_pres.Length, 10000);
                                }
                                catch (Exception ex) { /* ... */ }
                                //t.Speak(note_pres);
                            }
                            else 
                            {
                                var notas = oPresentation.Slides[oPowerPoint.ActiveWindow.Selection.SlideRange.SlideIndex].NotesPage.Shapes[2].TextFrame.TextRange.Text;
                                Socket socket = tcpClient.Client;

                                try
                                { // sends the text with timeout 10s
                                    Send(socket, Encoding.UTF8.GetBytes(notas), 0, notas.Length, 10000);
                                }
                                catch (Exception ex) { /* ... */ }
                                //t.Speak(notas);
                            }
                            break;
                    
                    }
                    break;

                case "1":
                    switch ((string)json.recognized[1].ToString())
                    {
                        case "CropI":
                            Console.WriteLine("DO CROP IN!");

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

                    }
                    break;

                case "7":
                    switch ((string)json.recognized[1].ToString())
                    {
                        case "ZoomI":
                            Console.WriteLine("DO ZOOM IN!");

                            tShape.ScaleHeight(1.2f, Microsoft.Office.Core.MsoTriState.msoFalse);
                            tShape.ScaleWidth(1.2f, Microsoft.Office.Core.MsoTriState.msoFalse);


                            break;
                    }
                    break;


                case "8":
                    switch ((string)json.recognized[1].ToString())
                    {
                        case "ZoomO":
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

        public static void Send(Socket socket, byte[] buffer, int offset, int size, int timeout)
        {
            int startTickCount = Environment.TickCount;
            int sent = 0;  // how many bytes is already sent
            do
            {
                if (Environment.TickCount > startTickCount + timeout)
                    throw new Exception("Timeout.");
                try
                {
                    sent += socket.Send(buffer, offset + sent, size - sent, SocketFlags.None);
                }
                catch (SocketException ex)
                {
                    if (ex.SocketErrorCode == SocketError.WouldBlock ||
                        ex.SocketErrorCode == SocketError.IOPending ||
                        ex.SocketErrorCode == SocketError.NoBufferSpaceAvailable)
                    {
                        // socket buffer is probably full, wait and try again
                        Thread.Sleep(30);
                    }
                    else
                        throw ex;  // any serious error occurr
                }
            } while (sent < size);
        }
    }
}
