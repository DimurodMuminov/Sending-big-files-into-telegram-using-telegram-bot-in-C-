using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using QRCoder;
using Telegram.Bot;
using Xceed.Words.NET;
using System.Reflection.Metadata;
using Telegram.Bot.Types;
using Xceed.Document.NET;
using PdfiumViewer;
using System.Diagnostics;
using WTelegram;
using System.Text;
using PdfSharp.Pdf.Advanced;
using static System.Net.WebRequestMethods;


namespace Urdu_library5
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            // Step 1: Open a File Dialog to select the PDF file
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "PDF files (*.pdf)|*.pdf";


            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string pdfFilePath = openFileDialog.FileName;

                // Step 2: Send the PDF to the Telegram channel
                string botToken = "bot_token";
                string channelUsername = "@channel_user_name";
                // test channel  channelUsername = "@test_library";

                progressBar1.Visible = true;
                label1.Visible = true;
                progressBar1.Value = 10;
                
                var pdfStream = new FileStream(pdfFilePath, FileMode.Open, FileAccess.Read);
                
                FileInfo fileInfo = new FileInfo(pdfFilePath);
                long fileSizeInBytes = fileInfo.Length;
                double fileSizeInMB = fileSizeInBytes / (1024.0 * 1024.0);  // Convert bytes to MB

                string telegramPostUrl;
                if(fileSizeInMB<50)
                {
                    //Telegram botClient bilan

                    var botClient =new TelegramBotClient(botToken);

                    string filename1 = Path.GetFileName(pdfFilePath);

                   var message = await botClient.SendDocumentAsync(
                        chatId: channelUsername,
                        document: InputFile.FromStream(pdfStream, filename1)
                    );
                    telegramPostUrl = $"https://t.me/{channelUsername.Substring(1)}/{message.MessageId}";

                }
                else
                {
                    //Wtelegram bilan
                    int apiId = 1;//apiID from telegram.org
                    string apiHash = "apiHash";
                    string botToken1 = botToken;

                    try
                    {
                        StreamWriter WTelegramLogs = new StreamWriter("WTelegramBot.log", true, Encoding.UTF8) { AutoFlush = true };
                        WTelegram.Helpers.Log = (lvl, str) => WTelegramLogs.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{"TDIWE!"[lvl]}] {str}");
                    }
                    catch
                    {
                    }

                    using var connection = new Microsoft.Data.Sqlite.SqliteConnection(@"Data Source=WTelegramBot.sqlite");

                    using var bot = new WTelegram.Bot(botToken1, apiId, apiHash, connection);

                    string filename = Path.GetFileName(pdfFilePath);

                    var message = await bot.SendDocument(
                        chatId: channelUsername,
                        document: InputFile.FromStream(pdfStream, filename)
                    );
                    telegramPostUrl = $"https://t.me/{channelUsername.Substring(1)}/{message.MessageId}";                    
                }

                
                // Step 3: Generate the QR Code
                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(telegramPostUrl, QRCodeGenerator.ECCLevel.Q);
                QRCode qrCode = new QRCode(qrCodeData);
                Bitmap qrCodeImage = qrCode.GetGraphic(20);
                string qrCodeImagePath = Path.Combine(Path.GetDirectoryName(pdfFilePath), "qr_code.png");
                qrCodeImage.Save(qrCodeImagePath, ImageFormat.Png);


                // Step 4: Extract the first page of the PDF as an image
                string firstPageImagePath = Path.Combine(Path.GetDirectoryName(pdfFilePath), "pdf_first_page.png");
                ExtractFirstPageAsImage(pdfFilePath, firstPageImagePath);

                // Step 5: Insert the QR code and the first page of the PDF into the Word document
                string wordDocumentPath = @"C:\kitoblar.docx";
                AddToWordDocument(wordDocumentPath, qrCodeImagePath, firstPageImagePath);

                progressBar1.Visible = false;
                label1.Visible = false;                          
            }
        }

        // Function to extract the first page of the PDF as an image
        private void ExtractFirstPageAsImage(string pdfFilePath, string outputImagePath)
        {
            // Load the PDF using PdfiumViewer
            using (var pdfDocument = PdfiumViewer.PdfDocument.Load(pdfFilePath))
            {
                // Render the first page as an image
                using (var image = pdfDocument.Render(0, 300, 300, PdfRenderFlags.CorrectFromDpi))
                {
                    // Save the image as PNG
                    image.Save(outputImagePath, ImageFormat.Png);
                }
            }
        }

        // Function to create or add to the Word document
        private void AddToWordDocument(string wordDocumentPath, string qrCodeImagePath, string firstPageImagePath)
        {
            // Create or open the Word document
            DocX doc;
            if (System.IO.File.Exists(wordDocumentPath))
            {
                doc = DocX.Load(wordDocumentPath);
            }
            else
            {
                doc = DocX.Create(wordDocumentPath);
            }

            /*
            doc.InsertParagraph("Kitob:");
            var pdfImage = doc.AddImage(firstPageImagePath);
            var pdfPicture = pdfImage.CreatePicture();
            doc.InsertParagraph().InsertPicture(pdfPicture);
            

            // Insert the QR code image
            doc.InsertParagraph("Yuqoridagi kitobning QR kodi:");
            var qrImage = doc.AddImage(qrCodeImagePath);
            var qrPicture = qrImage.CreatePicture();
            doc.InsertParagraph().InsertPicture(qrPicture);


            // Save the Word document
            doc.Save();*/

            // adding qr code and page in one paragraph
            // Load the first page image (assuming it's already saved to a file)
            var pdfImage = doc.AddImage(firstPageImagePath);
            var pdfPicture = pdfImage.CreatePicture();

            // Load the QR code image
            var qrImage = doc.AddImage(qrCodeImagePath);
            var qrPicture = qrImage.CreatePicture();

            // Resize the images to fit on the page appropriately
            float maxImageWidth = 250f; // Max width for the PDF page image
            if (pdfPicture.Width > maxImageWidth)
            {
                float scaleFactor = maxImageWidth / pdfPicture.Width;
                pdfPicture.Width = maxImageWidth;
                pdfPicture.Height *= scaleFactor;
            }

            // Resize QR code to a smaller size
            float qrMaxSize = 150f; // Max size for QR code
            if (qrPicture.Width > qrMaxSize)
            {
                float scaleFactor = qrMaxSize / qrPicture.Width;
                qrPicture.Width = qrMaxSize;
                qrPicture.Height *= scaleFactor;
            }

            // Insert both pictures into the same paragraph
            var paragraph = doc.InsertParagraph();
            paragraph.AppendPicture(pdfPicture);  // Append the first page picture
            paragraph.Append("  ");               // Add a space between the images
            paragraph.AppendPicture(qrPicture);   // Append the QR code image next to the first image

            // Save the Word document
            doc.Save();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label1.Visible = false;
            progressBar1.Visible = false;
            // progressBar1.Enabled = false;
            progressBar1.Minimum = 10;
            progressBar1.Maximum = 100;
            progressBar1.Step = 30;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string url = "https://t.me/Dilmurod_Muminov";
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true // This is required to open the URL in the default browser
                });
            }
            catch
            {
                MessageBox.Show("Telegram orqali https://t.me/Dilmurod_Muminov ga bog'laning ");
            }
        }
    }
}
