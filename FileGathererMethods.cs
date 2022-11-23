using System;
using System.Configuration;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
////ZIP
using System.IO.Compression;
////////Mail Libraries
using MimeKit;
using MailKit;
using MailKit.Search;
using MailKit.Security;
using MailKit.Net.Imap;
/////////SPIRE PDF
using Spire.Pdf;
using Spire.Pdf.Graphics;
///////Credentials
using CredentialManagement;
using MsgReader.Outlook;
using System.Diagnostics;

namespace PGNiG_FileProcessor
{
    public class FileGatherer
    {
        public class MyWebClient
        {
            public string GetPassword(String KeyPair)
            {
                try
                {
                    using (var cred = new Credential())
                    {
                        cred.Target = KeyPair;
                        cred.Load();
                        return cred.Password;
                    }
                }
                catch (Exception ex)
                {
                }
                return "";
            }
            public string GetUsername(String KeyPair)
            {
                try
                {
                    using (var cred = new Credential())
                    {
                        cred.Target = KeyPair;
                        cred.Load();
                        return cred.Username;
                    }
                }
                catch (Exception ex)
                {
                }
                return "";
            }

        }

        public static void ConvertOfficeFiles(string path)
        {
            string[] extensions = { ".docx", ".doc", ".xls", ".xlsx", ".rtf", ".ods", ".odt" };
            try
            {
                string[] officefiles = Directory.GetFiles(path, "*.*")
                .Where(f => extensions.Contains(new FileInfo(f).Extension.ToLower())).ToArray();
                foreach (string file in officefiles)
                {
                    LibreOfficeConverter.Run(file, path);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("There was an error: {0}", ex.Message);
            }

        }

        public static void CollectNetworkFiles()
        {
            string NetworkFolder = ConfigurationManager.AppSettings.Get("NetworkFolder");
            string InputClassificationFolder = ConfigurationManager.AppSettings.Get("InputClassificationFolder");
            string InitialFolder = ConfigurationManager.AppSettings.Get("InitialFolder");
            string ProcessedZIPFilesFolder = ConfigurationManager.AppSettings.Get("ProcessedZIPFiles");

            try
            {
                string[] array2 = Directory.GetFiles(NetworkFolder, "*.zip");
                foreach (string zipfile in array2)
                {
                    GetAttachmentsFromZIPFile(zipfile);
                    //File.Delete(zipfile);
                    File.Move(zipfile, ProcessedZIPFilesFolder+@"\"+Path.GetFileName(zipfile));
                    
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("There was an error: {0}", ex.Message);
            }
            

        }

        public static void GetAttachmentsFromZIPFile(string zipfile)
        {
            string InputClassificationFolder = ConfigurationManager.AppSettings.Get("InputClassificationFolder");
            string InitialFolder = ConfigurationManager.AppSettings.Get("InitialFolder");

            string path = $"{InitialFolder}\\temp";
            System.IO.Directory.CreateDirectory(path);
            ZipFile.ExtractToDirectory(zipfile, path);

            string[] array = Directory.GetFiles(path, "*.msg");
            MsgReader.Outlook.Storage.Message messagefile = new Storage.Message(array[0], FileAccess.ReadWrite);
            string date = messagefile.Headers.DateSent.ToLocalTime().ToString("yyyy'-'MM'-'dd'T'HH'-'mm'-'ss");
            

            CreateMailBodyPDFFile(messagefile.BodyText, path);
            //CreateSignalFile(path);
            messagefile.Dispose();
            //delete .msg file
            File.Delete(array[0]);
            MoveFolderForClassification(path, InputClassificationFolder, date);
            ////delete temp folder
            Directory.Delete(path,true);

        }

        public static void MoveFolderForClassification(string inputFolder, string outputFolder, string date)
        {
            string InputClassificationFolder = ConfigurationManager.AppSettings.Get("InputClassificationFolder");
            Directory.Move(inputFolder, $"{outputFolder}\\{date}");
            ConvertOfficeFiles($"{outputFolder}\\{date}");
            CreateSignalFile($"{outputFolder}\\{date}");
        }

        public static void CreateSignalFile(string path)
        {
            // Create a new file     
            using (FileStream fs = File.Create(path + "\\SignalFile.xml"))
            {
                // Add some text to file    
                byte[] author = new UTF8Encoding(true).GetBytes("1");
                fs.Write(author, 0, author.Length);
            }

        }

        public static void CreateMailBodyPDFFile(string message, string path)
        {
            string MailBody = message;
            PdfDocument pdf = new PdfDocument();
            PdfPageBase page = pdf.Pages.Add();
            PdfFont font = new PdfFont(PdfFontFamily.Helvetica, 11);
            PdfTextLayout textLayout = new PdfTextLayout();
            textLayout.Break = PdfLayoutBreakType.FitPage;
            textLayout.Layout = PdfLayoutType.Paginate;
            PdfStringFormat format = new PdfStringFormat();
            //format.Alignment = PdfTextAlignment.Justify;
            format.LineSpacing = 20f;
            PdfTextWidget textWidget = new PdfTextWidget(MailBody, font, PdfBrushes.Black);
            textWidget.StringFormat = format;
            RectangleF bounds = new RectangleF(new PointF(10, 25), page.Canvas.ClientSize);
            textWidget.Draw(page, bounds, textLayout);

            pdf.SaveToFile(path + "\\MailBody.pdf", Spire.Pdf.FileFormat.PDF);

        }

        public static void GetAttachments(MimeMessage message)
        {
            string InitialFolder = ConfigurationManager.AppSettings.Get("InitialFolder");
            string date = message.Date.DateTime.ToString("yyyy'-'MM'-'dd'T'HH'-'mm'-'ss");
            //string path = $"C:\\InitialFolder\\{date}";
            string path = $"{InitialFolder}\\{date}";

            System.IO.Directory.CreateDirectory(path);


            foreach (var attachment in message.Attachments)
            {
                var fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;


                using (var stream = File.Create(path + "\\" + fileName))
                {
                    if (attachment is MessagePart)
                    {
                        var rfc822 = (MessagePart)attachment;

                        rfc822.Message.WriteTo(stream);
                    }
                    else
                    {
                        var part = (MimePart)attachment;

                        part.Content.DecodeTo(stream);
                    }
                }
                if (attachment.ContentType.MimeType.Equals("application/x-zip-compressed"))
                {
                    string zipPath = path + "\\" + fileName;
                    string extractPath = path;
                    ZipFile.ExtractToDirectory(zipPath, extractPath);
                    File.Delete(zipPath);
                }

            }

            CreateMailBodyPDFFile(message.TextBody, path);
            //CreateSignalFile(path);

            string InputClassificationFolder = ConfigurationManager.AppSettings.Get("InputClassificationFolder");
            //Directory.Move(path, $"{InputClassificationFolder}\\{date}");
            MoveFolderForClassification(path, InputClassificationFolder, date);


        }


        public static void DownloadMessages()
        {
            using (var client = new ImapClient(new ProtocolLogger("imap.log")))
            {
                MyWebClient webClient = new MyWebClient();
                try
                {
                    //client.Connect("10.88.99.18", 993, SecureSocketOptions.Auto);
                    client.Connect("imap-akquinet.ogicom.pl", 993, SecureSocketOptions.SslOnConnect);
                    string CredentialPairName = ConfigurationManager.AppSettings.Get("CredentialPairName");
                    client.Authenticate(webClient.GetUsername(CredentialPairName), webClient.GetPassword(CredentialPairName));
                    client.Inbox.Open(MailKit.FolderAccess.ReadWrite);
                    //Console.WriteLine(client.ToString());
                    var subfolder = client.Inbox.GetSubfolder("Do importu");
                    subfolder.Open(MailKit.FolderAccess.ReadWrite);
                    ////
                    var items = subfolder.Fetch(0, -1, MessageSummaryItems.UniqueId | MessageSummaryItems.Size | MessageSummaryItems.Flags);
                    // iterate over all of the messages and fetch them by UID
                    foreach (var item in items)
                    {
                        var message = subfolder.GetMessage(item.UniqueId);
                        GetAttachments(message);
                        subfolder.MoveTo(item.UniqueId, client.Inbox.GetSubfolder("Zaimportowane"));
                    }
                    ////
                }
                catch (Exception ex)
                {
                    Console.WriteLine("There was an error: {0}", ex.Message);
                }

                client.Disconnect(true);
            }
        }
    }
}
