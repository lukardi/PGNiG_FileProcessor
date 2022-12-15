///////Credentials
using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MailKit.Security;
////////Mail Libraries
using MimeKit;
using MsgReader.Outlook;
/////////SPIRE PDF
using Spire.Pdf;
using Spire.Pdf.Graphics;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
////ZIP
using System.IO.Compression;
using System.Linq;
using System.Text;
//
using Microsoft.Win32;
using Spire.Pdf.Exporting.XPS.Schema;

namespace PGNiG_FileProcessor
{
    public class FileGatherer
    {
        const string RegisterKey = "HKEY_CURRENT_USER\\Software\\LukardiSettings";

        public static void Run()
        {
            CollectNetworkFiles();
            DownloadMessages();
            ProcessClassifiedPDFs();
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
                CleanUpLockLibreFiles(path);
            }
            catch (Exception ex)
            {
                Console.WriteLine("There was an error: {0}", ex.Message);
            }

        }

        public static void CollectNetworkFiles()
        {
            string NetworkFolder = ConfigurationManager.AppSettings.Get("NetworkFolder");
            string ProcessedZIPFilesFolder = ConfigurationManager.AppSettings.Get("ProcessedZIPFiles");

            try
            {
                string[] array2 = Directory.GetFiles(NetworkFolder, "*.zip");
                foreach (string zipfile in array2)
                {
                    GetAttachmentsFromZIPFile(zipfile);
                    //File.Delete(zipfile);
                    File.Move(zipfile, ProcessedZIPFilesFolder + @"\" + System.IO.Path.GetFileName(zipfile));

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
            Directory.CreateDirectory(path);
            ZipFile.ExtractToDirectory(zipfile, path);

            string[] array = Directory.GetFiles(path, "*.msg");
            Storage.Message messagefile = new Storage.Message(array[0], FileAccess.ReadWrite);
            string date = messagefile.Headers.DateSent.ToLocalTime().ToString("yyyy'-'MM'-'dd'T'HH'-'mm'-'ss");

            CreateMailBodyPDFFile(messagefile.BodyText, path);
            //CreateSignalFile(path);
            messagefile.Dispose();

            File.Delete(array[0]);
            MoveFolderForClassification(path, InputClassificationFolder, date);
            //Directory.Delete(path, true);

        }

        public static void MoveFolderForClassification(string inputFolder, string outputFolder, string date)
        {
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

        public static string CreateMailBodyPDFFile(string message, string path)
        {
            string file = path + "\\MailBody.pdf";
            string MailBody = message;
            PdfDocument pdf = new PdfDocument();
            PdfPageBase page = pdf.Pages.Add();
            PdfFont font = new PdfFont(PdfFontFamily.TimesRoman, 11);
            PdfTextLayout textLayout = new PdfTextLayout();
            textLayout.Break = PdfLayoutBreakType.FitPage;
            textLayout.Layout = PdfLayoutType.Paginate;
            PdfStringFormat format = new PdfStringFormat();
            //format.Alignment = PdfTextAlignment.Justify;
            format.LineSpacing = 20f;
            PdfTextWidget textWidget = new PdfTextWidget(MailBody, new PdfTrueTypeFont(new Font("Arial", 11), true), PdfBrushes.Black);
            textWidget.StringFormat = format;
            RectangleF bounds = new RectangleF(new PointF(10, 25), page.Canvas.ClientSize);
            textWidget.Draw(page, bounds, textLayout);


            pdf.SaveToFile(file, FileFormat.PDF);
            return file;
        }

        public static void GetAttachments(MimeMessage message, string date, string path)
        {
            foreach (var attachment in message.Attachments)
            {
                var fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;
                using (var stream = File.Create(path + "\\" + fileName))
                {
                    if (attachment is MessagePart rfc822)
                    {
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
        }

        public static void DownloadMessages()
        {
            string InitialFolder = ConfigurationManager.AppSettings.Get("InitialFolder");
            using (var client = new ImapClient(new ProtocolLogger("imap.log")))
            {
                MyWebClient webClient = new MyWebClient();
                try
                {
                    client.Connect("imap.pgnig.pl", 143, SecureSocketOptions.StartTls);
                    string CredentialPairName = ConfigurationManager.AppSettings.Get("CredentialPairName");
                    client.Authenticate(webClient.GetUsername(CredentialPairName), webClient.GetPassword(CredentialPairName));
                    client.Inbox.Open(FolderAccess.ReadWrite);
                    var folder = client.Inbox.GetSubfolders();
                    var subfolder = client.Inbox.GetSubfolder("Do Importu");
                    subfolder.Open(FolderAccess.ReadWrite);
                    var items = subfolder.Fetch(0, -1, MessageSummaryItems.UniqueId | MessageSummaryItems.Size | MessageSummaryItems.Flags);
                    foreach (var item in items)
                    {
                        var message = subfolder.GetMessage(item.UniqueId);
                        string mailPDF = null;
                        string date = message.Date.DateTime.ToString("yyyy'-'MM'-'dd'T'HH'-'mm'-'ss");
                        string path = $"{InitialFolder}\\{date}";
                        try
                        {
                            Directory.CreateDirectory(path);
                            mailPDF = CreateMailBodyPDFFile(message.Date.ToLocalTime() + Environment.NewLine + message.Subject + Environment.NewLine + message.TextBody, path);
                            GetAttachments(message, date, path);
                            string InputClassificationFolder = ConfigurationManager.AppSettings.Get("InputClassificationFolder");
                            MoveFolderForClassification(path, InputClassificationFolder, date);
                        }
                        catch (Exception msgEx)
                        {
                            SendErrorMail(message, mailPDF);
                            Console.WriteLine("There was an error: {0}", msgEx.Message);
                        }
                        finally
                        {
                            if (Directory.Exists(path))
                            {
                                Directory.Delete(path, true);
                            }
                            subfolder.MoveTo(item.UniqueId, client.Inbox.GetSubfolder("Zaimportowane"));
                        }
                    }
                }
                catch (Exception ex)
                {
                    //CreateMailBodyPDFFile(ex.Message, @"C:\test");
                    Console.WriteLine("There was an error: {0}", ex.Message);
                }

                client.Disconnect(true);
            }
        }

        public static void SendErrorMail(MimeMessage originalMessage, string mailPDF = null)
        {
            using (var client = new SmtpClient(new ProtocolLogger("smtp.log")))
            {
                MyWebClient webClient = new MyWebClient();
                try
                {
                    client.Connect(ConfigurationManager.AppSettings.Get("SMTPServer"), 25, SecureSocketOptions.Auto);
                    string CredentialPairName = ConfigurationManager.AppSettings.Get("CredentialPairName");
                    client.Authenticate(webClient.GetUsername(CredentialPairName), webClient.GetPassword(CredentialPairName));

                    var message = new MimeMessage();
                    message.From.Add(new MailboxAddress("Service Error", "EfakturaODtest@pgnig.pl"));
                    //
                    var emails = ConfigurationManager.AppSettings.Get("ErrorMailReceivers").Split(';').ToList();
                    emails.RemoveAll(s => string.IsNullOrEmpty(s));
                    foreach (var email in emails)
                    {
                        message.To.Add(new MailboxAddress(email, email));
                    }
                    message.Subject = "[Wystąpił błąd] " + originalMessage.Subject;
                    message.Body = new TextPart("plain")
                    {
                        Text = "Drogi Użytkowniku, faktura z załączonego maila nie została poprawnie przetworzona."
                        + Environment.NewLine
                        + "Proszę o ponowne podjęcie załącznika z maila i ponowne wprowadzenie do systemu."
                        + Environment.NewLine
                        + "* to jest powiadomienie systemowe, proszę na nie odpowiadać"
                    };
                    if (mailPDF != null)
                    {
                        var attachment = new MimePart("application", "pdf")
                        {
                            Content = new MimeContent(File.OpenRead(mailPDF)),
                            ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                            ContentTransferEncoding = ContentEncoding.Base64,
                            FileName = System.IO.Path.GetFileName(mailPDF)
                        };
                        message.Body = new Multipart("mixed")
                        {
                            message.Body,
                            attachment
                        };
                    }
                    client.Send(message);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("There was an error: {0}", ex.Message);
                }
                finally
                {
                    client.Disconnect(true);
                }
            }
        }

        public static void ReturnMessageToINBOX(IMessageSummary msgsummary)
        {

            using (var client = new ImapClient(new ProtocolLogger("imap.log")))
            {
                MyWebClient webClient = new MyWebClient();
                try
                {
                    client.Connect("imap.pgnig.pl", 143, SecureSocketOptions.StartTls);

                    string CredentialPairName = ConfigurationManager.AppSettings.Get("CredentialPairName");
                    client.Authenticate(webClient.GetUsername(CredentialPairName), webClient.GetPassword(CredentialPairName));
                    client.Inbox.Open(MailKit.FolderAccess.ReadWrite);
                    var folder = client.Inbox.GetSubfolders();

                    var subfolder = client.Inbox.GetSubfolder("Do Importu");
                    subfolder.Open(MailKit.FolderAccess.ReadWrite);

                    subfolder.MoveTo(msgsummary.UniqueId, client.Inbox);
                    ////
                }
                catch (Exception ex)
                {
                    //CreateMailBodyPDFFile(ex.Message, @"C:\test");
                    Console.WriteLine("There was an error: {0}", ex.Message);
                }

                client.Disconnect(true);
            }
        }

        public static void CleanUpLockLibreFiles(string path)
        {
            string[] extensions = { ".docx#", ".doc#", ".xls#", ".xlsx#", ".rtf#", ".ods#", ".odt#", ".docx", ".doc", ".xls", ".xlsx", ".rtf", ".ods", ".odt" };
            string[] officefiles = Directory.GetFiles(path, "*.*")
                .Where(f => extensions.Contains(new FileInfo(f).Extension.ToLower())).ToArray();

            foreach (string file in officefiles)
            {
                File.Delete(file);
            }
        }

        public static void ProcessClassifiedPDFs()
        {
            List<string> FVs = new List<string>();
            List<string> Attachments = new List<string>();
            string OutputClassificationFolder = ConfigurationManager.AppSettings.Get("OutputClassificationFolder");
            string[] pdfarray = Directory.GetFiles(OutputClassificationFolder, "*.pdf");

            foreach (string pdf in pdfarray)
            {
                if (pdf.Contains("F_"))
                {
                    FVs.Add(pdf);
                }
                else if (pdf.Contains("Z_"))
                {
                    Attachments.Add(pdf);
                }
            }
            foreach (string FV in FVs)
            {
                MergePDF(FV, Attachments);
            }
            CleanupOutputClassificationFolder(pdfarray);
        }

        public static void CleanupOutputClassificationFolder(string[] PDFs)
        {
            foreach (string pdf in PDFs)
            {
                File.Delete(pdf);
            }
        }

        public static void MergePDF(string FV, List<string> Attachments)
        {
            string FinalFVsDirectory = ConfigurationManager.AppSettings.Get("CompleteFVs");

            List<string> PDF_To_Merge = new List<string>();
            string bid = FV.Substring(GetNthIndex(FV, '_', 1) + 1);
            bid = bid.Split('_')[0];
            PDF_To_Merge.Add(FV);
            ///////////////////////
            foreach (string attachment in Attachments)
            {
                if (attachment.Contains(bid))
                {
                    PDF_To_Merge.Add(attachment);
                }
            }

            PdfDocumentBase mergedPDF = PdfDocument.MergeFiles(PDF_To_Merge.ToArray());
            string filename = System.IO.Path.GetFileName(FV);
            mergedPDF.Save(FinalFVsDirectory + filename);

            int year = DateTime.Today.Year;
            string FVnr = $"E{year}5" + GetNextFVNumber().ToString("D6");
            ///Start Barcoder
            System.Diagnostics.Process.Start(ConfigurationManager.AppSettings.Get("Barcoder"), $"\"{FinalFVsDirectory}{filename}\" {FVnr} barcode \"{FinalFVsDirectory}{filename}");

        }

        public static long GetNextFVNumber()
        {
            long FVNumber;
            RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\LukardiSettings");
            if (key != null)
            {
                if (key.GetValue("CurrentFVNumber") != null)
                {
                    FVNumber = (long)key.GetValue("CurrentFVNumber");
                    FVNumber++;
                    Registry.SetValue(RegisterKey, "CurrentFVNumber", FVNumber, RegistryValueKind.QWord);
                }
                else
                {
                    Registry.SetValue(RegisterKey, "CurrentFVNumber", 1, RegistryValueKind.QWord);
                    FVNumber = (long)key.GetValue("CurrentFVNumber");
                }
                key.Close();
            }
            else
            {
                Registry.CurrentUser.CreateSubKey(@"SOFTWARE\LukardiSettings");
                return 1;
            }
            return FVNumber;
        }

        public static int GetNthIndex(string s, char t, int n)
        {
            int count = 0;
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] == t)
                {
                    count++;
                    if (count == n)
                    {
                        return i;
                    }
                }
            }
            return -1;
        }

    }
}
