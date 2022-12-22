using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MailKit.Security;
//
using MimeKit;
using MsgReader.Outlook;
//
using Spire.Pdf;
using Spire.Pdf.Graphics;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
//
using System.IO.Compression;
using System.Linq;
using System.Text;
//
using Microsoft.Win32;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace PGNiG_FileProcessor
{
    public class FileGatherer
    {
        private static string RegisterKey = "SOFTWARE\\LukardiSettings";
        private static string RegisterValueKey;
        //
        private static string SourceInboxFolderName;
        private static string DestinationInboxFolderName;
        //
        private static string SMTPServer;
        private static int SMTPPort;
        private static string IMAPServer;
        private static int IMAPPort;
        //
        private static string NetworkFolder;
        private static string ProcessedZIPFiles;
        private static string ErrorZIPFiles;
        private static string InputClassificationFolder;
        private static string InitialFolder;
        private static string OutputClassificationFolder;
        private static string CompleteFVs;
        private static string TmpCompleteFVs;
        //
        private static string Barcoder;
        //
        private static string CredentialPairName;
        private static List<string> ErrorMailReceivers;
        private static bool SendErrors = false;
        //
        private static Regex bidRegex = new Regex(@"_bid(\d+)_");
        //
        private static string imapLog = "imap.log";
        private static string smtpLog = "smtp.log";
        //
        private static string MailBodyFileName = "MailBody.pdf";
        private static string SignalFileName = "SignalFile.xml";
        private static string dateFormat = "yyyy-MM-ddTHH-mm-ss";

        private static readonly string[] extensions = {
                ".docx#",
                ".doc#",
                ".xls#",
                ".xlsx#",
                ".rtf#",
                ".ods#",
                ".odt#", 
                //
                ".docx",
                ".doc",
                ".xls",
                ".xlsx",
                ".rtf",
                ".ods",
                ".odt"
            };

        public static void Init()
        {
            RegisterValueKey = ConfigurationManager.AppSettings.Get("RegisterValueKey");
            //
            SourceInboxFolderName = ConfigurationManager.AppSettings.Get("SourceFolderName");
            DestinationInboxFolderName = ConfigurationManager.AppSettings.Get("DestinationFolderName");
            //
            SMTPServer = ConfigurationManager.AppSettings.Get("SMTPServer");
            SMTPPort = int.Parse(ConfigurationManager.AppSettings.Get("SMTPPort"));
            IMAPServer = ConfigurationManager.AppSettings.Get("IMAPServer");
            IMAPPort = int.Parse(ConfigurationManager.AppSettings.Get("IMAPPort"));
            //
            NetworkFolder = ConfigurationManager.AppSettings.Get("NetworkFolder");
            ProcessedZIPFiles = ConfigurationManager.AppSettings.Get("ProcessedZIPFiles");
            ErrorZIPFiles = ConfigurationManager.AppSettings.Get("ErrorZIPFiles");
            InputClassificationFolder = ConfigurationManager.AppSettings.Get("InputClassificationFolder");
            InitialFolder = ConfigurationManager.AppSettings.Get("InitialFolder");
            OutputClassificationFolder = ConfigurationManager.AppSettings.Get("OutputClassificationFolder");
            CompleteFVs = ConfigurationManager.AppSettings.Get("CompleteFVs");
            TmpCompleteFVs = Path.Combine(CompleteFVs, "temp");
            //
            Barcoder = ConfigurationManager.AppSettings.Get("Barcoder");
            //
            CredentialPairName = ConfigurationManager.AppSettings.Get("CredentialPairName");
            ErrorMailReceivers = ConfigurationManager.AppSettings.Get("ErrorMailReceivers").Split(';').ToList();
            SendErrors = ConfigurationManager.AppSettings.Get("SendErrors") == "1";
            //
            imapLog = Logger.GenerateFile("imap");
            smtpLog = Logger.GenerateFile("smtp");
            foreach (var dir in new List<string>() {
                ProcessedZIPFiles,
                InputClassificationFolder,
                InitialFolder,
                OutputClassificationFolder,
                CompleteFVs,
                TmpCompleteFVs,
                ErrorZIPFiles
            })
            {
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
            }
        }

        public static void Run()
        {
            try
            {
                Logger.Debug("Classify pdfs...");
                ProcessClassifiedPDFs();
                Logger.Debug("Check network files...");
                CollectNetworkFiles();
                Logger.Debug("Download messages from inbox...");
                DownloadMessages();
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
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
                    Logger.Debug($"Convert to pdf: {file}");
                    LibreOfficeConverter.Run(file, path);
                }
                CleanUpLockLibreFiles(path);
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }

        }

        public static void CollectNetworkFiles()
        {
            foreach (string file in Directory.GetFiles(NetworkFolder, "*.zip"))
            {
                string destPath = null;
                try
                {
                    destPath = GetAttachmentsFromZIPFile(file);
                    MoveFileSafe(file, ProcessedZIPFiles);
                }
                catch (Exception ex)
                {
                    Logger.Error(ex);
                    //
                    string mailBodyFile = null;
                    foreach (var tmpPath in new string[] { Path.Combine(InitialFolder, "temp"), destPath })
                    {
                        var mailFile = tmpPath != null ? Path.Combine(tmpPath, MailBodyFileName) : null;
                        if (mailFile != null && File.Exists(mailFile))
                        {
                            mailBodyFile = mailFile;
                        }
                    }
                    SendErrorMail(new MimeMessage()
                    {
                        Subject = $"Archiwum {file}"
                    }, mailBodyFile);
                    //
                    MoveFileSafe(file, ErrorZIPFiles);
                }
            }
        }

        public static string GetAttachmentsFromZIPFile(string zipfile)
        {
            Logger.Debug($"Extracting file: {zipfile}");
            string path = Path.Combine(InitialFolder, "temp");
            if (!Directory.Exists(path))
            {
                Logger.Debug($"Creating temp directory: {path}");
                Directory.CreateDirectory(path);
            }
            else
            {
                Logger.Debug($"Cleanup temp directory: {path}");
                foreach (var item in Directory.GetDirectories(path))
                {
                    Directory.Delete(item, true);
                }
                foreach (var item in Directory.GetFiles(path))
                {
                    File.Delete(item);
                }
            }
            using (var archive = ZipFile.Open(zipfile, ZipArchiveMode.Read))
            {
                archive.ExtractToDirectory(path);
            }
            string[] files = Directory.GetFiles(path, "*.msg");
            if (files.Length <= 0)
            {
                throw new Exception($"Couldn't find msg file in directory: {path}");
            }
            string msgFile = files.First();
            Logger.Debug($"Found mail message file: {msgFile}");
            string date;
            using (Storage.Message messagefile = new Storage.Message(msgFile, FileAccess.ReadWrite))
            {
                date = messagefile.Headers.DateSent.ToLocalTime().ToString(dateFormat);
                Logger.Debug($"Fetched date: {date}");
                CreateMailBodyPDFFile(messagefile.BodyText, path);
            }
            Logger.Debug($"Removing message file: {msgFile}");
            File.Delete(msgFile);
            var destPath = Path.Combine(InputClassificationFolder, date);
            MoveFolderForClassification(path, ref destPath);
            return destPath;
        }

        public static void MoveFolderForClassification(string inputFolder, ref string outputFolder)
        {
            outputFolder = MoveDirectorySafe(inputFolder, outputFolder);
            ConvertOfficeFiles(outputFolder);
            CreateSignalFile(outputFolder);
        }

        public static void CreateSignalFile(string path)
        {
            Logger.Debug($"Creating singal file in: {path}");
            using (FileStream fs = File.Create(Path.Combine(path, SignalFileName)))
            {
                byte[] author = new UTF8Encoding(true).GetBytes("1");
                fs.Write(author, 0, author.Length);
            }
        }

        public static string CreateMailBodyPDFFile(string message, string path)
        {
            string file = Path.Combine(path, MailBodyFileName);
            Logger.Debug($"Generating mail body file: {file}");
            string MailBody = message;
            PdfDocument pdf = new PdfDocument();
            PdfPageBase page = pdf.Pages.Add();
            PdfTextLayout textLayout = new PdfTextLayout
            {
                Break = PdfLayoutBreakType.FitPage,
                Layout = PdfLayoutType.Paginate
            };
            PdfStringFormat format = new PdfStringFormat
            {
                LineSpacing = 20f
            };
            PdfTextWidget textWidget = new PdfTextWidget(MailBody, new PdfTrueTypeFont(new Font("Arial", 11), true), PdfBrushes.Black)
            {
                StringFormat = format
            };
            RectangleF bounds = new RectangleF(new PointF(10, 25), page.Canvas.ClientSize);
            textWidget.Draw(page, bounds, textLayout);
            Logger.Debug($"Save mail to pdf: {file}");
            pdf.SaveToFile(file, FileFormat.PDF);
            if (!File.Exists(file))
            {
                throw new Exception($"Couldn't save mail body to pdf: {file}");
            }
            return file;
        }

        public static void GetAttachments(MimeMessage message, string path)
        {
            foreach (var attachment in message.Attachments)
            {
                var fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;
                var file = Path.Combine(path, fileName);
                using (var stream = File.Create(file))
                {
                    Logger.Debug($"Saving attachement file: {file}");
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
                    Logger.Debug($"Attachement iz zip so extract it: {file} to {path}");
                    ZipFile.ExtractToDirectory(file, path);
                    Logger.Debug("Delete source zip file");
                    File.Delete(file);
                }

            }
        }

        public static void DownloadMessages()
        {
            using (var client = new ImapClient(new ProtocolLogger(imapLog)))
            {
                MyWebClient webClient = new MyWebClient();
                try
                {
                    Logger.Debug($"Connecting to IMAP server: {IMAPServer}:{IMAPPort}");
                    client.Connect(IMAPServer, IMAPPort, SecureSocketOptions.Auto);
                    client.Authenticate(webClient.GetUsername(CredentialPairName), webClient.GetPassword(CredentialPairName));
                    client.Inbox.Open(FolderAccess.ReadWrite);
                    var folder = client.Inbox.GetSubfolders();
                    var subfolder = client.Inbox.GetSubfolder(SourceInboxFolderName);
                    subfolder.Open(FolderAccess.ReadWrite);
                    var items = subfolder.Fetch(0, -1, MessageSummaryItems.UniqueId | MessageSummaryItems.Size | MessageSummaryItems.Flags);
                    Logger.Debug($"Found {items.Count} mails");
                    foreach (var item in items)
                    {
                        Logger.Debug($"Processing mail with UID: {item.UniqueId}");
                        var message = subfolder.GetMessage(item.UniqueId);
                        string mailPDF = null;
                        string date = message.Date.DateTime.ToString(dateFormat);
                        string path = Path.Combine(InitialFolder, date);
                        try
                        {
                            Directory.CreateDirectory(path);
                            mailPDF = CreateMailBodyPDFFile(message.Date.ToLocalTime() + Environment.NewLine + message.Subject + Environment.NewLine + message.TextBody, path);
                            GetAttachments(message, path);
                            var destPath = Path.Combine(InputClassificationFolder, date);
                            MoveFolderForClassification(path, ref destPath);
                        }
                        catch (Exception msgEx)
                        {
                            Logger.Error(msgEx);
                            SendErrorMail(message, mailPDF);
                        }
                        finally
                        {
                            if (Directory.Exists(path))
                            {
                                Directory.Delete(path, true);
                            }
                            subfolder.MoveTo(item.UniqueId, client.Inbox.GetSubfolder(DestinationInboxFolderName));
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error(ex);
                }

                client.Disconnect(true);
            }
        }

        public static void SendErrorMail(MimeMessage originalMessage, string mailPDF = null)
        {
            if (!SendErrors)
            {
                Logger.Debug($"Sending errors is disabled. {originalMessage.Subject}");
                return;
            }
            using (var client = new SmtpClient(new ProtocolLogger(smtpLog)))
            {
                MyWebClient webClient = new MyWebClient();
                try
                {
                    Logger.Debug($"Sending error mail with SMTP server: {SMTPServer}:{SMTPPort}");
                    client.Connect(SMTPServer, SMTPPort, SecureSocketOptions.Auto);
                    client.Authenticate(webClient.GetUsername(CredentialPairName), webClient.GetPassword(CredentialPairName));

                    var message = new MimeMessage();
                    message.From.Add(new MailboxAddress("Service Error", "EfakturaODtest@pgnig.pl"));
                    Logger.Debug($"Recipients: {ErrorMailReceivers}");
                    foreach (var email in ErrorMailReceivers)
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
                            FileName = Path.GetFileName(mailPDF)
                        };
                        Logger.Debug($"Attaching file: {mailPDF}");
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
                    Logger.Error(ex);
                }
                finally
                {
                    client.Disconnect(true);
                }
            }
        }

        public static void ReturnMessageToINBOX(IMessageSummary msgsummary)
        {

            using (var client = new ImapClient(new ProtocolLogger(imapLog)))
            {
                MyWebClient webClient = new MyWebClient();
                try
                {
                    Logger.Debug($"Move back message with UID: {msgsummary.UniqueId}");
                    client.Connect(IMAPServer, IMAPPort, SecureSocketOptions.StartTls);
                    client.Authenticate(webClient.GetUsername(CredentialPairName), webClient.GetPassword(CredentialPairName));
                    client.Inbox.Open(FolderAccess.ReadWrite);
                    var folder = client.Inbox.GetSubfolders();
                    var subfolder = client.Inbox.GetSubfolder(SourceInboxFolderName);
                    subfolder.Open(FolderAccess.ReadWrite);
                    subfolder.MoveTo(msgsummary.UniqueId, client.Inbox);
                }
                catch (Exception ex)
                {
                    Logger.Error(ex);
                }

                client.Disconnect(true);
            }
        }

        public static void CleanUpLockLibreFiles(string path)
        {
            string[] officefiles = Directory.GetFiles(path, "*.*")
                .Where(f => extensions.Contains(new FileInfo(f).Extension.ToLower())).ToArray();

            foreach (string file in officefiles)
            {
                Logger.Debug($"Removing lock file: {file}");
                File.Delete(file);
            }
        }

        public static string GetBid(string name)
        {
            var matches = bidRegex.Matches(name);
            if (matches.Count > 0 && matches[0].Success)
            {

                return matches[0].Groups[0].Value;
            }
            return null;
        }

        public static void ProcessClassifiedPDFs()
        {
            var invoices = new Dictionary<string, List<string>>();
            var attachments = new Dictionary<string, List<string>>();
            foreach (string pdf in Directory.GetFiles(OutputClassificationFolder, "*.pdf"))
            {
                var bid = GetBid(pdf);
                if (bid == null || bid == "")
                {
                    continue;
                }
                Logger.Debug($"BID: {bid}");
                var filename = Path.GetFileName(pdf);
                if (filename.StartsWith("F_"))
                {
                    Logger.Debug($"Found invoice: {pdf}");
                    if (!invoices.ContainsKey(bid))
                    {
                        invoices.Add(bid, new List<string>());
                    }
                    invoices[bid].Add(pdf);
                }
                else if (filename.StartsWith("Z_"))
                {
                    Logger.Debug($"Found attachement: {pdf}");
                    if (!attachments.ContainsKey(bid))
                    {
                        attachments.Add(bid, new List<string>());
                    }
                    attachments[bid].Add(pdf);
                }
                else
                {
                    Logger.Debug($"Unknown file type: {pdf}");
                }
            }
            var toRemove = new List<string>();
            foreach (string bid in invoices.Keys)
            {
                foreach (string file in invoices[bid])
                {
                    try
                    {
                        var files = attachments.ContainsKey(bid) ? new List<string>(attachments[bid]) : new List<string>();
                        files.Insert(0, file);
                        MergePDF(file, files);
                        toRemove.AddRange(files);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex);
                    }
                }
            }
            CleanupOutputClassificationFolder(toRemove.Distinct().ToArray());
        }

        public static void CleanupOutputClassificationFolder(string[] PDFs)
        {
            foreach (string pdf in PDFs)
            {
                Logger.Debug($"Cleaning: {pdf}");
                File.Delete(pdf);
            }
        }

        public static void MergePDF(string FV, List<string> Attachments)
        {
            Logger.Debug($"Merging files for: {FV}");
            Logger.Debug($"Files to merge: {Environment.NewLine}- {string.Join($"{Environment.NewLine}- ", Attachments)}");
            PdfDocumentBase mergedPDF = PdfDocument.MergeFiles(Attachments.ToArray());
            string filename = Path.GetFileName(FV);
            var file = Path.Combine(TmpCompleteFVs, filename);
            mergedPDF.Save(file);
            if (!File.Exists(file))
            {
                throw new Exception($"Coludn't merge pdfs. {file}");
            }
            int year = DateTime.Today.Year;
            string FVnr = $"E{year}5" + GetNextFVNumber().ToString("D6");
            var barcodeFile = Path.Combine(CompleteFVs, filename);
            if (!RunBarcode(file, FVnr, barcodeFile))
            {
                throw new Exception("Something went wrong with barcode");
            }
            File.Delete(file);
        }

        public static bool RunBarcode(string file, string number, string outputFile)
        {
            Logger.Debug($"Running barcode for file: {file}, number: {number}, output: {outputFile}");
            using (var process = Process.Start(Barcoder, "\"" + string.Join("\" \"", new List<string>() {
                file,
                number,
                "barcode",
                outputFile
            }) + "\""))
            {
                process.WaitForExit();
                Logger.Debug($"Barcode exit code: {process.ExitCode}");
                return process.ExitCode == 0;
            }
        }

        public static long GetNextFVNumber()
        {
            long FVNumber;
            RegistryKey key = Registry.CurrentUser.OpenSubKey(RegisterKey);
            if (key != null)
            {
                if (key.GetValue(RegisterValueKey) != null)
                {
                    Logger.Debug($"Found value for key: {RegisterValueKey}");
                    FVNumber = (long)key.GetValue(RegisterValueKey);
                    FVNumber++;
                    Registry.SetValue($"HKEY_CURRENT_USER\\{RegisterKey}", RegisterValueKey, FVNumber, RegistryValueKind.QWord);
                }
                else
                {
                    Logger.Debug("Key doesn't exists.");
                    Registry.SetValue($"HKEY_CURRENT_USER\\{RegisterKey}", RegisterValueKey, 1, RegistryValueKind.QWord);
                    FVNumber = (long)key.GetValue(RegisterValueKey);
                }
                key.Close();
            }
            else
            {
                Registry.CurrentUser.CreateSubKey(RegisterKey);
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

        public static string MoveFileSafe(string file, string destPath)
        {
            string destFile;
            var time = (int)DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1)).TotalSeconds;
            do
            {
                destFile = Path.Combine(destPath, Path.GetFileNameWithoutExtension(file) + $"-{time++}{Path.GetExtension(file)}");
            }
            while (File.Exists(destFile));
            Logger.Debug($"Safe moving file: {file} -> {destFile}");
            File.Move(file, destFile);
            if (!File.Exists(destFile))
            {
                throw new Exception("Safe move file failed!");
            }
            return destFile;
        }

        public static string MoveDirectorySafe(string inputFolder, string outputFolder)
        {
            var n = 0;
            string tmpOutputFolder;
            do
            {
                tmpOutputFolder = $"{outputFolder}-{n++:D5}";
            } while (Directory.Exists(tmpOutputFolder));
            Logger.Debug($"Safe moving directory: {inputFolder} -> {tmpOutputFolder}");
            Directory.Move(inputFolder, tmpOutputFolder);
            if (!Directory.Exists(tmpOutputFolder))
            {
                throw new Exception("Safe move directory failed!");
            }
            return tmpOutputFolder;
        }
    }
}
