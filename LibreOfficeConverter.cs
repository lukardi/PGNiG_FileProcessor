using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Configuration;
using System.IO;

namespace PGNiG_FileProcessor
{
    class LibreOfficeConverter
    {
        private static readonly string appPath = AppDomain.CurrentDomain.BaseDirectory;
        //private static readonly string exePath = Properties.Settings.Default.LibreOfficePath + @"\program\soffice.exe";
        private static readonly string exePath = ConfigurationManager.AppSettings.Get("LibreOfficePath") + @"\program\soffice.exe";
        // private static readonly string userData = appPath + @"UserData";
        private static readonly string userData = ConfigurationManager.AppSettings.Get("UserDataFolder");

        public static void Run(string inputPath, string outputPath)
        {
            CheckUserData();
            if (!Directory.Exists(outputPath))
            {
                Directory.CreateDirectory(outputPath);
            }
            Process(inputPath, outputPath);
        }

        public static void CheckUserData()
        {
            if (Directory.Exists(userData))
            {
                return;
            }
            Directory.CreateDirectory(userData);
            Exec(new List<string> {
                    "--terminate_after_init"
                });
            File.Copy(appPath + @"Macro\Module1.xba", userData + @"\user\basic\Standard\Module1.xba", true);
        }

        public static void Process(string inputPath, string outputPath)
        {
            if (!Directory.Exists(inputPath))
            {
                if (!File.Exists(inputPath))
                {
                    throw new Exception("Input file does not exist");
                }
                string outputFile = $"{outputPath}\\{Path.GetFileNameWithoutExtension(inputPath)}.pdf";
                if (File.Exists(outputFile))
                {
                    File.Delete(outputFile);
                }
            }
            else
            {
                inputPath = Path.Combine(inputPath, "*");
            }
            //Run macro to set page scale
            Exec(new List<string> {
                $"\"{inputPath}\"",
                "\"macro:///Standard.Module1.FitToPage\"",
            });
            //Run conversion
            Exec(new List<string> {
                "--convert-to pdf",
                $"--outdir \"{outputPath}\"",
                $"\"{inputPath}\""
            });
        }

        public static int Exec(List<string> args)
        {
            var process = new Process();
            process.StartInfo.FileName = exePath;
            args.Add("--headless");
            args.Add("--norestore");
            args.Add("--nofirststartwizard");
            var userDataDir = $"file:///{userData.Replace(@"\", @"/")}";
            args.Add($"-env:UserInstallation=\"{userDataDir}\"");
            process.StartInfo.Arguments = string.Join(" ", args);
            process.StartInfo.UseShellExecute = true;
            process.StartInfo.CreateNoWindow = true;
            process.Start();
            if (!process.WaitForExit(60 * 1000))
            {
                process.Kill();
            }
            return process.ExitCode;
        }
    }
}
