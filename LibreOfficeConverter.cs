using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Configuration;
using System.IO;

namespace PGNiG_FileProcessor
{
    class LibreOfficeConverter
    {
        private static readonly string appPath = AppDomain.CurrentDomain.BaseDirectory;
        private static readonly string exePath = ConfigurationManager.AppSettings.Get("LibreOfficePath") + @"\program\soffice.exe";
        private static readonly string userData = ConfigurationManager.AppSettings.Get("UserDataFolder");
        public static readonly bool headless = ConfigurationManager.AppSettings.Get("LibreOfficePathHeadlessMode") == "true";
        private static readonly int timeout = 60 * 1000;

        /// <summary>
        /// Runs as background task.
        /// </summary>
        /// <param name="inputPath"></param>
        /// <param name="outputPath"></param>
        /// <returns></returns>
        public static Task RunAsync(string inputPath, string outputPath)
        {
            return Task.Run(() =>
            {
                Run(inputPath, outputPath);
            });
        }

        /// <summary>
        /// Converts a file using LibreOffice.
        /// </summary>
        /// <param name="inputPath"></param>
        /// <param name="outputPath"></param>
        /// <returns>Exec Exit code.</returns>
        public static int Run(string inputPath, string outputPath)
        {
            CheckUserData();
            if (!Directory.Exists(outputPath))
            {
                Logger.Debug($"Creating directory {outputPath}");
                Directory.CreateDirectory(outputPath);
            }
            return Process(inputPath, outputPath);
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

        public static int Process(string inputPath, string outputPath)
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
            Logger.Debug($"Macro for: {inputPath}");
            //Run macro to set page scale
            Exec(new List<string> {
                $"\"{inputPath}\"",
                "\"macro:///Standard.Module1.FitToPage\"",
            });
            Logger.Debug($"Convert for: {inputPath}");
            //Run conversion
            return Exec(new List<string> {
                "--convert-to pdf",
                $"--outdir \"{outputPath}\"",
                $"\"{inputPath}\""
            });
        }

        public static int Exec(List<string> args)
        {
            var process = new Process();
            process.StartInfo.FileName = exePath;
            if (headless)
            {
                args.Add("--headless");
            }
            args.Add("--norestore");
            args.Add("--nofirststartwizard");
            var userDataDir = $"file:///{userData.Replace(@"\", @"/")}";
            args.Add($"-env:UserInstallation=\"{userDataDir}\"");
            process.StartInfo.Arguments = string.Join(" ", args);
            process.StartInfo.UseShellExecute = headless;
            process.StartInfo.CreateNoWindow = headless;
            process.Start();
            if (!process.WaitForExit(timeout) && !process.HasExited)
            {
                if (headless)
                {
                    Logger.Info($"Killing LibreOffice process: {process.Id}");
                    process.Kill();
                }
                else
                {
                    process.WaitForExit();
                }
            }
            Logger.Debug($"LibreOffice exit code: {process.ExitCode}");
            return process.ExitCode;
        }
    }
}
