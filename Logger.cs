using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;

namespace PGNiG_FileProcessor
{
    class Logger
    {

        public const int TYPE_DEBUG = 0;
        public const int TYPE_INFO = 10;
        public const int TYPE_ERROR = 20;

        public static bool console = false;
        public static string logFile;
        public static int SaveLevel = 0;
        public static string LogsPath = ConfigurationManager.AppSettings.Get("LogsPath");

        private static string path;

        public static void Init()
        {
            path = Path.GetFullPath(LogsPath);
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            NewRun();
        }

        public static void Debug(string content)
        {
            Write(TYPE_DEBUG, content);
        }

        public static void Info(string content)
        {
            Write(TYPE_INFO, content);
        }

        public static void Error(string content)
        {
            Write(TYPE_ERROR, content);
        }

        public static void Error(Exception ex)
        {
            Write(TYPE_ERROR, ex.Message + Environment.NewLine + ex.StackTrace);
        }

        public static string GetTypeLabel(int type)
        {
            if (type == TYPE_DEBUG)
            {
                return "debug";
            }
            else if (type == TYPE_INFO)
            {
                return "info";
            }
            else if (type == TYPE_ERROR)
            {
                return "error";
            }
            return "";
        }

        public static void Write(int type, string content)
        {
            if (type < SaveLevel)
            {
                return;
            }
            string message = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.ffffff}][{GetTypeLabel(type).ToUpper()}] {content}";
            if (console)
            {
                Console.WriteLine(message);
            }
            WriteToFile(message);
        }

        public static void WriteToFile(string message)
        {
            File.AppendAllText(logFile, message + Environment.NewLine);
        }

        public static string GenerateFileName(string name = "app")
        {
            return $"{name}_pid{Process.GetCurrentProcess().Id}_{DateTime.Now:yyyyMMdd_HHmmss}.log";
        }

        public static string GenerateFile(string name = "app")
        {
            var file = Path.Combine(path, GenerateFileName(name));
            File.Create(file).Close();
            return file;
        }

        public static void NewRun()
        {
            logFile = GenerateFile();
        }
    }
}
