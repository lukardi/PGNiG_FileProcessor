using System;
using System.ServiceProcess;

namespace PGNiG_FileProcessor
{
    static class Program
    {

        /// <summary>
        /// Główny punkt wejścia dla aplikacji.
        /// </summary>
        static void Main(string[] args)
        {
            if (Environment.UserInteractive)
            {
                RunTestService(args);
                Console.WriteLine("Press any key to stop...");
                Console.ReadKey();
            }
            else
            {
                RunService(args);
            }
        }

        /// <summary>
        /// Uruchamia usługę systemową.
        /// </summary>
        /// <param name="args"></param>
        static void RunService(string[] args)
        {

            ServiceBase[] service = new ServiceBase[]
            {
                new FileProcessor()
            };
            ServiceBase.Run(service);
        }

        /// <summary>
        /// Uruchamia usługę systemową w konsoli.
        /// </summary>
        /// <param name="args"></param>
        static void RunTestService(string[] args)
        {
            FileProcessor service = new FileProcessor();
            service.OnStartTest(args);
        }
    }
}
