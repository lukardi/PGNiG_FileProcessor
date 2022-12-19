using System.Configuration;
using System.ServiceProcess;
using System.Timers;

namespace PGNiG_FileProcessor
{
    public partial class FileProcessor : ServiceBase
    {
        public FileProcessor()
        {
            InitializeComponent();
        }

        public void OnStartTest(string[] args)
        {
            OnStart(args);
        }

        protected override void OnStart(string[] args)
        {
            Timer timer = new Timer
            {
                Interval = int.Parse(ConfigurationManager.AppSettings.Get("TimerInterval")) * 1000
            };
            timer.Elapsed += new ElapsedEventHandler((object sender, ElapsedEventArgs timerArgs) =>
            {
                timer.Stop();
                OnTimer();
                timer.Start();
            });
            timer.Start();
        }

        public void OnTimer()
        {
            FileGatherer.Run();
        }

        protected override void OnStop()
        {

        }
    }
}
