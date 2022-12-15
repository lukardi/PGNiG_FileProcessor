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
            Spire.License.LicenseProvider.SetLicenseFileName(@"C:\FileGathererFiles\license.elic");
            Timer timer = new Timer
            {
                Interval = 30 * 1000 // 30 seconds
            };
            timer.Elapsed += new ElapsedEventHandler(OnTimer);
            timer.Start();
        }

        public void OnTimer(object sender, ElapsedEventArgs args)
        {
            FileGatherer.Run();
        }

        protected override void OnStop()
        {

        }
    }
}
