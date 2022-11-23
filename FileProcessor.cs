using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace PGNiG_FileProcessor
{
    public partial class FileProcessor : ServiceBase
    {
        public FileProcessor()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
           // CheckFolders();

            Timer timer = new Timer();
            //timer.Interval = 120000; // 120 seconds
            timer.Interval = 30000; // 120 seconds
            timer.Elapsed += new ElapsedEventHandler(this.OnTimer);
            timer.Start();
        }

        public void OnTimer(object sender, ElapsedEventArgs args)
        {
            FileGatherer.CollectNetworkFiles();
            FileGatherer.DownloadMessages();
           
        }

        protected override void OnStop()
        {
            
        }

      
    }
}
