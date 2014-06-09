using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using Mgmt.SGM;

namespace SGM
{
    public partial class SGMsvc : ServiceBase
    {
        private SGMService sgmService = null;

        public SGMsvc()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            // Create Thread
            Thread mThread = new Thread(new ThreadStart(delegate
            {
                fireItUp();
            })
            );

            // Start Thread
            mThread.Start();
        }

        protected override void OnStop()
        {
            //sgmService.Stop();
            ExitCode = 0;
        }

        private void fireItUp()
        {
            while (true)
            {
                int ThreadSleep;
                using (sgmService = new SGMService())
                {
                    sgmService.Start();
                    // Convert Seconds to Milliseconds
                    ThreadSleep = Convert.ToInt16(sgmService.Get("interval")) * 1000;
                }
                // Put the thread to sleep
                Thread.Sleep(ThreadSleep);
            }
        }
    }
}
