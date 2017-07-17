using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;

namespace MailMergeService
{
    public partial class MergeService : ServiceBase
    {
        private System.Threading.Thread workerThread;
        private bool serviceStarted = false;
        int intervalSec=60;
     
        public MergeService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            try
            {

                string threadInterval = System.Configuration.ConfigurationManager.AppSettings["threadIntervalSec"];
                
                if (!string.IsNullOrEmpty(threadInterval))
                {
                    if (!int.TryParse(threadInterval, out intervalSec))
                        intervalSec = 60;
                }               
                

               
                System.Threading.ThreadStart st = new System.Threading.ThreadStart(WorkerFunction);
                serviceStarted = true;
                workerThread = new System.Threading.Thread(st);
                workerThread.Start();
            }
            catch (Exception ex)
            {
                

                this.Stop();
            }
        }
        void WorkerFunction()
        {
            try
            {
                while (serviceStarted)
                {
                    if (!MailMerger.Merger.Running)
                    {
                        MailMerger.Merger.Run();
                    }

                    if (serviceStarted)
                    {
                        System.Threading.Thread.Sleep(intervalSec * 1000);
                    }
                }

                // time to end the thread
                System.Threading.Thread.CurrentThread.Abort();
            }
            catch (Exception)
            {
            }
        }

        protected override void OnStop()
        {
            try
            {
                serviceStarted = false;
                workerThread.Join();
                MailMerger.Merger.CleanUp();
            }
            catch (Exception)
            {
            }
        }
    }
}
