using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using PS.BusinessLogic;

namespace PS.ImportExcelService
{
    public partial class ExcelService : ServiceBase
    {
        private System.Timers.Timer timerExcute = new System.Timers.Timer();
        public ExcelService()
        {
            InitializeComponent();
        }


        protected override void OnStart(string[] args)
        {
            try
            {
                Utility.WriteLog("Start Window service ", "Normal");
                var interval = ConfigurationManager.AppSettings["Interval"];
                Utility.WriteLog("Interval =>" + interval.ToString(), "Normal");
                this.timerExcute.Interval = Convert.ToInt32(interval);
                this.timerExcute.Elapsed += new System.Timers.ElapsedEventHandler(this.ExcuteWork);
                this.timerExcute.Start();
            }
            catch (Exception ex)
            {
                Utility.WriteLog("Service Error : " + ex.Message, "Error");
            }
            
        }
        private void ExcuteWork(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                Utility.WriteLog("Start Timer tick ", "Normal");
                MainProcess main = new MainProcess();
                main.RunningProcessModelAndPart();
                main.RunningProcessServiceManual();
            }
            catch (Exception ex)
            {
                Utility.WriteLog("Error execute time tick : " + ex.Message, "Error");
            }
            
        }

        protected override void OnStop()
        {
            Utility.WriteLog("Service stopped " +System.DateTime.Now.ToString(), "Normal");
            this.timerExcute.Stop();
            this.timerExcute = null;
        }
    }
}
