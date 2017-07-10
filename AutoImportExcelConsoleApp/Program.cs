using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using PS.BusinessLogic;
using PS.DataModel;
using System.IO;
using Microsoft.SharePoint;
using PS.DataService;
using System.Threading;
using PS.ImportExcelService;

namespace PS.ConsoleApp
{ 
    class Program
    {
        static void Main(string[] args)
        {
            Utility.WriteLog("############################################################", "Normal");
            Utility.WriteLog("ImportExcel Console started " + System.DateTime.Now.ToString(), "Normal");
            MainProcess Main = new MainProcess();
            Main.RunningProcessModelAndPart();
            Main.RunningProcessServiceManual();
            Main.RunningProcessModelBL();
            Utility.WriteLog("ImportExcel Console ended " + System.DateTime.Now.ToString(), "Normal");
            Utility.WriteLog("############################################################", "Normal");

            Environment.Exit(0);


        }

       
    }
}
