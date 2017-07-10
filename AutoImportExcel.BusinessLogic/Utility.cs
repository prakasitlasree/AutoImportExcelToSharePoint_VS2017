using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using PS.DataModel;
using System.IO;
using System.Diagnostics;
using System.Configuration;

namespace PS.BusinessLogic
{
    public static class Utility
    {
        public static void KillExcel()
        {
            Process[] AllProcesses = Process.GetProcessesByName("excel");


            foreach (Process ExcelProcess in AllProcesses)
            {
                ExcelProcess.Kill();
            }

            AllProcesses = null;
        }
        public static List<Config> GetConfiguration()
        {
            try
            {

                var sourcePathModel = System.Configuration.ConfigurationManager.AppSettings["SourcePathModel"];
                var sourcePathServiceManual = System.Configuration.ConfigurationManager.AppSettings["SourcePathServiceManual"];
                var sourcePathModelBL = System.Configuration.ConfigurationManager.AppSettings["SourcePathModelBL"];
                var backupPath = System.Configuration.ConfigurationManager.AppSettings["BackupPath"];
                var BackupPathSVM = System.Configuration.ConfigurationManager.AppSettings["BackupPathSVM"];
                var BackupPathBL = System.Configuration.ConfigurationManager.AppSettings["BackupPathBL"];
                var errorPath = System.Configuration.ConfigurationManager.AppSettings["ErrorPath"];
                var Log = System.Configuration.ConfigurationManager.AppSettings["LogPath"];

                var list = new List<Config>();
                if (sourcePathModel!= null)
                {
                    var obj = new Config();
                    obj.Name = Constants.SourcePathModel;
                    obj.Values = sourcePathModel;
                    list.Add(obj);
                }
                if (backupPath != null)
                {
                    var obj = new Config();
                    obj.Name = Constants.BackupPath;
                    obj.Values = backupPath;
                    list.Add(obj);
                }
                if (BackupPathSVM != null)
                {
                    var obj = new Config();
                    obj.Name = Constants.BackupPathSVM;
                    obj.Values = BackupPathSVM;
                    list.Add(obj);
                }
                if (BackupPathBL != null)
                {
                    var obj = new Config();
                    obj.Name = Constants.BackupPathBL;
                    obj.Values = BackupPathBL;
                    list.Add(obj);
                }
                if (Log != null)
                {
                    var obj = new Config();
                    obj.Name = Constants.Log;
                    obj.Values = Log;
                    list.Add(obj);
                }
                if (sourcePathServiceManual != null)
                {
                    var obj = new Config();
                    obj.Name = Constants.SourceServiceManual;
                    obj.Values = sourcePathServiceManual;
                    list.Add(obj);
                }
                if (errorPath != null)
                {
                    var obj = new Config();
                    obj.Name = Constants.Error;
                    obj.Values = errorPath;
                    list.Add(obj);
                }
                if (sourcePathModelBL != null)
                {
                    var obj = new Config();
                    obj.Name = Constants.SourcePathModelBL;
                    obj.Values = sourcePathModelBL;
                    list.Add(obj);
                }
                return list;
            }
            catch (Exception ex)
            {
                return new List<Config>();
            }
        }

        public static void WriteLog(string Message,string type)
        {
            try
            {
                var listConfig = GetConfiguration();
                foreach (var item in listConfig)
                {
                    if (!System.IO.Directory.Exists(item.Values))
                    {
                        System.IO.Directory.CreateDirectory(item.Values);
                    }

                 
                }
                StringBuilder templateLog = new StringBuilder();
                
               
                switch (type)
                {
                   
                    case "Normal":
                        templateLog.AppendLine(Message);

                        
                        break;
                    case "Success":
                        templateLog.AppendLine("Success");
                        templateLog.AppendLine(System.DateTime.Now.ToString());
                        templateLog.AppendLine(Message);
                        
                        break;
                    case "Error":
                        templateLog.AppendLine("Error");
                        templateLog.AppendLine(System.DateTime.Now.ToString());
                        templateLog.AppendLine(Message);
                       
                        break;
                }
               


                string logPath = System.Configuration.ConfigurationManager.AppSettings.Get("LogPath");
                string path = logPath+@"\PS-LogUpload_"+System.DateTime.Now.ToString("MM-dd-yy")+".txt";
                try
                {

                   File.AppendAllText(path,templateLog.ToString() + Environment.NewLine); 
                   
                }
                catch(Exception ex)
                {

                }
    
            }
            catch (Exception ex)
            {
              
            }
        }
    }
}
