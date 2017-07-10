using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PS.DataModel;
using PS.DataService;
using System.IO;
using System.Diagnostics;

namespace PS.BusinessLogic
{
    public class MainProcess
    {
        public string listLogName = "UPLOAD_LOG";

        public void RunningProcessModelAndPart()
        {
            var siteUrl = System.Configuration.ConfigurationManager.AppSettings.Get("SPList");
             
            try
            {
                Utility.WriteLog("Reading Model " + System.DateTime.Now.ToString(), "Normal"); 
                var listConfig = Utility.GetConfiguration();
                var sourcePath = listConfig.Where(x => x.Name == Constants.SourcePathModel).FirstOrDefault();
                var backupPath = listConfig.Where(x => x.Name == Constants.BackupPath).FirstOrDefault();
                var errorPath = listConfig.Where(x => x.Name == Constants.Error).FirstOrDefault();
                
                if (sourcePath != null)
                { 
                    if (Directory.GetFiles(sourcePath.Values, "*", SearchOption.AllDirectories).Length > 0)
                    {
                        foreach (var item in Directory.GetFiles(sourcePath.Values))
                        {
                            string fileName = Path.GetFileName(item);
                            int counterInsertSuccess = 0;
                            int counterUpdateSuccess = 0;
                            DataAccessModel SPListData = SPDataAccess.GetListModelData("SERVICE_PART_LIST", siteUrl);
                            string allModel = "";
                            if (item.ToUpper().Contains(".XLS") || item.ToUpper().Contains(".XLSX"))
                            {
                                var list_model = FileService.GetListModelData(item);

                                var groupedModel = list_model.OrderBy(x => x.Model_Name).GroupBy(x => x.Model_Name).ToList();

                                foreach (var name in groupedModel)
                                {
                                    allModel += " " + name.Key;
                                }

                                var listModelPart = (from p in list_model
                                                     where p.Part_No != ""
                                                     select p).ToList();

                                foreach (var model in groupedModel)
                                {
                                    if (model.Key != "")
                                    {
                                        var listFilter = (from p in list_model
                                                          where p.Model_Name == model.Key && p.Part_No != ""
                                                          select p).ToList();
                                        if (listFilter.Count > 0)
                                        {
                                            SPDataAccess.DeleteModel(model.Key, listFilter, "SERVICE_PART_LIST", siteUrl);
                                        } 
                                    } 
                                }
                                foreach (var itemModel in list_model)
                                {
                                    if (itemModel.Model_Name != "" && itemModel.Part_No != "")
                                    {

                                        var listOldModel = (from p in SPListData.listModelData
                                                            where p.Model_Name == itemModel.Model_Name && p.Part_No == itemModel.Part_No
                                                            select p).ToList();
 
                                        //####### Rewrite Model #######
                                        try
                                        { 
                                            SPDataAccess.UpdateAllPart(itemModel, "SERVICE_PART_LIST", siteUrl);
                                            if (listOldModel.Count > 0)
                                            {
                                                foreach (var oldModel in listOldModel)
                                                {
                                                    SPDataAccess.UpdateListModel(oldModel, itemModel, "SERVICE_PART_LIST", siteUrl, fileName);

                                                    counterUpdateSuccess++;
                                                } 
                                            }
                                            else
                                            {
                                                SPDataAccess.AddListModel(itemModel, "SERVICE_PART_LIST", siteUrl, fileName);

                                                counterInsertSuccess++;
                                            } 
                                        }
                                        catch (Exception ex)
                                        {
                                            Utility.WriteLog(itemModel.Model_Name + " => " + ex.Message, "Add Error");
                                        } 
                                    }
                                    else if (itemModel.Model_Name == "" && itemModel.Part_No != "")
                                    {
                                        try
                                        {
                                            SPDataAccess.UpdateAllPart(itemModel, "SERVICE_PART_LIST", siteUrl);

                                            counterUpdateSuccess++;
                                        }
                                        catch (Exception ex)
                                        {
                                            Utility.WriteLog(itemModel.Part_No + " => " + ex.Message, "Update Error");
                                        }
                                    }
                                    else if (itemModel.Model_Name != "" && itemModel.Part_No == "")
                                    {
                                        try
                                        {
                                            SPDataAccess.UpdateAllModel(itemModel, "SERVICE_PART_LIST", siteUrl);

                                            counterUpdateSuccess++;
                                        }
                                        catch (Exception ex)
                                        {
                                            Utility.WriteLog(itemModel.Part_No + " => " + ex.Message, "Update Error");
                                        }
                                    }
                                }
                                File.Move(item, backupPath.Values + @"\" + Path.GetFileNameWithoutExtension(item) + "_" + System.DateTime.Now.ToString("MM-dd-yy_HHmmss") + Path.GetExtension(item));
                                Utility.WriteLog("File Name:" + Path.GetFileName(item) + " (Model: " + allModel + " Actual Data: " + list_model.Count + " records. Inserted: " + counterInsertSuccess + " records. Update: " + counterUpdateSuccess + " records.)", "Success");
                                SPDataAccess.AddLog(Path.GetFileName(item), " (Model: " + allModel + " Actual Data: " + list_model.Count + " records. Inserted: " + counterInsertSuccess + " records. Update: " + counterUpdateSuccess + " records.) Success", listLogName, siteUrl, "Service Part List");
                            }
                            else
                            {

                                File.Move(item, errorPath.Values + @"\" + Path.GetFileNameWithoutExtension(item) + "_" + System.DateTime.Now.ToString("MM-dd-yy_HHmmss") + Path.GetExtension(item));
                                Utility.WriteLog("File: " + item + " is not Excel format", "Error");
                                SPDataAccess.AddLog(Path.GetFileName(item), "File: " + item + " is not Excel format  Error", listLogName, siteUrl, "Service Part List");
                            } 
                        }
                    } 
                }  
            }
            catch (Exception ex)
            {
                var st = new StackTrace(ex, true);
                var frame = st.GetFrame(0);
                var line = frame.GetFileLineNumber();
                Utility.WriteLog("Running Model : " + ex.ToString() + "Line: " + line.ToString(), "Error");
                SPDataAccess.AddLog("Exception", ex.Message, "UPLOAD_LOG", siteUrl, "Service Part List");
            }
            finally
            {
                //  Utility.KillExcel();
            }
        }

        public void RunningProcessServiceManual()
        {
            var siteUrl = System.Configuration.ConfigurationManager.AppSettings.Get("SPList");
            try
            {
                Utility.WriteLog("Reading Service Manual " + System.DateTime.Now.ToString(), "Normal");
                int counterInsertSuccess = 0;
                int counterUpdateSuccess = 0;
                var listConfig = Utility.GetConfiguration();
                var sourcePath = listConfig.Where(x => x.Name == Constants.SourceServiceManual).FirstOrDefault();
                var backupPath = listConfig.Where(x => x.Name == Constants.BackupPathSVM).FirstOrDefault();
                var errorPath = listConfig.Where(x => x.Name == Constants.Error).FirstOrDefault();

                if (sourcePath != null)
                {

                    if (Directory.GetFiles(sourcePath.Values, "*", SearchOption.AllDirectories).Length > 0)
                    {
                        DataAccessModel oldListData = SPDataAccess.GetListServiceManualData("SERVICE_MANUAL_LIST", siteUrl);
                        int maxNumber = 0;
                        if (oldListData.listSVMData != null)
                        {
                            maxNumber = oldListData.listSVMData.Count;
                        }

                        foreach (var item in Directory.GetFiles(sourcePath.Values))
                        {
                            if (item.ToUpper().Contains(".XLS") || item.ToUpper().Contains(".XLSX"))
                            {
                                var list_svm = FileService.GetListSVMData(item);
                                string fileName = Path.GetFileName(item);
                                if (list_svm.Count > 0)
                                {
                                    int index = 0;
                                    int tempIndex = 0;
                                    foreach (var itemSVM in list_svm)
                                    {
                                        int DuplicateIndex = oldListData.listSVMData.FindIndex(a => a.Indoor_Model_Name == itemSVM.Indoor_Model_Name && a.Outdoor_Model_Name == itemSVM.Outdoor_Model_Name);

                                        if (DuplicateIndex == -1) // case Insert
                                        {
                                            try
                                            {
                                                index = index + 1;
                                                tempIndex = index + maxNumber;
                                                itemSVM.Title = tempIndex.ToString("00000");
                                                SPDataAccess.AddListServiceManual(itemSVM, "SERVICE_MANUAL_LIST", siteUrl, fileName);
                                                counterInsertSuccess++;
                                            }
                                            catch (Exception ex)
                                            {
                                                Utility.WriteLog(itemSVM.Indoor_Model_Name + " And " + itemSVM.Outdoor_Model_Name + " => " + ex.Message, "Error");
                                            }

                                        }
                                        else if (DuplicateIndex >= 0) // case Update
                                        {
                                            try
                                            {
                                                SPDataAccess.UpdateListServiceManual(oldListData.listSVMData[DuplicateIndex], itemSVM, "SERVICE_MANUAL_LIST", siteUrl, fileName);
                                                counterUpdateSuccess++;
                                            }
                                            catch (Exception ex)
                                            {
                                                Utility.WriteLog(itemSVM.Indoor_Model_Name + " And " + itemSVM.Outdoor_Model_Name + " => " + ex.Message, "Error");
                                            }
                                        }
                                    }
                                    File.Move(item, backupPath.Values + @"\" + Path.GetFileNameWithoutExtension(item) + "_" + System.DateTime.Now.ToString("MM-dd-yy_HHmmss") + Path.GetExtension(item));
                                    Utility.WriteLog("File Name:" + Path.GetFileName(item) + " (Actual Data: " + list_svm.Count + " records. Inserted: " + counterInsertSuccess + " records. Updated: " + counterUpdateSuccess + " records)", "Success");
                                    SPDataAccess.AddLog(Path.GetFileName(item), "(Actual Data: " + list_svm.Count + " records. Inserted: " + counterInsertSuccess + " records. Updated: " + counterUpdateSuccess + " records) Success", listLogName, siteUrl, "Service Manual");
                                }
                                else
                                {
                                    File.Move(item, errorPath.Values + @"\" + Path.GetFileNameWithoutExtension(item) + "_" + System.DateTime.Now.ToString("MM-dd-yy_HHmmss") + Path.GetExtension(item));
                                    Utility.WriteLog("File: " + item + " is not correct format", "Error");
                                    SPDataAccess.AddLog(Path.GetFileName(item), "File: " + item + " is not correct format Error", listLogName, siteUrl, "Service Manual");
                                }

                            }
                            else
                            {

                                File.Move(item, errorPath.Values + @"\" + Path.GetFileNameWithoutExtension(item) + "_" + System.DateTime.Now.ToString("MM-dd-yy_HHmmss") + Path.GetExtension(item));
                                Utility.WriteLog("File: " + item + " is not Excel format", "Error");
                                SPDataAccess.AddLog(Path.GetFileName(item), "File: " + item + " is not correct format Error", "UPLOAD_LOG", siteUrl, "Service Manual");

                            } 
                        }
                    } 
                }


            }
            catch (Exception ex)
            {
                var st = new StackTrace(ex, true);
                var frame = st.GetFrame(0);
                var line = frame.GetFileLineNumber();
                Utility.WriteLog("Running Service Manual : " + ex.ToString() + "Line: " + line.ToString(), "Error");
                SPDataAccess.AddLog("Exception", ex.Message, listLogName, siteUrl, "Service Manual");
            }
            finally
            {
                // Utility.KillExcel();
            }
        }

        public void RunningProcessModelBL()
        {
            var siteUrl = System.Configuration.ConfigurationManager.AppSettings.Get("BL-SPList");
            try
            {
                Utility.WriteLog("Reading Model BusinessLounge " + System.DateTime.Now.ToString(), "Normal");
                int counterInsertSuccess = 0;
                int counterUpdateSuccess = 0;
                var listConfig = Utility.GetConfiguration();
                var sourcePath = listConfig.Where(x => x.Name == Constants.SourcePathModelBL).FirstOrDefault();
                var backupPath = listConfig.Where(x => x.Name == Constants.BackupPathBL).FirstOrDefault();
                var errorPath = listConfig.Where(x => x.Name == Constants.Error).FirstOrDefault();

                if (sourcePath != null)
                {

                    if (Directory.GetFiles(sourcePath.Values, "*", SearchOption.AllDirectories).Length > 0)
                    {
                        DataAccessModel oldListData = SPDataAccess.GetListModelBLData("PRODUCTS", siteUrl);

                        foreach (var item in Directory.GetFiles(sourcePath.Values))
                        {
                            if (item.ToUpper().Contains(".XLS") || item.ToUpper().Contains(".XLSX"))
                            {
                                var list_ModelBL = FileService.GetListModelBL(item);
                                string fileName = Path.GetFileName(item);
                                 
                                if (list_ModelBL.Count > 0)
                                {
                                    foreach (var itemModelBL in list_ModelBL)
                                    {
                                        int DuplicateIndex = oldListData.listModelBLData.FindIndex(a => a.INDOOR == itemModelBL.INDOOR && a.OUTDOOR == a.OUTDOOR);
                                        //int[] duplicateSet = oldListData.listModelBLData.Select((b, i) => b.INDOOR == itemModelBL.INDOOR && b.OUTDOOR == itemModelBL.OUTDOOR ? i : -1).Where(i => i != -1).ToArray();
                                        if (DuplicateIndex == -1) // case Insert
                                        {
                                            try
                                            {
                                                SPDataAccess.AddListModelBL(itemModelBL, "PRODUCTS", siteUrl, fileName);
                                                counterInsertSuccess++;
                                            }
                                            catch (Exception ex)
                                            {
                                                Utility.WriteLog(itemModelBL.Id + " => " + ex.Message, "Error");
                                            }
                                        }
                                        else if (DuplicateIndex >= 0) // case Update
                                        {
                                            try
                                            {
                                                SPDataAccess.UpdateListModelBL(oldListData.listModelBLData[DuplicateIndex], itemModelBL, "PRODUCTS", siteUrl, fileName);
                                                counterUpdateSuccess++;
                                            }
                                            catch (Exception ex)
                                            {
                                                Utility.WriteLog(itemModelBL.Id + " => " + ex.Message, "Error");
                                            }
                                        }
                                    }
                                    File.Move(item, backupPath.Values + @"\" + Path.GetFileNameWithoutExtension(item) + "_" + System.DateTime.Now.ToString("MM-dd-yy_HHmmss") + Path.GetExtension(item));
                                    Utility.WriteLog("File Name:" + Path.GetFileName(item) + " (Actual Data: " + list_ModelBL.Count + " records. Inserted: " + counterInsertSuccess + " records. Updated: " + counterUpdateSuccess + " records)", "Success");
                                    SPDataAccess.AddLog(Path.GetFileName(item), "(Actual Data: " + list_ModelBL.Count + " records. Inserted: " + counterInsertSuccess + " records. Updated: " + counterUpdateSuccess + " records) Success", listLogName, siteUrl, "Model BL");
                                }
                                else
                                {
                                    File.Move(item, errorPath.Values + @"\" + Path.GetFileNameWithoutExtension(item) + "_" + System.DateTime.Now.ToString("MM-dd-yy_HHmmss") + Path.GetExtension(item));
                                    Utility.WriteLog("File: " + item + " is not correct format", "Error");
                                    SPDataAccess.AddLog(Path.GetFileName(item), "File: " + item + " is not correct format Error", listLogName, siteUrl, "Models BL");
                                }
                            }
                            else
                            {
                                File.Move(item, errorPath.Values + @"\" + Path.GetFileNameWithoutExtension(item) + "_" + System.DateTime.Now.ToString("MM-dd-yy_HHmmss") + Path.GetExtension(item));
                                Utility.WriteLog("File: " + item + " is not Excel format", "Error");
                                SPDataAccess.AddLog(Path.GetFileName(item), "File: " + item + " is not Excel format", listLogName, siteUrl, "Models BL");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var st = new StackTrace(ex, true);
                var frame = st.GetFrame(0);
                var line = frame.GetFileLineNumber();
                Utility.WriteLog("Running Model BL : " + ex.ToString() + "Line: " + line.ToString(), "Error");
                SPDataAccess.AddLog("Exception", ex.Message, listLogName, siteUrl, "Models BL");
            }
            finally
            {
                // Utility.KillExcel();
            }
        }

    }
}
