using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using PS.DataModel;
using System.Data.OleDb;
using System.Data;
using Microsoft.SharePoint;
using System.IO;
using System.Diagnostics;

namespace PS.BusinessLogic
{
    public static class FileService
    {
        //private static DataTable GetDataTable(string fullPath)
        //{

        //    Microsoft.Office.Interop.Excel.Application m_excelApp = new Excel.Application();
        //    m_excelApp.Workbooks.Open(fullPath, false, true);
        //    Microsoft.Office.Interop.Excel.Worksheet xlsSheet = ((Microsoft.Office.Interop.Excel.Worksheet)(m_excelApp.Sheets[1]));
        //    DataTable dtResult = null;
        //    dtResult = WorksheetToDataTable(xlsSheet);
        //    Marshal.ReleaseComObject(xlsSheet);
        //    Marshal.ReleaseComObject(m_excelApp);


        //    return dtResult;

        //}
        //private static System.Data.DataTable WorksheetToDataTable(Microsoft.Office.Interop.Excel.Worksheet worksheet)
        //{
        //    int rows = worksheet.UsedRange.Rows.Count;
        //    int cols = worksheet.UsedRange.Columns.Count;
        //    System.Data.DataTable dt = new System.Data.DataTable();
        //    int noofrow = 1;
        //    for (int c = 1; (c <= cols); c++)
        //    {
        //        string colname = ("F" + c.ToString());
        //        dt.Columns.Add(colname);
        //        noofrow = 1;
        //    }

        //    for (int r = noofrow; (r <= rows); r++)
        //    {
        //        DataRow dr = dt.NewRow();
        //        for (int c = 1; (c <= cols); c++)
        //        {
        //            dr[(c - 1)] = worksheet.Cells[r, c].Text;
        //        }

        //        dt.Rows.Add(dr);
        //    }
        //    return dt;
        //}

        public static List<Model> GetListModelData(string fullPath)
        {

            List<Model> listData = new List<Model>();
            try
            {
                DataTable dtResult = new DataTable();

                int totalSheet = 0;
                using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fullPath + ";Extended Properties='Excel 12.0;HDR=NO;IMEX=1;';"))
                {
                    objConn.Open();
                    OleDbCommand cmd = new OleDbCommand();
                    OleDbDataAdapter oleda = new OleDbDataAdapter();
                    DataSet ds = new DataSet();
                    DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string sheetName = string.Empty;
                    if (dt != null)
                    {
                        var tempDataTable = (from dataRow in dt.AsEnumerable()
                                             where !dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")
                                             select dataRow).CopyToDataTable();
                        dt = tempDataTable;
                        totalSheet = dt.Rows.Count;
                        sheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                    }
                    cmd.Connection = objConn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                    oleda = new OleDbDataAdapter(cmd);
                    oleda.Fill(ds, "excelData");
                    dtResult = ds.Tables["excelData"];
                    objConn.Close();
                }


                if (dtResult.Rows.Count > 0 && dtResult.Columns.Count == 29)
                {
                    var listModelData = (from table in dtResult.AsEnumerable()
                                         select new Model
                                         {
                                             Model_Name = table["F1"].ToString(),
                                             Model_Description = table["F2"].ToString(),
                                             Category_Code = table["F3"].ToString(),
                                             First_Production_Date = table["F4"].ToString(),
                                             Exploded_Diagram_NO = table["F5"].ToString(),
                                             Part_No = table["F7"].ToString(),
                                             Location_No = table["F6"].ToString(),
                                             RoHS = table["F8"].ToString(),
                                             Description = table["F9"].ToString(),
                                             Drawing_NO = table["F10"].ToString(),
                                             Qty = table["F11"].ToString(),
                                             Price_USD = table["F12"].ToString(),
                                             Price_THB = table["F13"].ToString(),
                                             Price_EUR = table["F14"].ToString(),
                                             Part_Group = table["F15"].ToString(),
                                             Net_Weight = table["F16"].ToString(),
                                             Type_No = table["F17"].ToString(),
                                             Country_Of_Origin = table["F18"].ToString(),
                                             STC_Mark = table["F19"].ToString(),
                                             EMC_Code = table["F20"].ToString(),
                                             ECCN_Code = table["F21"].ToString(),
                                             Page_No = table["F22"].ToString(),
                                             Zone_Code = table["F23"].ToString(),
                                             Last_Production_date = table["F24"].ToString(),
                                             Retention_Period = table["F25"].ToString(),
                                             Recomment = table["F26"].ToString(),
                                             Substituted = table["F27"].ToString(),
                                             Lead_Time = table["F28"].ToString(),
                                             Final_Buy_Date = table["F29"].ToString()

                                         }).ToList();

                    listModelData.RemoveAt(0);
                    listData = listModelData;
                    Utility.WriteLog("Read : " + fullPath, "Success");
                    return listData;
                }
                else
                {
                    return listData;
                }


            }
            catch (Exception ex)
            {
                Utility.WriteLog(ex.Message.ToString(), "Error");
                return listData;
            }

        }

        public static List<ServiceManual> GetListSVMData(string fullPath)
        {

            List<ServiceManual> listData = new List<ServiceManual>();
            try
            {
                DataTable dtResult = new DataTable();

                int totalSheet = 0;
                using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fullPath + ";Extended Properties='Excel 12.0;HDR=NO;IMEX=1;';"))
                {
                    objConn.Open();
                    OleDbCommand cmd = new OleDbCommand();
                    OleDbDataAdapter oleda = new OleDbDataAdapter();
                    DataSet ds = new DataSet();
                    DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string sheetName = string.Empty;
                    if (dt != null)
                    {
                        var tempDataTable = (from dataRow in dt.AsEnumerable()
                                             where !dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")
                                             select dataRow).CopyToDataTable();
                        dt = tempDataTable;
                        totalSheet = dt.Rows.Count;
                        sheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                    }
                    cmd.Connection = objConn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                    oleda = new OleDbDataAdapter(cmd);
                    oleda.Fill(ds, "excelData");
                    dtResult = ds.Tables["excelData"];
                    objConn.Close();
                }

                if (dtResult.Rows.Count > 0 && dtResult.Columns.Count == 6)
                {
                    var listSVMData = (from table in dtResult.AsEnumerable()
                                       select new ServiceManual
                                       {
                                           Indoor_Model_Name = table["F1"].ToString(),
                                           Outdoor_Model_Name = table["F2"].ToString(),
                                           Issue_Date = table["F5"].ToString(),
                                           MDC_Code = table["F4"].ToString(),
                                           SVM_Remark = table["F6"].ToString(),
                                           SVM_FileName = table["F3"].ToString()

                                       }).ToList();

                    listSVMData.RemoveAt(0);
                    listData = listSVMData;
                    Utility.WriteLog("Read : " + fullPath, "Success");
                    return listData;
                }
                else
                {
                    return listData;
                }


            }
            catch (Exception ex)
            {
                Utility.WriteLog(ex.Message.ToString(), "Error");
                return listData;
            }

        }

        public static List<ModelBL> GetListModelBL(string fullPath)
        {

            List<ModelBL> listData = new List<ModelBL>();
            try
            {
                DataTable dtResult = new DataTable();

                int totalSheet = 0;
                using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fullPath + ";Extended Properties='Excel 12.0;HDR=NO;IMEX=1;';"))
                {
                    objConn.Open();
                    OleDbCommand cmd = new OleDbCommand();
                    OleDbDataAdapter oleda = new OleDbDataAdapter();
                    DataSet ds = new DataSet();
                    DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string sheetName = string.Empty;
                    if (dt != null)
                    {
                        var tempDataTable = (from dataRow in dt.AsEnumerable()
                                             where !dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")
                                             select dataRow).CopyToDataTable();
                        dt = tempDataTable;
                        totalSheet = dt.Rows.Count;
                        sheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                    }
                    cmd.Connection = objConn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                    oleda = new OleDbDataAdapter(cmd);
                    oleda.Fill(ds, "excelData");
                    dtResult = ds.Tables["excelData"];
                    objConn.Close();
                }
                dtResult = dtResult.Rows.Cast<DataRow>().Where(row => !row.ItemArray.All(field => field is DBNull || string.IsNullOrWhiteSpace(field as string))).CopyToDataTable();

                if (dtResult.Rows.Count > 0 && dtResult.Rows[0]["F9"].ToString() == "Manual" && dtResult.Rows[2]["F1"].ToString() == "BRAND")
                {
                    var listModelBLData = (from table in dtResult.AsEnumerable()
                                           select new ModelBL
                                           {
                                               Title = GetTitle(table["F7"].ToString(), table["F8"].ToString()),
                                               BRAND = table["F1"].ToString(),
                                               BL_CATEGORY = table["F2"].ToString(),
                                               BL_PRODUCT_TYPE = table["F3"].ToString(),
                                               BL_PRODUCT_SIZE = table["F5"].ToString(),
                                               REFRIGERANT = table["F6"].ToString(),
                                               INDOOR = table["F7"].ToString(),
                                               OUTDOOR = table["F8"].ToString(),
                                               INSTALLATION = table["F9"].ToString(),
                                               OWNER = table["F10"].ToString(),
                                               DISC = table["F11"].ToString(),
                                               SPECIFICATION = table["F12"].ToString(),
                                               BULLETIN = table["F13"].ToString(),
                                               DATABOOK = table["F14"].ToString(),
                                               VDO = table["F15"].ToString(),
                                               PRESENTATION = table["F16"].ToString(),
                                               IMAGE_LOW = table["F17"].ToString(),
                                               IMAGE_HD = table["F18"].ToString(),
                                               CATALOGUE = table["F19"].ToString()

                                           }).ToList();

                    listModelBLData.RemoveAt(0);
                    listModelBLData.RemoveAt(0);
                    listModelBLData.RemoveAt(0);


                    listData = listModelBLData;
                    Utility.WriteLog("Read : " + fullPath, "Success");
                    return listData;
                }
                else
                {
                    return listData;
                }


            }
            catch (Exception ex)
            {
                Utility.WriteLog(ex.Message.ToString(), "Error");
                return listData;
            }

        }

        private static string GetTitle(string indoor, string outdoor)
        {
            var result = "";
            try
            {
                if (indoor =="" || indoor == "*" || indoor =="-")
                {
                    result = outdoor;
                }
                else
                {
                    result = indoor;
                }
            }
            catch (Exception)
            {
                result = "";
            }
            return result;
        }

        public static void UploadServiceManual(string fullPath, string docLibName, string url)
        {


            using (SPSite oSite = new SPSite(url))
            {
                using (SPWeb oWeb = oSite.OpenWeb())
                {

                    SPFolder myLibrary = oWeb.Folders[docLibName];

                    // Prepare to upload
                    String fileName = System.IO.Path.GetFileName(fullPath);
                    FileStream fileStream = File.OpenRead(fullPath);

                    // Upload document
                    SPFile spfile = myLibrary.Files.Add(fileName, fileStream, false);

                    // Commit 
                    myLibrary.Update();

                    fileStream.Close();
                    fileStream.Dispose();

                }
            }
        }
    }
}
