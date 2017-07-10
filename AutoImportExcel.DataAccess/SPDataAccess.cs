using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PS.DataModel;
using Microsoft.SharePoint;
using System.Data;
using System.Globalization;

namespace PS.DataService
{
    public static class SPDataAccess
    {
        #region ====== Service Part List ======

        public static DataAccessModel GetListModelData(string listName, string url)
        {
            var outPutData = new DataAccessModel();
            try
            {
                using (SPSite oSite = new SPSite(url))
                {
                    using (SPWeb oWeb = oSite.OpenWeb())
                    {
                        List<string> camlQuery = new List<string>();
                        camlQuery = GetPartCAML();

                        SPQuery query = new SPQuery();
                        query.Query = camlQuery[0].ToString();
                        query.ViewFields = camlQuery[1].ToString();
                        SPList oList = oWeb.Lists[listName];
                        SPListItemCollection items = oList.GetItems(query);
                        var dt = items.GetDataTable();
                        var list = (from table in dt.AsEnumerable()
                                    select new Model
                                    {
                                        Id = table["ID"].ToString(),
                                        Model_Name = table["Model_Name"].ToString(),
                                        Model_Description = table["Model_Description"].ToString(),
                                        Category_Code = table["Category_Code"].ToString(),
                                        First_Production_Date = table["First_Production_date"].ToString(),
                                        Exploded_Diagram_NO = table["Exploded_Diagram_ref_no"].ToString(),
                                        Part_No = table["Part_no"].ToString(),
                                        Location_No = table["Location_no"].ToString(),
                                        RoHS = table["RoHS"].ToString(),
                                        Description = table["More_Description"].ToString(),
                                        Drawing_NO = table["Drawing_no"].ToString(),
                                        Qty = table["Qty"].ToString(),
                                        Price_USD = table["Price_USD"].ToString(),
                                        Price_THB = table["Price_THB"].ToString(),
                                        Price_EUR = table["Price_EUR"].ToString(),
                                        Part_Group = table["Part_Group"].ToString(),
                                        Net_Weight = table["Net_weight"].ToString(),
                                        Type_No = table["Type_No"].ToString(),
                                        Country_Of_Origin = table["Country_of_origin"].ToString(),
                                        STC_Mark = table["STC_mark"].ToString(),
                                        EMC_Code = table["EMC_code"].ToString(),
                                        ECCN_Code = table["ECCN_code"].ToString(),
                                        Import_Filename = table["IMPORT_FILE"].ToString()

                                    }).ToList();

                        outPutData.listModelData = list;
                        outPutData.Status = true;
                        outPutData.Reason = "Success";

                    }
                }
                return outPutData;
            }
            catch (Exception ex)
            {
                outPutData.listModelData = new List<Model>();
                outPutData.Status = false;
                outPutData.Reason = ex.Message.ToString();
                return outPutData;
            }
        }

        public static void AddListModel(Model itemModel, string listName, string url, string fileName)
        {
            var outPutData = new DataAccessModel();

            using (SPSite oSiteCollection = new SPSite(url))
            {
                using (SPWeb oWeb = oSiteCollection.OpenWeb())
                {
                    SPList oList = oWeb.Lists[listName];

                    SPListItem oListItem = oList.Items.Add();
                    oListItem["Title"] = itemModel.Model_Name;
                    oListItem["Model_Name"] = itemModel.Model_Name;
                    oListItem["Model_Description"] = itemModel.Model_Description;
                    oListItem["Category_Code"] = itemModel.Category_Code;
                    oListItem["First_Production_date"] = itemModel.First_Production_Date;
                    oListItem["Exploded_Diagram_ref_no"] = itemModel.Exploded_Diagram_NO; ;
                    oListItem["Part_no"] = itemModel.Part_No;
                    oListItem["Location_no"] = itemModel.Location_No;
                    oListItem["RoHS"] = itemModel.RoHS;
                    oListItem["More_Description"] = itemModel.Description;
                    oListItem["Drawing_no"] = itemModel.Drawing_NO;
                    oListItem["Qty"] = itemModel.Qty;
                    oListItem["Price_USD"] = itemModel.Price_USD;
                    oListItem["Price_THB"] = itemModel.Price_THB;
                    oListItem["Price_EUR"] = itemModel.Price_EUR;
                    oListItem["Part_Group"] = itemModel.Part_Group;
                    oListItem["Net_weight"] = itemModel.Net_Weight;
                    oListItem["Type_No"] = itemModel.Type_No;
                    oListItem["Country_of_origin"] = itemModel.Country_Of_Origin;
                    oListItem["STC_mark"] = itemModel.STC_Mark;
                    oListItem["EMC_code"] = itemModel.EMC_Code;
                    oListItem["ECCN_code"] = itemModel.ECCN_Code;
                    oListItem["PAGE_NO"] = itemModel.Page_No;
                    oListItem["Substituted"] = itemModel.Substituted;
                    oListItem["Sub_Part_Price"] = itemModel.Sub_Part_Price;
                    DateTime outDateFinal = new DateTime();
                    DateTime outDateFirst = new DateTime();
                    try
                    {
                        if (itemModel.Final_Buy_Date != "")
                        {
                            var dtFinal = DateTime.TryParse(itemModel.Final_Buy_Date, out outDateFinal);
                            if (!dtFinal)
                            {
                                outDateFinal = DateTime.ParseExact(itemModel.Final_Buy_Date,
                                "yyyyMMdd",
                                 CultureInfo.InvariantCulture);
                                oListItem["Retention_Period"] = outDateFinal.Year + 8;
                            }
                            else
                            {
                                oListItem["Retention_Period"] = outDateFinal.Year + 8;
                            }
                        }
                        else if (itemModel.First_Production_Date != "")
                        {
                            var dtFirst = DateTime.TryParse(itemModel.First_Production_Date, out outDateFirst);
                            if (!dtFirst)
                            {
                                outDateFirst = DateTime.ParseExact(itemModel.First_Production_Date,
                                "yyyyMMdd",
                                 CultureInfo.InvariantCulture);
                                oListItem["Retention_Period"] = outDateFirst.Year + 8;
                            }
                            else
                            {
                                oListItem["Retention_Period"] = outDateFirst.Year + 8;
                            } 
                        }
                    }
                    catch
                    {

                        oListItem["Retention_Period"] = itemModel.Retention_Period == "" ? oListItem["Retention_Period"] : itemModel.Retention_Period;
                    }
                    oListItem["Required_Per_Unit"] = itemModel.Required_Per_Unit;
                    oListItem["Final_Buy_Date"] = itemModel.Final_Buy_Date;
                    oListItem["Recomment"] = itemModel.Recomment;
                    oListItem["Lead_Time"] = itemModel.Lead_Time;
                    oListItem["Last_Production_date"] = itemModel.Last_Production_date;
                    oListItem["Zone_Code"] = itemModel.Zone_Code;
                    oListItem["IMPORT_FILE"] = fileName;
                    oListItem.Update();

                }
            }

        }

        public static void AddLog(string importFile, string Message, string listName, string url, string type)
        {
            using (SPSite oSiteCollection = new SPSite(url))
            {
                using (SPWeb oWeb = oSiteCollection.OpenWeb())
                {
                    SPList oList = oWeb.Lists[listName];

                    SPListItem oListItem = oList.Items.Add();
                    oListItem["Title"] = type;
                    oListItem["IMPORT_FILE"] = importFile;
                    oListItem["MESSAGE"] = Message;

                    oListItem.Update();
                }
            }
        }

        public static void UpdateListModel(Model oldItemModel, Model itemModel, string listName, string url, string fileName)
        {
            var outPutData = new DataAccessModel();
            using (SPSite oSiteCollection = new SPSite(url))
            {
                using (SPWeb oWeb = oSiteCollection.OpenWeb())
                {
                    SPList oList = oWeb.Lists[listName];
                    SPListItem oListItem = oList.GetItemById(Convert.ToInt32(oldItemModel.Id));
                    oListItem["Title"] = itemModel.Model_Name == "" ? oListItem["Title"] : itemModel.Model_Name;
                    oListItem["Model_Name"] = itemModel.Model_Name == "" ? oListItem["Model_Name"] : itemModel.Model_Name;
                    oListItem["Model_Description"] = itemModel.Model_Description == "" ? oListItem["Model_Description"] : itemModel.Model_Description;
                    oListItem["Category_Code"] = itemModel.Category_Code == "" ? oListItem["Category_Code"] : itemModel.Category_Code;
                    oListItem["First_Production_date"] = itemModel.First_Production_Date == "" ? oListItem["First_Production_date"] : itemModel.First_Production_Date;
                    oListItem["Exploded_Diagram_ref_no"] = itemModel.Exploded_Diagram_NO == "" ? oListItem["Exploded_Diagram_ref_no"] : itemModel.Exploded_Diagram_NO;
                    oListItem["Part_no"] = itemModel.Part_No == "" ? oListItem["Part_no"] : itemModel.Part_No;
                    oListItem["Location_no"] = itemModel.Location_No == "" ? oListItem["Location_no"] : itemModel.Location_No;
                    oListItem["RoHS"] = itemModel.RoHS == "" ? oListItem["RoHS"] : itemModel.RoHS;
                    oListItem["More_Description"] = itemModel.Description == "" ? oListItem["More_Description"] : itemModel.Description;
                    oListItem["Drawing_no"] = itemModel.Drawing_NO == "" ? oListItem["Drawing_no"] : itemModel.Drawing_NO;
                    oListItem["Qty"] = itemModel.Qty == "" ? oListItem["Qty"] : itemModel.Qty;
                    oListItem["Price_USD"] = itemModel.Price_USD == "" ? oListItem["Price_USD"] : itemModel.Price_USD;
                    oListItem["Price_THB"] = itemModel.Price_THB == "" ? oListItem["Price_THB"] : itemModel.Price_THB;
                    oListItem["Price_EUR"] = itemModel.Price_EUR == "" ? oListItem["Price_EUR"] : itemModel.Price_EUR;
                    oListItem["Part_Group"] = itemModel.Part_Group == "" ? oListItem["Part_Group"] : itemModel.Part_Group;
                    oListItem["Net_weight"] = itemModel.Net_Weight == "" ? oListItem["Net_weight"] : itemModel.Net_Weight;
                    oListItem["Type_No"] = itemModel.Type_No == "" ? oListItem["Type_No"] : itemModel.Type_No;
                    oListItem["Country_of_origin"] = itemModel.Country_Of_Origin == "" ? oListItem["Country_of_origin"] : itemModel.Country_Of_Origin;
                    oListItem["STC_mark"] = itemModel.STC_Mark == "" ? oListItem["STC_mark"] : itemModel.STC_Mark;
                    oListItem["EMC_code"] = itemModel.EMC_Code == "" ? oListItem["EMC_code"] : itemModel.EMC_Code;
                    oListItem["ECCN_code"] = itemModel.ECCN_Code == "" ? oListItem["ECCN_code"] : itemModel.ECCN_Code;
                    oListItem["PAGE_NO"] = itemModel.Page_No == "" ? oListItem["PAGE_NO"] : itemModel.Page_No;
                    oListItem["Substituted"] = itemModel.Substituted == "" ? oListItem["Substituted"] : itemModel.Substituted;
                    oListItem["Sub_Part_Price"] = itemModel.Sub_Part_Price == "" ? oListItem["Sub_Part_Price"] : itemModel.Sub_Part_Price;
                    DateTime outDateFinal = new DateTime();
                    DateTime outDateFirst = new DateTime();
                    try
                    {
                        if (itemModel.Final_Buy_Date != "")
                        {
                            var dtFinal = DateTime.TryParse(itemModel.Final_Buy_Date, out outDateFinal);
                            if (!dtFinal)
                            {
                                outDateFinal = DateTime.ParseExact(itemModel.Final_Buy_Date,
                                "yyyyMMdd",
                                 CultureInfo.InvariantCulture);
                                oListItem["Retention_Period"] = outDateFinal.Year + 8;
                            }
                            else
                            {
                                oListItem["Retention_Period"] = outDateFinal.Year + 8;
                            }
                        }
                        else if (itemModel.First_Production_Date != "")
                        {
                            var dtFirst = DateTime.TryParse(itemModel.First_Production_Date, out outDateFirst);
                            if (!dtFirst)
                            {
                                outDateFirst = DateTime.ParseExact(itemModel.First_Production_Date,
                                "yyyyMMdd",
                                 CultureInfo.InvariantCulture);
                                oListItem["Retention_Period"] = outDateFirst.Year + 8;
                            }
                            else
                            {
                                oListItem["Retention_Period"] = outDateFirst.Year + 8;
                            }
                        }
                    }
                    catch
                    {
                        oListItem["Retention_Period"] = itemModel.Retention_Period == "" ? oListItem["Retention_Period"] : itemModel.Retention_Period;
                    }
                    oListItem["Required_Per_Unit"] = itemModel.Required_Per_Unit == "" ? oListItem["Required_Per_Unit"] : itemModel.Required_Per_Unit;
                    oListItem["Final_Buy_Date"] = itemModel.Final_Buy_Date == "" ? oListItem["Final_Buy_Date"] : itemModel.Final_Buy_Date;
                    oListItem["Recomment"] = itemModel.Recomment == "" ? oListItem["Recomment"] : itemModel.Recomment;
                    oListItem["Lead_Time"] = itemModel.Lead_Time == "" ? oListItem["Lead_Time"] : itemModel.Lead_Time;
                    oListItem["Last_Production_date"] = itemModel.Last_Production_date == "" ? oListItem["Last_Production_date"] : itemModel.Last_Production_date;
                    oListItem["Zone_Code"] = itemModel.Zone_Code == "" ? oListItem["Zone_Code"] : itemModel.Zone_Code;
                    oListItem["IMPORT_FILE"] = fileName;
                    oListItem.Update();
                } 
            } 
        }

        public static void UpdateAllPart(Model partNoModel, string listName, string url)
        {
            var outPutData = new DataAccessModel();
            using (SPSite oSiteCollection = new SPSite(url))
            {
                using (SPWeb oWeb = oSiteCollection.OpenWeb())
                {
                    List<string> camlQuery = new List<string>();
                    camlQuery = GetPartCAML();
                    SPQuery query = new SPQuery();
                    query.Query = camlQuery[0].ToString();
                    query.ViewFields = camlQuery[1].ToString();
                    SPList oList = oWeb.Lists[listName];
                    SPListItemCollection oListItem = oList.GetItems(query);

                    var listPart = (from SPListItem p in oListItem
                                    where p["Part_no"] != null
                                    select p).ToList();

                    var listItem = (from SPListItem p in listPart
                                    where p["Part_no"].ToString() == partNoModel.Part_No
                                    select p).ToList();

                    foreach (var item in listItem)
                    {
                        item["RoHS"] = partNoModel.RoHS == "" ? item["RoHS"] : partNoModel.RoHS;
                        item["More_Description"] = partNoModel.Description == "" ? item["More_Description"] : partNoModel.Description;
                        item["Drawing_no"] = partNoModel.Drawing_NO == "" ? item["Drawing_no"] : partNoModel.Drawing_NO;
                        item["Qty"] = partNoModel.Qty == "" ? item["Qty"] : partNoModel.Qty;
                        item["Price_USD"] = partNoModel.Price_USD == "" ? item["Price_USD"] : partNoModel.Price_USD;
                        item["Price_THB"] = partNoModel.Price_THB == "" ? item["Price_THB"] : partNoModel.Price_THB;
                        item["Price_EUR"] = partNoModel.Price_EUR == "" ? item["Price_EUR"] : partNoModel.Price_EUR;
                        item["Part_Group"] = partNoModel.Part_Group == "" ? item["Part_Group"] : partNoModel.Part_Group;
                        item["Net_weight"] = partNoModel.Net_Weight == "" ? item["Net_weight"] : partNoModel.Net_Weight;
                        item["Type_No"] = partNoModel.Type_No == "" ? item["Type_No"] : partNoModel.Type_No;
                        item["Country_of_origin"] = partNoModel.Country_Of_Origin == "" ? item["Country_of_origin"] : partNoModel.Country_Of_Origin;
                        item["STC_mark"] = partNoModel.STC_Mark == "" ? item["STC_mark"] : partNoModel.STC_Mark;
                        item["EMC_code"] = partNoModel.EMC_Code == "" ? item["EMC_code"] : partNoModel.EMC_Code;
                        item["ECCN_code"] = partNoModel.ECCN_Code == "" ? item["ECCN_code"] : partNoModel.ECCN_Code;
                        item["PAGE_NO"] = partNoModel.Page_No == "" ? item["PAGE_NO"] : partNoModel.Page_No;
                        item["Substituted"] = partNoModel.Substituted == "" ? item["Substituted"] : partNoModel.Substituted;
                        item["Sub_Part_Price"] = partNoModel.Sub_Part_Price == "" ? item["Sub_Part_Price"] : partNoModel.Sub_Part_Price;
                        DateTime outDateFinal = new DateTime();
                        DateTime outDateFirst = new DateTime();
                        try
                        {
                            if (partNoModel.Final_Buy_Date != "")
                            {
                                var dtFinal = DateTime.TryParse(partNoModel.Final_Buy_Date, out outDateFinal);
                                if (!dtFinal)
                                {
                                    outDateFinal = DateTime.ParseExact(partNoModel.Final_Buy_Date,
                                    "yyyyMMdd",
                                     CultureInfo.InvariantCulture);
                                    item["Retention_Period"] = outDateFinal.Year + 8;
                                }
                                else
                                {
                                    item["Retention_Period"] = outDateFinal.Year + 8;
                                }
                            }
                            else if (partNoModel.First_Production_Date != "")
                            {
                                var dtFirst = DateTime.TryParse(partNoModel.First_Production_Date, out outDateFirst);
                                if (!dtFirst)
                                {
                                    outDateFirst = DateTime.ParseExact(partNoModel.First_Production_Date,
                                    "yyyyMMdd",
                                     CultureInfo.InvariantCulture);
                                    item["Retention_Period"] = outDateFirst.Year + 8;
                                }
                                else
                                {
                                    item["Retention_Period"] = outDateFirst.Year + 8;
                                }
                            }
                        }
                        catch
                        { 
                            item["Retention_Period"] = partNoModel.Retention_Period == "" ? item["Retention_Period"] : partNoModel.Retention_Period;
                        }
                        item["Required_Per_Unit"] = partNoModel.Required_Per_Unit == "" ? item["Required_Per_Unit"] : partNoModel.Required_Per_Unit;
                        item["Final_Buy_Date"] = partNoModel.Final_Buy_Date == "" ? item["Final_Buy_Date"] : partNoModel.Final_Buy_Date;
                        item["Recomment"] = partNoModel.Recomment == "" ? item["Recomment"] : partNoModel.Recomment;
                        item["Lead_Time"] = partNoModel.Lead_Time == "" ? item["Lead_Time"] : partNoModel.Lead_Time;
                        item["Last_Production_date"] = partNoModel.Last_Production_date == "" ? item["Last_Production_date"] : partNoModel.Last_Production_date;
                        item["Zone_Code"] = partNoModel.Zone_Code == "" ? item["Zone_Code"] : partNoModel.Zone_Code;
                        item.Update();
                    }
                } 
            } 
        }

        public static void UpdateAllModel(Model modelItem, string listName, string url)
        {
            DataAccessModel outPutData = new DataAccessModel();
            using (SPSite oSiteCollection = new SPSite(url))
            {
                using (SPWeb oWeb = oSiteCollection.OpenWeb())
                {
                    List<string> camlQuery = new List<string>();
                    camlQuery = GetPartCAML();
                    SPQuery query = new SPQuery();
                    query.Query = camlQuery[0].ToString();
                    query.ViewFields = camlQuery[1].ToString();
                    SPList oList = oWeb.Lists[listName];
                    SPListItemCollection oListItem = oList.GetItems(query);

                    var listPart = (from SPListItem p in oListItem
                                    where p["Model_Name"] != null
                                    select p).ToList();

                    var listItem = (from SPListItem p in listPart
                                    where p["Model_Name"].ToString() == modelItem.Model_Name
                                    select p).ToList();

                    foreach (var item in listItem)
                    {
                        item["Model_Description"] = modelItem.Model_Description == "" ? item["Model_Description"] : modelItem.Model_Description;
                        item["Category_Code"] = modelItem.Category_Code == "" ? item["Category_Code"] : modelItem.Category_Code;
                        item["First_Production_date"] = modelItem.First_Production_Date == "" ? item["First_Production_date"] : modelItem.First_Production_Date;
                        item["Exploded_Diagram_ref_no"] = modelItem.Exploded_Diagram_NO == "" ? item["Exploded_Diagram_ref_no"] : modelItem.Exploded_Diagram_NO;
                        item["Location_no"] = modelItem.Location_No == "" ? item["Location_no"] : modelItem.Location_No;
                        item.Update();
                    }


                }

            }


        }

        public static void DeleteModel(string ModelName, List<Model> listNotDel, string listName, string url)
        {
            using (SPSite oSiteCollection = new SPSite(url))
            {
                using (SPWeb oWeb = oSiteCollection.OpenWeb())
                {
                    List<string> camlQuery = new List<string>();
                    camlQuery = GetPartCAML();
                    SPQuery query = new SPQuery();
                    query.Query = camlQuery[0].ToString();
                    query.ViewFields = camlQuery[1].ToString();
                    SPList oList = oWeb.Lists[listName];
                    SPListItemCollection items = oList.GetItems(query);

                    var listItem = (from SPListItem p in items
                                    where p["Model_Name"] != null && p["Part_no"] != null
                                    select p).ToList();


                    var listModel = (from SPListItem p in listItem
                                     where p["Model_Name"].ToString() == ModelName
                                     select p).ToList();
                    List<SPListItem> listKeep = new List<SPListItem>();

                    foreach (var itemNotDel in listNotDel)
                    {
                        listKeep.AddRange((from SPListItem p in listModel
                                           where p["Part_no"].ToString() == itemNotDel.Part_No
                                           select p).ToList());
                    }
                    foreach (var itemKeep in listKeep)
                    {
                        listModel.Remove(itemKeep);
                    }
                    foreach (var item in listModel)
                    {

                        item.Delete();
                        item.Update();
                    }


                }
            }
        }

        public static void DeleteModelBL(string indoorName, string outdoorName, string listName, string url)
        {
            using (SPSite oSiteCollection = new SPSite(url))
            {
                using (SPWeb oWeb = oSiteCollection.OpenWeb())
                {
                    List<string> camlQuery = new List<string>();
                    camlQuery = GetModelBLCAML();
                    SPQuery query = new SPQuery();
                    query.Query = camlQuery[0].ToString();
                    query.ViewFields = camlQuery[1].ToString();
                    SPList oList = oWeb.Lists[listName];
                    SPListItemCollection items = oList.GetItems(query);
                    string caseDelete = "";
                    if (!string.IsNullOrEmpty(indoorName) && !string.IsNullOrEmpty(outdoorName))
                    {
                        caseDelete = "Both";
                    }
                    else if (string.IsNullOrEmpty(indoorName) && !string.IsNullOrEmpty(outdoorName))
                    {
                        caseDelete = "OnlyIndoor";
                    }
                    else if (!string.IsNullOrEmpty(indoorName) && string.IsNullOrEmpty(outdoorName))
                    {
                        caseDelete = "OnlyOutdoor";
                    }
                    else
                    {
                        caseDelete = "NoData";
                    }

                    switch (caseDelete)
                    {
                        case "Both":
                            var listItem = (from SPListItem p in items
                                            where p["INDOOR"] != null && p["OUTDOOR"] != null
                                            select p).ToList();
                            foreach (var item in listItem)
                            {
                                if (item["INDOOR"].ToString() == indoorName && item["outdoorName"].ToString() == outdoorName)
                                {
                                    item.Delete();
                                }

                            }
                            break;
                        case "OnlyIndoor":
                            var listItemIndoor = (from SPListItem p in items
                                                  where p["INDOOR"] != null
                                                  select p).ToList();
                            foreach (var item in listItemIndoor)
                            {
                                if (item["INDOOR"].ToString() == indoorName)
                                {
                                    item.Delete();
                                }

                            }
                            break;
                        case "OnlyOutdoor":
                            var listItemOutdoor = (from SPListItem p in items
                                                   where p["OUTDOOR"] != null
                                                   select p).ToList();
                            foreach (var item in listItemOutdoor)
                            {
                                if (item["OUTDOOR"].ToString() == outdoorName)
                                {
                                    item.Delete();
                                }

                            }
                            break;
                    }
                }
            }
        }

        #endregion

        #region ====== Service Manual ======

        public static DataAccessModel GetListServiceManualData(string listName, string url)
        {
            DataAccessModel outPutData = new DataAccessModel();
            try
            {
                using (SPSite oSite = new SPSite(url))
                {
                    using (SPWeb oWeb = oSite.OpenWeb())
                    {
                        List<string> camlQuery = new List<string>();
                        camlQuery = GetServiceManualCAML();

                        SPQuery query = new SPQuery();
                        query.Query = camlQuery[0].ToString();
                        query.ViewFields = camlQuery[1].ToString();
                        SPList oList = oWeb.Lists[listName];
                        SPListItemCollection items = oList.GetItems(query);
                        var dt = items.GetDataTable();
                        var list = (from table in dt.AsEnumerable()
                                    select new ServiceManual
                                    {
                                        Id = table["ID"].ToString(),
                                        Title = table["Title"].ToString(),
                                        SVM_FileName = table["SVM_FILENAME"].ToString(),
                                        Indoor_Model_Name = table["INDOOR_MODEL_NAME"].ToString(),
                                        Outdoor_Model_Name = table["OUTDOOR_MODEL_NAME"].ToString(),
                                        SVM_Remark = table["SVM_REMARK"].ToString(),
                                        Issue_Date = table["ISSUE_DATE"].ToString(),
                                        MDC_Code = table["MDC_CODE"].ToString(),
                                        Import_Filename = table["IMPORT_FILE"].ToString()

                                    }).ToList();

                        outPutData.listSVMData = list;
                        outPutData.Status = true;
                        outPutData.Reason = "Success";
                    }
                }
                return outPutData;
            }
            catch (Exception ex)
            {
                outPutData.listSVMData = new List<ServiceManual>();
                outPutData.Status = false;
                outPutData.Reason = ex.Message.ToString();
                return outPutData;
            }
        }

        public static void AddListServiceManual(ServiceManual itemSVM, string listName, string url, string fileName)
        {
            DataAccessModel outPutData = new DataAccessModel();

            using (SPSite oSiteCollection = new SPSite(url))
            {
                using (SPWeb oWeb = oSiteCollection.OpenWeb())
                {
                    SPList oList = oWeb.Lists[listName];

                    SPListItem oListItem = oList.Items.Add();
                    oListItem["Title"] = itemSVM.Title;
                    oListItem["MDC_CODE"] = itemSVM.MDC_Code;
                    oListItem["SVM_REMARK"] = itemSVM.SVM_Remark;
                    oListItem["OUTDOOR_MODEL_NAME"] = itemSVM.Outdoor_Model_Name;
                    oListItem["INDOOR_MODEL_NAME"] = itemSVM.Indoor_Model_Name;
                    oListItem["ISSUE_DATE"] = itemSVM.Issue_Date;
                    oListItem["SVM_FILENAME"] = itemSVM.SVM_FileName;
                    oListItem["IMPORT_FILE"] = fileName;
                    oListItem.Update();

                }
            }
        }

        public static void UpdateListServiceManual(ServiceManual oldItemSVM, ServiceManual itemSVM, string listName, string url, string fileName)
        {
            DataAccessModel outPutData = new DataAccessModel();
            using (SPSite oSiteCollection = new SPSite(url))
            {
                using (SPWeb oWeb = oSiteCollection.OpenWeb())
                {
                    SPList oList = oWeb.Lists[listName];
                    SPListItem oListItem = oList.GetItemById(Convert.ToInt32(oldItemSVM.Id));
                    oListItem["MDC_CODE"] = itemSVM.MDC_Code;
                    oListItem["SVM_REMARK"] = itemSVM.SVM_Remark;
                    oListItem["ISSUE_DATE"] = itemSVM.Issue_Date;
                    oListItem["SVM_FILENAME"] = itemSVM.SVM_FileName;
                    oListItem["IMPORT_FILE"] = fileName;
                    oListItem.Update();
                }
            }
        }

        #endregion

        #region ====== Business Louge ======

        public static DataAccessModel GetListModelBLData(string listName, string url)
        {
            DataAccessModel outPutData = new DataAccessModel();
            try
            {

                using (SPSite oSite = new SPSite(url))
                {
                    using (SPWeb oWeb = oSite.OpenWeb())
                    {
                        List<string> camlQuery = new List<string>();
                        camlQuery = GetModelBLCAML();

                        SPQuery query = new SPQuery();
                        query.Query = camlQuery[0].ToString();
                        query.ViewFields = camlQuery[1].ToString();
                        SPList oList = oWeb.Lists[listName];
                        SPListItemCollection items = oList.GetItems(query);

                        var dt = items.GetDataTable();
                        var list = (from table in dt.AsEnumerable()
                                    select new ModelBL
                                    {
                                        Id = table["Id"].ToString(),
                                        Title = table["Title"].ToString(),
                                        BL_CATEGORY = table["BL_CATEGORY"].ToString(),
                                        BRAND = table["BRAND"].ToString(),
                                        BL_PRODUCT_TYPE = table["BL_PRODUCT_TYPE"].ToString(),
                                        BL_PRODUCT_SIZE = table["BL_PRODUCT_SIZE"].ToString(),
                                        REFRIGERANT = table["REFRIGERANT"].ToString(),
                                        INDOOR = table["INDOOR"].ToString(),
                                        OUTDOOR = table["OUTDOOR"].ToString(),
                                        INSTALLATION = table["INSTALLATION"].ToString(),
                                        OWNER = table["OWNER"].ToString(),
                                        DISC = table["DISC"].ToString(),
                                        SPECIFICATION = table["SPECIFICATION"].ToString(),
                                        BULLETIN = table["BULLETIN"].ToString(),
                                        DATABOOK = table["DATABOOK"].ToString(),
                                        VDO = table["VDO"].ToString(),
                                        PRESENTATION = table["PRESENTATION"].ToString(),
                                        IMAGE_LOW = table["IMAGE_LOW"].ToString(),
                                        IMAGE_HD = table["IMAGE_HD"].ToString(),
                                        CATALOGUE = table["CATALOGUE"].ToString()

                                    }).ToList();

                        outPutData.listModelBLData = list;
                        outPutData.Status = true;
                        outPutData.Reason = "Success";

                    }
                }
                return outPutData;
            }
            catch (Exception ex)
            {
                outPutData.listModelBLData = new List<ModelBL>();
                outPutData.Status = false;
                outPutData.Reason = ex.Message.ToString();
                return outPutData;
            }
        }

        public static void AddListModelBL(ModelBL itemBL, string listName, string url, string fileName)
        {
            var outPutData = new DataAccessModel();
            using (SPSite oSiteCollection = new SPSite(url))
            {
                using (SPWeb oWeb = oSiteCollection.OpenWeb())
                {
                    SPList oList = oWeb.Lists[listName];

                    SPListItem oListItem = oList.Items.Add();
                    oListItem["Title"] = itemBL.Title;
                    oListItem["BL_CATEGORY"] = itemBL.BL_CATEGORY;
                    oListItem["BRAND"] = itemBL.BRAND;
                    oListItem["BL_PRODUCT_TYPE"] = itemBL.BL_PRODUCT_TYPE;
                    oListItem["BL_PRODUCT_SIZE"] = itemBL.BL_PRODUCT_SIZE;
                    oListItem["REFRIGERANT"] = itemBL.REFRIGERANT;
                    oListItem["INDOOR"] = itemBL.INDOOR;
                    oListItem["OUTDOOR"] = itemBL.OUTDOOR;
                    oListItem["INSTALLATION"] = itemBL.INSTALLATION;
                    oListItem["OWNER"] = itemBL.OWNER;
                    oListItem["DISC"] = itemBL.DISC;
                    oListItem["SPECIFICATION"] = itemBL.SPECIFICATION;
                    oListItem["BULLETIN"] = itemBL.BULLETIN;
                    oListItem["DATABOOK"] = itemBL.DATABOOK;
                    oListItem["VDO"] = itemBL.VDO;
                    oListItem["PRESENTATION"] = itemBL.PRESENTATION;
                    oListItem["IMAGE_LOW"] = itemBL.IMAGE_LOW;
                    oListItem["CATALOGUE"] = itemBL.CATALOGUE;
                    oListItem["IMPORT_FILE"] = fileName;

                    oListItem.Update();
                }
            }

        }

        public static void UpdateListModelBL(ModelBL oldItemBL, ModelBL itemBL, string listName, string url, string fileName)
        {
            DataAccessModel outPutData = new DataAccessModel();
            using (SPSite oSiteCollection = new SPSite(url))
            {
                using (SPWeb oWeb = oSiteCollection.OpenWeb())
                {
                    SPList oList = oWeb.Lists[listName];
                    SPListItem oListItem = oList.GetItemById(Convert.ToInt32(oldItemBL.Id));

                    oListItem["BL_CATEGORY"] = itemBL.BL_CATEGORY;
                    oListItem["BRAND"] = itemBL.BRAND;
                    oListItem["BL_PRODUCT_TYPE"] = itemBL.BL_PRODUCT_TYPE;
                    oListItem["BL_PRODUCT_SIZE"] = itemBL.BL_PRODUCT_SIZE;
                    oListItem["REFRIGERANT"] = itemBL.REFRIGERANT;
                    oListItem["INDOOR"] = itemBL.INDOOR;
                    oListItem["OUTDOOR"] = itemBL.OUTDOOR;
                    oListItem["INSTALLATION"] = itemBL.INSTALLATION;
                    oListItem["OWNER"] = itemBL.OWNER;
                    oListItem["DISC"] = itemBL.DISC;
                    oListItem["SPECIFICATION"] = itemBL.SPECIFICATION;
                    oListItem["BULLETIN"] = itemBL.BULLETIN;
                    oListItem["DATABOOK"] = itemBL.DATABOOK;
                    oListItem["VDO"] = itemBL.VDO;
                    oListItem["PRESENTATION"] = itemBL.PRESENTATION;
                    oListItem["IMAGE_LOW"] = itemBL.IMAGE_LOW;
                    oListItem["CATALOGUE"] = itemBL.CATALOGUE;
                    oListItem["IMPORT_FILE"] = fileName;
                    oListItem.Update();
                }
            }
        }

        #endregion

        public static List<string> GetPartCAML()
        {
            List<string> query = new List<string>();

            StringBuilder camlQuery = new StringBuilder();
            camlQuery.Append("<View><Query><Where><Eq>");
            camlQuery.Append("<FieldRef Name='Title'/>");
            camlQuery.Append("<Value Type='Text'>All</Value>");
            camlQuery.Append("</Eq></Where></Query></View>");
            query.Add(camlQuery.ToString());

            StringBuilder camlField = new StringBuilder();
            camlField.Append("<FieldRef Name='ID' />");
            camlField.Append("<FieldRef Name='Title' />");
            camlField.Append("<FieldRef Name='Model_Name' />");
            camlField.Append("<FieldRef Name='Model_Description' />");
            camlField.Append("<FieldRef Name='Category_Code' />");
            camlField.Append("<FieldRef Name='First_Production_date' />");
            camlField.Append("<FieldRef Name='Exploded_Diagram_ref_no' />");
            camlField.Append("<FieldRef Name='Location_no' />");
            camlField.Append("<FieldRef Name='Part_no' />");
            camlField.Append("<FieldRef Name='RoHS' />");
            camlField.Append("<FieldRef Name='More_Description' />");
            camlField.Append("<FieldRef Name='Drawing_no' />");
            camlField.Append("<FieldRef Name='Qty' />");
            camlField.Append("<FieldRef Name='Price_USD' />");
            camlField.Append("<FieldRef Name='Price_THB' />");
            camlField.Append("<FieldRef Name='Price_EUR' />");
            camlField.Append("<FieldRef Name='Part_Group' />");
            camlField.Append("<FieldRef Name='Net_weight' />");
            camlField.Append("<FieldRef Name='Type_No' />");
            camlField.Append("<FieldRef Name='Country_of_origin' />");
            camlField.Append("<FieldRef Name='STC_mark' />");
            camlField.Append("<FieldRef Name='EMC_code' />");
            camlField.Append("<FieldRef Name='ECCN_code' />");
            camlField.Append("<FieldRef Name='PAGE_NO' />");
            camlField.Append("<FieldRef Name='Substituted' />");
            camlField.Append("<FieldRef Name='Sub_Part_Price' />");
            camlField.Append("<FieldRef Name='Retention_Period' />");
            camlField.Append("<FieldRef Name='Required_Per_Unit' />");
            camlField.Append("<FieldRef Name='Final_Buy_Date' />");
            camlField.Append("<FieldRef Name='Recomment' />");
            camlField.Append("<FieldRef Name='Lead_Time' />");
            camlField.Append("<FieldRef Name='Last_Production_date' />");
            camlField.Append("<FieldRef Name='Zone_Code' />");
            camlField.Append("<FieldRef Name='IMPORT_FILE' />");

            query.Add(camlField.ToString());

            return query;
        }
    
        public static List<string> GetServiceManualCAML()
        {
            List<string> query = new List<string>();

            StringBuilder camlQuery = new StringBuilder();
            camlQuery.Append("<View><Query><Where><Eq>");
            camlQuery.Append("<FieldRef Name='Title'/>");
            camlQuery.Append("<Value Type='Text'>All</Value>");
            camlQuery.Append("</Eq></Where></Query></View>");
            query.Add(camlQuery.ToString());

            StringBuilder camlField = new StringBuilder();
            camlField.Append("<FieldRef Name='ID' />");
            camlField.Append("<FieldRef Name='Title' />");
            camlField.Append("<FieldRef Name='INDOOR_MODEL_NAME' />");
            camlField.Append("<FieldRef Name='OUTDOOR_MODEL_NAME' />");
            camlField.Append("<FieldRef Name='SVM_FILENAME' />");
            camlField.Append("<FieldRef Name='MDC_CODE' />");
            camlField.Append("<FieldRef Name='ISSUE_DATE' />");
            camlField.Append("<FieldRef Name='SVM_REMARK' />");
            camlField.Append("<FieldRef Name='IMPORT_FILE' />");

            query.Add(camlField.ToString());

            return query;
        }

        public static List<string> GetModelBLCAML()
        {
            List<string> query = new List<string>();

            StringBuilder camlQuery = new StringBuilder();
            camlQuery.Append("<View><Query><Where><Eq>");
            camlQuery.Append("<FieldRef Name='Title'/>");
            camlQuery.Append("<Value Type='Text'>All</Value>");
            camlQuery.Append("</Eq></Where></Query></View>");
            query.Add(camlQuery.ToString());

            StringBuilder camlField = new StringBuilder();
            camlField.Append("<FieldRef Name='Id' />");
            camlField.Append("<FieldRef Name='Title' />");
            camlField.Append("<FieldRef Name='BL_CATEGORY' />");
            camlField.Append("<FieldRef Name='BRAND' />");
            camlField.Append("<FieldRef Name='BL_PRODUCT_TYPE' />");
            camlField.Append("<FieldRef Name='BL_PRODUCT_SIZE' />");
            camlField.Append("<FieldRef Name='REFRIGERANT' />");
            camlField.Append("<FieldRef Name='INDOOR' />");
            camlField.Append("<FieldRef Name='OUTDOOR' />");
            camlField.Append("<FieldRef Name='INSTALLATION' />");
            camlField.Append("<FieldRef Name='OWNER' />");
            camlField.Append("<FieldRef Name='DISC' />");
            camlField.Append("<FieldRef Name='SPECIFICATION' />");
            camlField.Append("<FieldRef Name='BULLETIN' />");
            camlField.Append("<FieldRef Name='DATABOOK' />");
            camlField.Append("<FieldRef Name='VDO' />");
            camlField.Append("<FieldRef Name='PRESENTATION' />");
            camlField.Append("<FieldRef Name='IMAGE_HD' />");
            camlField.Append("<FieldRef Name='IMAGE_LOW' />");
            camlField.Append("<FieldRef Name='CATALOGUE' />");
            camlField.Append("<FieldRef Name='IMPORT_FILE' />");
            query.Add(camlField.ToString());

            return query;
        }

    }
}
