using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace PS.DataModel
{
    public class DataAccessModel
    {
        public List<Model> listModelData { get; set; }

        public List<ServiceManual> listSVMData { get; set; }

        public List<ModelBL> listModelBLData { get; set; }
        public bool Status { get; set; }

        public string Reason { get; set; }
    }
    public class Model
    {
        public string Id { get; set; }
        public string Model_Name { get; set; }

        public string Model_Description { get; set; }
        public string Description { get; set; }
        public string Category_Code { get; set; }
        public string First_Production_Date { get; set; }
        public string Exploded_Diagram_NO { get; set; }
        public string Location_No { get; set; }
        public string Part_No { get; set; }
        public string RoHS { get; set; }
        public string Information { get; set; }
        public string Drawing_NO { get; set; }
        public string Qty { get; set; }
        public string Price_USD { get; set; }
        public string Price_THB { get; set; }
        public string Price_EUR { get; set; }
        public string Part_Group { get; set; }
        public string Net_Weight { get; set; }
        public string Type_No { get; set; }
        public string Country_Of_Origin { get; set; }
        public string STC_Mark { get; set; }
        public string EMC_Code { get; set; }
        public string ECCN_Code { get; set; }

        public string Page_No { get; set; }
        public string Substituted { get; set; }
        public string Sub_Part_Price { get; set; }
        public string Retention_Period { get; set; }
        public string Required_Per_Unit { get; set; }
        public string Final_Buy_Date { get; set; }
        public string Recomment { get; set; }
        public string Lead_Time { get; set; }
        public string Last_Production_date { get; set; }
        public string Zone_Code { get; set; } 
        public string Import_Filename { get; set; }


    }

    public class ServiceManual
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string Indoor_Model_Name { get; set; }
        public string Outdoor_Model_Name { get; set; }
        public string Issue_Date { get; set; }
        public string SVM_Remark { get; set; }
        public string SVM_FileName { get; set; }
        public string MDC_Code { get; set; }
        public string Import_Filename { get; set; }

    }
    public class ModelBL
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string BL_CATEGORY { get; set; }
        public string BRAND { get; set; }
        public string BL_PRODUCT_TYPE { get; set; }
        public string BL_PRODUCT_SIZE { get; set; }
        public string REFRIGERANT { get; set; }
        public string INDOOR { get; set; }
        public string OUTDOOR { get; set; }
        public string INSTALLATION { get; set; }
        public string OWNER { get; set; }
        public string DISC { get; set; }
        public string SPECIFICATION { get; set; }
        public string BULLETIN { get; set; }
        public string DATABOOK { get; set; }
        public string VDO { get; set; }
        public string PRESENTATION { get; set; }
        public string IMAGE_HD { get; set; }
        public string IMAGE_LOW { get; set; }
        public string CATALOGUE { get; set; }
        public string Import_Filename { get; set; }
    }
}
