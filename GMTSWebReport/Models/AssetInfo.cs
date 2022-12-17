using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GMTSWebReport.Models
{
    public class AssetInfo
    {
        public string Company { get; set; }
        public string ComSortName { get; set; }
        public string Unit { get; set; }
        public int Unit_Order { get; set; }
        public string Item_Description { get; set; }
      
        public string AssetNo { get; set; }
        public string AssetType { get; set; }
        public string Item { get; set; }
        public decimal Qty { get; set; }
        public decimal Price { get; set; }
        public string EmployeeCode { get; set; }
        public string Employee { get; set; }
        public string Department { get; set; }
        public string Designation { get; set; }
        public string ItemCategory { get; set; }
        public int ItemCategoryDisplayOrder { get; set; }
        public string Custodian { get; set; }
        public string SubCustodian { get; set; }
        public int Display_Order { get; set; }
        public string Supplier { get; set; }
        public int NoOfUser { get; set; }
        public decimal AMCAmount { get; set; }
        public bool IsMcApplicable { get; set; }
    }

  
}