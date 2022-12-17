using GMTSWebReport.Models;
using Microsoft.Reporting.WebForms;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GMTSWebReport.Controllers
{
    public class AssetController : Controller
    {
        // GET: Asset
        public ActionResult Index()
        {
            return View();
        }


        public ActionResult AssetReportByCompanyRegister()
        {
            return View();
        }

        public ActionResult AssetReportByCompanyUnit()
        {
            return View();
        }

        public ActionResult AssetReportByCompany()
        {
            return View();
        }


        public ActionResult AssetReportByItemCategory()
        {
            return View();
        }

        public ActionResult AssetReportByItemCategoryStock()
        {
            return View();
        }


        public ActionResult TotalAssetsValueReport()
        {
            return View();
        }

        public ActionResult InTransibleAssetsReport()
        {
            return View();
        }


        public ActionResult AssetReportByDepartment()
        {
            return View();
        }




        private List<AssetInfo> GetIntransibleAssetValueDataList()
        {
            var objList = new List<AssetInfo>();

            using (SqlConnection con = new SqlConnection(SQLHelper.GetConnectionString()))
            {
                con.Open();
                var connection = con;
                var transaction = connection.BeginTransaction();
                try
                {
                    Hashtable htbl = new Hashtable();

                    DataSet dset = new DataSet();
                    dset = SQLHelper.ExecuteProcedureAsFromDataAdapter("SPGetIntransibleAssetItemInfo", htbl, connection, transaction);
                    transaction.Commit();
                    DataTable dtResult = dset.Tables[0];

                    foreach (DataRow dr in dtResult.Rows)
                    {
                        AssetInfo obj = new AssetInfo();

                        obj.Company = dr["Company"].ToString();
                        obj.ComSortName = dr["ComSortName"].ToString();
                        obj.Item = dr["Item"].ToString();
                        obj.Supplier = dr["Supplier"].ToString();

                        if (dr["NoOfUser"].ToString() != "")
                            obj.NoOfUser = Convert.ToInt32(dr["NoOfUser"].ToString());

                        if (dr["IsMcApplicable"].ToString() != "")
                            obj.IsMcApplicable = Convert.ToBoolean(dr["IsMcApplicable"].ToString());

                        if (dr["Price"].ToString() != "")
                            obj.Price = Convert.ToDecimal(dr["Price"].ToString());
                        if (dr["AMCAmount"].ToString() != "")
                            obj.AMCAmount = Convert.ToDecimal(dr["AMCAmount"].ToString());
                        objList.Add(obj);
                    }
                }
                catch (Exception ex)
                {
                    string errormsg = ex.Message;
                    transaction.Rollback();
                }
                con.Close();
            }
            return objList;
        }

        private List<AssetInfo> GetTotalAssetValueDataList()
        {
            var objList = new List<AssetInfo>();

            using (SqlConnection con = new SqlConnection(SQLHelper.GetConnectionString()))
            {
                con.Open();
                var connection = con;
                var transaction = connection.BeginTransaction();
                try
                {
                    Hashtable htbl = new Hashtable();

                    DataSet dset = new DataSet();
                    dset = SQLHelper.ExecuteProcedureAsFromDataAdapter("SPTotalAssetsValue", htbl, connection, transaction);
                    transaction.Commit();
                    DataTable dtResult = dset.Tables[0];

                    foreach (DataRow dr in dtResult.Rows)
                    {
                        AssetInfo obj = new AssetInfo();

                        obj.AssetType = dr["AssetType"].ToString();
                        if (dr["Price"].ToString() != "")
                            obj.Price = Convert.ToDecimal(dr["Price"].ToString());
                        objList.Add(obj);
                    }
                }
                catch (Exception ex)
                {
                    string errormsg = ex.Message;
                    transaction.Rollback();
                }
                con.Close();
            }
            return objList;
        }

        private List<AssetInfo> GetAssetReportDataList(int? CompanyId, int? UnitId , int? DepartmentId , int? CustodianId, int? SubCustodianId, int? AssetItemId)
        {
            var objList = new List<AssetInfo>();

            using (SqlConnection con = new SqlConnection(SQLHelper.GetConnectionString()))
            {
                con.Open();
                var connection = con;
                var transaction = connection.BeginTransaction();
                try
                {
                    Hashtable htbl = new Hashtable();
                    htbl.Add("@CompanyId", CompanyId);
                    htbl.Add("@UnitId", UnitId);
                    htbl.Add("@DepartmentId", DepartmentId);
                    htbl.Add("@CustodianId", CustodianId);
                    htbl.Add("@SubCustodianId", SubCustodianId);
                    htbl.Add("@AssetItemId", AssetItemId);
                    //htbl.Add("@LocationID", LocationID);

                    DataSet dset = new DataSet();
                    dset = SQLHelper.ExecuteProcedureAsFromDataAdapter("SPGetAssetBookInfo", htbl, connection, transaction);
                    transaction.Commit();
                    DataTable dtResult = dset.Tables[0];

                    foreach (DataRow dr in dtResult.Rows)
                    {
                        AssetInfo obj = new AssetInfo();

                        obj.Company = dr["Company"].ToString();

                        obj.ComSortName = dr["ComSortName"].ToString();

                        obj.Unit = dr["Unit"].ToString();
                        obj.Custodian = dr["Custodian"].ToString();
                        obj.SubCustodian = dr["SubCustodian"].ToString();
                        obj.AssetNo = dr["AssetNo"].ToString();

                        obj.ItemCategory = dr["ItemCategory"].ToString();
                        obj.EmployeeCode = dr["EmployeeCode"].ToString();
                        obj.Employee = dr["Employee"].ToString();
                        obj.Department = dr["Department"].ToString();
                        obj.Designation = dr["Designation"].ToString();
                        obj.Item_Description = dr["Item_Description"].ToString();
                        if (dr["Unit_Order"].ToString() != "")
                            obj.Unit_Order = Convert.ToInt32(dr["Unit_Order"].ToString());


                        if (dr["Display_Order"].ToString() != "")
                            obj.Display_Order = Convert.ToInt32(dr["Display_Order"].ToString());
                        if (dr["Qty"].ToString() != "")
                            obj.Qty = Convert.ToDecimal(dr["Qty"].ToString());
                        obj.Item = dr["Item"].ToString();

                        if (dr["Price"].ToString() != "")
                            obj.Price = Convert.ToDecimal(dr["Price"].ToString());

                        objList.Add(obj);
                    }
                }
                catch (Exception ex)
                {
                    string errormsg = ex.Message;
                    transaction.Rollback();
                }
                con.Close();
            }
            return objList;
        }

        private List<AssetInfo> GetAssetStockReportDataList(int? CompanyId, int? AssetItemId)
        {
            var objList = new List<AssetInfo>();

            using (SqlConnection con = new SqlConnection(SQLHelper.GetConnectionString()))
            {
                con.Open();
                var connection = con;
                var transaction = connection.BeginTransaction();
                try
                {
                    Hashtable htbl = new Hashtable();
                    htbl.Add("@CompanyId", CompanyId);
                    htbl.Add("@AssetItemId", AssetItemId);
                    //htbl.Add("@LocationID", LocationID);

                    DataSet dset = new DataSet();
                    dset = SQLHelper.ExecuteProcedureAsFromDataAdapter("SPGetAssetBookStockInfo", htbl, connection, transaction);
                    transaction.Commit();
                    DataTable dtResult = dset.Tables[0];

                    foreach (DataRow dr in dtResult.Rows)
                    {
                        AssetInfo obj = new AssetInfo();

                        obj.Company = dr["Company"].ToString();

                        obj.ComSortName = dr["ComSortName"].ToString();

                        obj.Unit = dr["Unit"].ToString();
                        obj.Custodian = dr["Custodian"].ToString();
                        obj.SubCustodian = dr["SubCustodian"].ToString();
                        obj.AssetNo = dr["AssetNo"].ToString();

                        obj.ItemCategory = dr["ItemCategory"].ToString();
                        obj.EmployeeCode = dr["EmployeeCode"].ToString();
                        obj.Employee = dr["Employee"].ToString();
                        obj.Department = dr["Department"].ToString();
                        obj.Designation = dr["Designation"].ToString();
                        obj.Item_Description = dr["Item_Description"].ToString();
                        if (dr["Unit_Order"].ToString() != "")
                            obj.Unit_Order = Convert.ToInt32(dr["Unit_Order"].ToString());


                        if (dr["Display_Order"].ToString() != "")
                            obj.Display_Order = Convert.ToInt32(dr["Display_Order"].ToString());
                        if (dr["Qty"].ToString() != "")
                            obj.Qty = Convert.ToDecimal(dr["Qty"].ToString());
                        obj.Item = dr["Item"].ToString();

                        if (dr["Price"].ToString() != "")
                            obj.Price = Convert.ToDecimal(dr["Price"].ToString());

                        objList.Add(obj);
                    }
                }
                catch (Exception ex)
                {
                    string errormsg = ex.Message;
                    transaction.Rollback();
                }
                con.Close();
            }
            return objList;
        }

        //Company Register
        public ActionResult AssetReportByCompanyRegisterExport(string FileType = "Excel", string ContentType = "application/vnd.ms-excel", int? CompanyId = 0, int? UnitId = 0, 
            int? DepartmentId = 0, int? CustodianId = 0, int? SubCustodianId = 0, int? AssetItemId = 0)
        {
            try
            {
                DateTime dTime = DateTime.Now;
                string strExelFileName = string.Format("AssetReportByCompanyRegister");
                string savePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\downloads\\";
                string pathToCheck = savePath + strExelFileName;
                var ExportDate = DateTime.Today.Date;

                List<AssetInfo> assetData = GetAssetReportDataList(CompanyId, UnitId, DepartmentId, CustodianId, SubCustodianId, AssetItemId);

                string _ReportName = "AssetReportByCompanyRegister.rdlc";
                DataTable tbl = new ReportDS.AssetReportByCustodianDataTable();
                if (assetData != null)
                {
                    tbl = ReportHelper.GetDataTableFromList(assetData, tbl);
                }
                string reportPath = string.Format("~/Content/Reports/Assets/{0}", _ReportName);

                return ExportData(Server.MapPath(reportPath), strExelFileName, tbl, FileType, ContentType);
            }
            catch (Exception ex)
            {

                return Json(new { Status = "Error" }, JsonRequestBehavior.AllowGet);
            }
        }


        //Unit
        public ActionResult AssetReportByCompanyUnitExport(string FileType = "Excel", string ContentType = "application/vnd.ms-excel", int? CompanyId = 0,
            int? UnitId = 0, int? DepartmentId = 0, int? CustodianId = 0, int? SubCustodianId = 0, int? AssetItemId = 0)
        {
            try
            {
                DateTime dTime = DateTime.Now;
                string strExelFileName = string.Format("AssetReportByCompanyUnit");
                string savePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\downloads\\";
                string pathToCheck = savePath + strExelFileName;
                var ExportDate = DateTime.Today.Date;

                List<AssetInfo> assetData = GetAssetReportDataList(CompanyId, UnitId, DepartmentId, CustodianId, SubCustodianId, AssetItemId);

                string _ReportName = "AssetReportByCompanyUnit.rdlc";
                DataTable tbl = new ReportDS.AssetReportByCustodianDataTable();
                if (assetData != null)
                {
                    tbl = ReportHelper.GetDataTableFromList(assetData, tbl);
                }
                string reportPath = string.Format("~/Content/Reports/Assets/{0}", _ReportName);

                return ExportData(Server.MapPath(reportPath), strExelFileName, tbl, FileType, ContentType);
            }
            catch (Exception ex)
            {

                return Json(new { Status = "Error" }, JsonRequestBehavior.AllowGet);
            }
        }

        //Company
        public ActionResult AssetReportByCompanyExport(string FileType = "Excel", string ContentType = "application/vnd.ms-excel", int? CompanyId = 0,
              int? UnitId = 0, int? DepartmentId = 0, int? CustodianId = 0, int? SubCustodianId = 0, int? AssetItemId = 0)
        {
            try
            {
                DateTime dTime = DateTime.Now;
                string strExelFileName = string.Format("AssetReportByCompany");
                string savePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\downloads\\";
                string pathToCheck = savePath + strExelFileName;
                var ExportDate = DateTime.Today.Date;

                List<AssetInfo> assetData = GetAssetReportDataList(CompanyId, UnitId, DepartmentId, CustodianId, SubCustodianId, AssetItemId);

                string _ReportName = "AssetReportByCompany.rdlc";
                DataTable tbl = new ReportDS.AssetReportByCustodianDataTable();
                if (assetData != null)
                {
                    tbl = ReportHelper.GetDataTableFromList(assetData, tbl);
                }
                string reportPath = string.Format("~/Content/Reports/Assets/{0}", _ReportName);

                return ExportData(Server.MapPath(reportPath), strExelFileName, tbl, FileType, ContentType);
            }
            catch (Exception ex)
            {

                return Json(new { Status = "Error" }, JsonRequestBehavior.AllowGet);
            }
        }

        //Item Category
        public ActionResult AssetReportByItemCategoryExport(string FileType = "Excel", string ContentType = "application/vnd.ms-excel", int? CompanyId = 0,
              int? UnitId = 0, int? DepartmentId = 0, int? CustodianId = 0, int? SubCustodianId = 0, int? AssetItemId = 0)
        {
            try
            {
                DateTime dTime = DateTime.Now;
                string strExelFileName = string.Format("AssetReportByItemCategory");
                string savePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\downloads\\";
                string pathToCheck = savePath + strExelFileName;
                var ExportDate = DateTime.Today.Date;

                List<AssetInfo> assetData = GetAssetReportDataList(CompanyId, UnitId, DepartmentId, CustodianId, SubCustodianId, AssetItemId);

                string _ReportName = "AssetReportByItemCategory.rdlc";
                DataTable tbl = new ReportDS.AssetReportByCustodianDataTable();
                if (assetData != null)
                {
                    tbl = ReportHelper.GetDataTableFromList(assetData, tbl);
                }
                string reportPath = string.Format("~/Content/Reports/Assets/{0}", _ReportName);

                return ExportData(Server.MapPath(reportPath), strExelFileName, tbl, FileType, ContentType);
            }
            catch (Exception ex)
            {

                return Json(new { Status = "Error" }, JsonRequestBehavior.AllowGet);
            }
        }

        //Item Category Stock
        public ActionResult AssetReportByItemCategoryStockExport(string FileType = "Excel", string ContentType = "application/vnd.ms-excel", int? CompanyId = 0, int? AssetItemId = 0)
        {
            try
            {
                DateTime dTime = DateTime.Now;
                string strExelFileName = string.Format("AssetReportByItemCategoryStock");
                string savePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\downloads\\";
                string pathToCheck = savePath + strExelFileName;
                var ExportDate = DateTime.Today.Date;

                List<AssetInfo> assetData = GetAssetStockReportDataList(CompanyId, AssetItemId);

                string _ReportName = "AssetReportByItemCategoryStock.rdlc";
                DataTable tbl = new ReportDS.AssetReportByCustodianDataTable();
                if (assetData != null)
                {
                    tbl = ReportHelper.GetDataTableFromList(assetData, tbl);
                }
                string reportPath = string.Format("~/Content/Reports/Assets/{0}", _ReportName);

                return ExportData(Server.MapPath(reportPath), strExelFileName, tbl, FileType, ContentType);
            }
            catch (Exception ex)
            {

                return Json(new { Status = "Error" }, JsonRequestBehavior.AllowGet);
            }
        }


        //Total Assets Value
        public ActionResult TotalAssetsReportExport(string FileType = "Excel", string ContentType = "application/vnd.ms-excel")
        {
            try
            {
                DateTime dTime = DateTime.Now;
                string strExelFileName = string.Format("TotalAssetsValueReport");
                string savePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\downloads\\";
                string pathToCheck = savePath + strExelFileName;
                var ExportDate = DateTime.Today.Date;

                List<AssetInfo> assetData = GetTotalAssetValueDataList();

                string _ReportName = "TotalAssetsValueReport.rdlc";
                DataTable tbl = new ReportDS.AssetReportByCustodianDataTable();
                if (assetData != null)
                {
                    tbl = ReportHelper.GetDataTableFromList(assetData, tbl);
                }
                string reportPath = string.Format("~/Content/Reports/Assets/{0}", _ReportName);

                return ExportData(Server.MapPath(reportPath), strExelFileName, tbl, FileType, ContentType);
            }
            catch (Exception ex)
            {
                return Json(new { Status = "Error" }, JsonRequestBehavior.AllowGet);
            }
        }

        //Intransible Assets Value
        public ActionResult IntransibleAssetReportExport(string FileType = "Excel", string ContentType = "application/vnd.ms-excel")
        {
            try
            {
                DateTime dTime = DateTime.Now;
                string strExelFileName = string.Format("IntansibleAssetReport");
                string savePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\downloads\\";
                string pathToCheck = savePath + strExelFileName;
                var ExportDate = DateTime.Today.Date;

                List<AssetInfo> assetData = GetIntransibleAssetValueDataList();

                string _ReportName = "IntansibleAssetReport.rdlc";
                DataTable tbl = new ReportDS.AssetReportByCustodianDataTable();
                if (assetData != null)
                {
                    tbl = ReportHelper.GetDataTableFromList(assetData, tbl);
                }
                string reportPath = string.Format("~/Content/Reports/Assets/{0}", _ReportName);

                return ExportData(Server.MapPath(reportPath), strExelFileName, tbl, FileType, ContentType);
            }
            catch (Exception ex)
            {
                return Json(new { Status = "Error" }, JsonRequestBehavior.AllowGet);
            }
        }

        //Custodian wise report
        public ActionResult AssetReportByDepartmentExport(string FileType = "Excel", string ContentType = "application/vnd.ms-excel", int? CompanyId = 0,
            int? UnitId = 0, int? DepartmentId = 0, int? CustodianId = 0, int? SubCustodianId = 0, int? AssetItemId = 0)
        {
            try
            {
                DateTime dTime = DateTime.Now;
                string strExelFileName = string.Format("AssetReportByDepartment");
                string savePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\downloads\\";
                string pathToCheck = savePath + strExelFileName;
                var ExportDate = DateTime.Today.Date;

                List<AssetInfo> assetData = GetAssetReportDataList(CompanyId, UnitId, DepartmentId, CustodianId, SubCustodianId, AssetItemId);

                string _ReportName = "AssetReportByCustodian.rdlc";
                DataTable tbl = new ReportDS.AssetReportByCustodianDataTable();
                if (assetData != null)
                {
                    tbl = ReportHelper.GetDataTableFromList(assetData, tbl);
                }
                string reportPath = string.Format("~/Content/Reports/Assets/{0}", _ReportName);

                return ExportData(Server.MapPath(reportPath), strExelFileName, tbl, FileType, ContentType);
            }
            catch (Exception ex)
            {

                return Json(new { Status = "Error" }, JsonRequestBehavior.AllowGet);
            }
        }


        private ActionResult ExportData(string ReportPath, string FileName, DataTable ReportData, string FileType = "Excel", string ContentType = "application/vnd.ms-excel")
        {
            LocalReport localReport = new LocalReport();
            localReport.ReportPath = ReportPath;

            ReportDataSource reportDataSource = new ReportDataSource();
            reportDataSource.Name = "ReportDB";
            reportDataSource.Value = ReportData;
            localReport.DataSources.Add(reportDataSource);


            string reportType = FileType;
            string mimeType;
            string encoding;
            string fileNameExtension;
            //The DeviceInfo settings should be changed based on the reportType            
            //http://msdn2.microsoft.com/en-us/library/ms155397.aspx            
            string deviceInfo = "<DeviceInfo>" +
                "  <OutputFormat>" + FileType + "</OutputFormat>" +
                //"  <PageWidth>36in</PageWidth>" +
                //"  <PageHeight>11in</PageHeight>" +
                //"  <MarginTop>0.5in</MarginTop>" +
                //"  <MarginLeft>0.5in</MarginLeft>" +
                //"  <MarginRight>0.5in</MarginRight>" +
                //"  <MarginBottom>0.5in</MarginBottom>" +
                "</DeviceInfo>";

            Warning[] warnings;
            string[] streams;
            byte[] renderedBytes = localReport.Render(reportType, deviceInfo, out mimeType, out encoding, out fileNameExtension, out streams, out warnings);
            return File(renderedBytes, ContentType, string.Format("{0}.{1}", FileName, fileNameExtension));
        }
    }
}