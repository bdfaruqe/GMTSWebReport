using Microsoft.Reporting.WinForms;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Web;

namespace GMTSWebReport.Models
{
    public class ReportHelper
    {
        public static byte[] GetReportToBytes(string ReportPath, DataTable ReportData, string FileType = "Excel", string ContentType = "application/vnd.ms-excel")
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
                "  <PageWidth>16.5in</PageWidth>" +
                "  <PageHeight>11in</PageHeight>" +
                "  <MarginTop>0.5in</MarginTop>" +
                "  <MarginLeft>1in</MarginLeft>" +
                "  <MarginRight>1in</MarginRight>" +
                "  <MarginBottom>0.5in</MarginBottom>" +
                "</DeviceInfo>";

            Warning[] warnings;
            string[] streams;
            byte[] renderedBytes = localReport.Render(reportType, deviceInfo, out mimeType, out encoding, out fileNameExtension, out streams, out warnings);
            return renderedBytes;
            //return File(renderedBytes, ContentType, string.Format("PoleUpdateInformation.{0}", fileNameExtension));
        }

        public static List<T> ConvertDatatableToList<T>(DataTable dt)
        {
            var columnNames = dt.Columns.Cast<DataColumn>()
                    .Select(c => c.ColumnName)
                    .ToList();
            var properties = typeof(T).GetProperties();
            return dt.AsEnumerable().Select(row =>
            {
                var objT = Activator.CreateInstance<T>();
                foreach (var pro in properties)
                {
                    if (columnNames.Contains(pro.Name))
                    {
                        PropertyInfo pI = objT.GetType().GetProperty(pro.Name);
                        pro.SetValue(objT, row[pro.Name] == DBNull.Value ? null : Convert.ChangeType(row[pro.Name], pI.PropertyType));
                    }
                }
                return objT;
            }).ToList();
        }


        public static DataTable ListToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            //Get all the properties by using reflection   
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names  
                dataTable.Columns.Add(prop.Name);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {

                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }

            return dataTable;
        }


        public static DataTable GetDataTableFromList(IEnumerable<object> data, DataTable tbl)
        {

            foreach (var itm in data)
            {
                if (itm != null)
                {
                    var tType = itm.GetType();
                    var tpro = tType.GetProperties();

                    var dr = tbl.NewRow();
                    foreach (DataColumn col in tbl.Columns)
                    {
                        var cname = col.ColumnName;

                        var tp = tpro.Where(p => p.Name == cname).FirstOrDefault();
                        if (tp != null)
                        {
                            var value = tp.GetValue(itm);
                            if (value != null)
                            {
                                dr[cname] = value;
                            }
                        }

                    }
                    tbl.Rows.Add(dr);

                }


            }

            return tbl;

        }

        public static DataTable GetDataTableFromList(object data, DataTable tbl)
        {

            //  foreach (var itm in data)
            //   {
            var tType = data.GetType();
            var tpro = tType.GetProperties();

            var dr = tbl.NewRow();
            foreach (DataColumn col in tbl.Columns)
            {
                var cname = col.ColumnName;

                var tp = tpro.Where(p => p.Name == cname).FirstOrDefault();
                if (tp != null)
                {
                    var value = tp.GetValue(data);
                    if (value != null)
                    {
                        dr[cname] = value;
                    }
                }

            }
            tbl.Rows.Add(dr);
            // }

            return tbl;

        }

        public DateTime? GetLastDateOfMonth(int year, int MonthNo)
        {
            DateTime? returnDate = null;
            int dayno = DateTime.DaysInMonth(year, MonthNo);
            if (dayno > 0)
            {
                returnDate = Convert.ToDateTime(string.Format("{0}-{1}-{2}", year, MonthNo, dayno));
            }

            return returnDate;
        }

        public static string GetConnectionString()
        {
            return ConfigurationManager.ConnectionStrings["AmanGMTSCon"].ConnectionString;
        }
        public static DataTable GetDataTableFromSQL(string queryString)
        {
            try
            {
                DataTable tbl = new DataTable();
                string connectionString = ReportHelper.GetConnectionString();

                SqlDataAdapter adapter = new SqlDataAdapter(queryString, connectionString);
                adapter.Fill(tbl);
                return tbl;
            }
            catch (Exception ex)
            {
                throw new Exception();
            }
        }


        public static string returnstring(int ReportID, int BUID, DateTime? date)
        {
            try
            {
                string connectionString = ReportHelper.GetConnectionString();
                SqlConnection conn = new SqlConnection(connectionString);

                SqlCommand cmd = new SqlCommand("SELECT dbo.FnReportURL(@ReportID,@BUID,@EffectiveDate)", conn);

                // cmd.CommandType=CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@ReportID", ReportID);
                cmd.Parameters.AddWithValue("@BUID", BUID);
                cmd.Parameters.AddWithValue("@EffectiveDate", date);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                string str = dt.Rows[0][0].ToString();
                return str;


            }
            catch (Exception ex)
            {
                throw new Exception();
            }
        }





        public static DataSet ExecuteProcedureAsFromDataAdapter(string SpName, Hashtable SPParameterHashTable, SqlConnection connection, SqlTransaction transaction)
        {

            DataSet dset = new DataSet();
            using (var command = new SqlCommand(SpName, connection, transaction))
            {
                command.CommandType = CommandType.StoredProcedure;
                foreach (DictionaryEntry val in SPParameterHashTable)
                {
                    command.Parameters.AddWithValue((string)val.Key, val.Value);
                }
                using (SqlDataAdapter da = new SqlDataAdapter(command))
                {
                    da.Fill(dset);
                }

            }
            return dset;
        }
        public static DataTable GetExcelData(string _FilePath, string ExcelSheetName)
        {
            DataTable tbl = new DataTable();

            //"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + PrmPathExcelFile + @";Extended Properties=""Excel 8.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text"""
            //string excelConnString = @"provider=Microsoft.ACE.OLEDB.12.0;data source=" + _FilePath + ";extended properties=" + "\"excel 12.0;hdr=yes;IMEX=1\"";

            //string excelConnString = @"provider=Microsoft.ACE.OLEDB.12.0;data source=" + _FilePath + ";extended properties=" + "\"excel 12.0 ;hdr=yes;IMEX=1\"";
            //string excelConnString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _FilePath + @";Extended Properties=""Excel 8.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text""";
            string excelConnString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _FilePath + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";";


            //Create Connection to Excel work book 
            using (OleDbConnection excelConnection = new OleDbConnection(excelConnString))
            {
                ///----------------------------
                string sql = "Select * " + " from [" + ExcelSheetName + "$]";
                OleDbDataAdapter da = new OleDbDataAdapter(sql, excelConnection);
                da.Fill(tbl);

                ///Remove Temp File
                System.IO.File.Delete(_FilePath);
            }






            return tbl;
        }
    }
}