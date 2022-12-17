using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace GMTSWebReport.Models
{
    public class SQLHelper
    {


        public static object ExecuteProcedureAsScaler(string SpName, Hashtable SPParameterHashTable, SqlConnection connection, SqlTransaction transaction)
        {

            object ReturnScaler = null;
            using (var command = new SqlCommand(SpName, connection, transaction))
            {
                command.CommandType = CommandType.StoredProcedure;
                foreach (DictionaryEntry val in SPParameterHashTable)
                {
                    command.Parameters.AddWithValue((string)val.Key, val.Value);
                }

                var rval = command.ExecuteScalar();

                //ReturnScaler = command.ExecuteScalar();
            }
            return ReturnScaler;
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
                    da.SelectCommand.CommandTimeout = 180;
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

        public static string GetConnectionString()
        {
            return ConfigurationManager.ConnectionStrings["AmanGMTSCon"].ConnectionString;
        }

      

        public static bool ExecuteCommand(string queryString)
        {
            try
            {
                string connectionString = SQLHelper.GetConnectionString();
                using (SqlConnection connection = new SqlConnection(
                           connectionString))
                {
                    SqlCommand command = new SqlCommand(queryString, connection);
                    command.Connection.Open();
                    command.ExecuteNonQuery();
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }


        public static DataTable GetDataTableFromSQL(string queryString)
        {
            try
            {
                DataTable tbl = new DataTable();
                string connectionString = SQLHelper.GetConnectionString();

                SqlDataAdapter adapter = new SqlDataAdapter(queryString, connectionString);
                adapter.Fill(tbl);
                return tbl;
            }
            catch (Exception ex)
            {
                throw new Exception();
            }
        }

        public static DataSet ExecuteProcedureAsFromDataAdapter(string SpName, Hashtable SPParameterHashTable)
        {

            DataSet dset = new DataSet();
            string connectionString = GetConnectionString();
            SqlConnection connection = new SqlConnection(connectionString);
            using (var command = new SqlCommand(SpName, connection))
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


    }
}