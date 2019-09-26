using Read_Excel_Webapi.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace Read_Excel_Webapi.Controllers
{
    public class HomeController : ApiController
    {
        public string Get()
        {
            return "Welcome To Web API";
        }
        public List<string> Get(int Id)
        {
            return new List<string> {
                "Data1",
                "Data2"
            };
        }
       
        public List<EmployeeModel> GetEmpDetailsFromExcel()
        {
            List<EmployeeModel> employees = new List<EmployeeModel>();
            string filePath = string.Empty;
            string path = ConfigurationManager.AppSettings["filepath"].ToString();

            string filename=Path.GetFileName(path);
            string extension = Path.GetExtension(path);

            string conString = string.Empty;
            switch (extension)
            {
                case ".xls": //Excel 97-03.
                    conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                    break;
                case ".xlsx": //Excel 07 and above.
                    conString = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                    break;
            }

            conString = string.Format(conString, path);

            using (OleDbConnection connExcel = new OleDbConnection(conString))
            {
                using (OleDbCommand cmdExcel = new OleDbCommand())
                {
                    using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                    {
                        DataTable dt = new DataTable();
                        cmdExcel.Connection = connExcel;

                        //Get the name of First Sheet.
                        connExcel.Open();
                        DataTable dtExcelSchema;
                        dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                        connExcel.Close();

                        //Read Data from First Sheet.
                        connExcel.Open();
                        cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                        odaExcel.SelectCommand = cmdExcel;
                        odaExcel.Fill(dt);
                        connExcel.Close();

                        foreach (DataRow row in dt.Rows)
                        {
                            employees.Add(new EmployeeModel
                            {
                                EmployeeId = Convert.ToInt32(row["Id"]),
                                Name = row["Name"].ToString(),
                                Address = row["Address"].ToString(),
                                EmailId = row["EmailId"].ToString()
                            });
                        }
                    }
                }
            }


         
            return employees;
        }

        public List<EmployeeModel> GetEmpDetailsFromExcel(int id)
        {
            List<EmployeeModel> employees = new List<EmployeeModel>();
            string filePath = string.Empty;
            string path = ConfigurationManager.AppSettings["filepath"].ToString();

            string filename = Path.GetFileName(path);
            string extension = Path.GetExtension(path);

            string conString = string.Empty;
            switch (extension)
            {
                case ".xls": //Excel 97-03.
                    conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                    break;
                case ".xlsx": //Excel 07 and above.
                    conString = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                    break;
            }

            conString = string.Format(conString, path);

            using (OleDbConnection connExcel = new OleDbConnection(conString))
            {
                using (OleDbCommand cmdExcel = new OleDbCommand())
                {
                    using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                    {
                        DataTable dt = new DataTable();
                        cmdExcel.Connection = connExcel;

                        //Get the name of First Sheet.
                        connExcel.Open();
                        DataTable dtExcelSchema;
                        dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                        connExcel.Close();

                        //Read Data from First Sheet.
                        connExcel.Open();
                        cmdExcel.CommandText = "SELECT * From [" + sheetName + "] where Id="+id;
                        odaExcel.SelectCommand = cmdExcel;
                        odaExcel.Fill(dt);
                        connExcel.Close();

                        foreach (DataRow row in dt.Rows)
                        {
                            employees.Add(new EmployeeModel
                            {
                                EmployeeId = Convert.ToInt32(row["Id"]),
                                Name = row["Name"].ToString(),
                                Address = row["Address"].ToString(),
                                EmailId = row["EmailId"].ToString()
                            });
                        }
                    }
                }
            }



            return employees;
        }
    }
}
