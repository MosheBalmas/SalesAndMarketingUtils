using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Configuration;
using System.Web.Services;
using System.Reflection;

namespace SalesAndMarketingUtilsServices
{
    /// <summary>
    /// Summary description for OpsInvoicesServices
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class SalesAndMarketingUtilsServices : System.Web.Services.WebService
    {
  
        

        public string CreateDemandCoverageFile()
        {
            string userName = System.Web.HttpContext.Current.User.Identity.Name;

            if (String.IsNullOrEmpty(userName))
            {
                userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            }

            string spreadsheetPath = Path.Combine(@"c:\temp\", string.Concat("Demand_Coverage_", DateTime.Now.ToString("yyyyMMdd_HHmmss"), ".xlsx"));
            string fileName = string.Concat("Demand_Coverage_", DateTime.Now.ToString("yyyyMMdd_HHmmss"), ".xlsx");
            File.Delete(spreadsheetPath);
            FileInfo spreadsheetInfo = new FileInfo(spreadsheetPath);
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;


            using (ExcelPackage pck = new ExcelPackage(spreadsheetInfo))
            {

                try
                {
                    var planDetails = CreateNewDemandCoverageData(userName);

                    var JDPlanWorksheet = pck.Workbook.Worksheets.Add("Demand_Coverage");

                    if (planDetails.Item1 == -1)
                    {
                        throw new Exception("Could not generate file data");

                    }

                    var headerDt = CreateDemandCoverageTopHeaderDataTable();
                    JDPlanWorksheet.Cells["A2"].LoadFromDataTable(headerDt, false);


                    JDPlanWorksheet.Cells["A1"].Value = "Meeting ID";
                    JDPlanWorksheet.Cells["B1"].Value = planDetails.Item1;


                    JDPlanWorksheet.Cells["E2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    JDPlanWorksheet.Cells["F2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    JDPlanWorksheet.Cells["G2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    //Borders
                    SetBorders(JDPlanWorksheet, "E2");
                    SetBorders(JDPlanWorksheet, "F2");
                    SetBorders(JDPlanWorksheet, "G2");
                    SetBorders(JDPlanWorksheet, "H2");


                    //Background
                    SetBackgroud(JDPlanWorksheet, "E2", "#EBEDEF");
                    SetBackgroud(JDPlanWorksheet, "F2", "#FDC168");
                    SetBackgroud(JDPlanWorksheet, "G2", "#D3E788");
                    SetBackgroud(JDPlanWorksheet, "H2", "#EBEDEF");


                    JDPlanWorksheet.Cells["A2:H2"].Style.Font.Bold = true;

                    //Secondary header
                    headerDt = CreateDemandCoverageSecondaryHeaderDataTable(planDetails.Item2);
                    JDPlanWorksheet.Cells["A3"].LoadFromDataTable(headerDt, false);

                    //Align 
                    JDPlanWorksheet.Cells["E3:G3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    JDPlanWorksheet.Cells["A3:D3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    JDPlanWorksheet.Cells["H3:H3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    //Borders
                    SetBorders(JDPlanWorksheet, "A3:H3");

                    //Background
                    SetBackgroud(JDPlanWorksheet, "E3", "#EBEDEF");
                    SetBackgroud(JDPlanWorksheet, "F3", "#FDC168");
                    SetBackgroud(JDPlanWorksheet, "G3", "#D3E788");
                    SetBackgroud(JDPlanWorksheet, "A3:D3", "#EBEDEF");
                    SetBackgroud(JDPlanWorksheet, "H3", "#EBEDEF");


                    JDPlanWorksheet.Cells["A3:G3"].Style.Font.Bold = true;


                    // populate spreadsheet with data

                    var ds = ReadDemandCoverageDataFromDB(planDetails.Item1);
                    var dt = ds.Tables[0];
                    var tabRowcount = dt.Rows.Count;

                    JDPlanWorksheet.Cells["A4"].LoadFromDataTable(dt, false);

                    JDPlanWorksheet.Cells[JDPlanWorksheet.Dimension.Address].AutoFitColumns();

                    //Borders
                    SetBorders(JDPlanWorksheet, "A4:H" + (3 + tabRowcount).ToString());


                    //Align 
                    JDPlanWorksheet.Cells["A4:D" + (3 + tabRowcount).ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    JDPlanWorksheet.Cells["E4:G" + (3 + tabRowcount).ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    JDPlanWorksheet.Cells["H4:H" + (3 + tabRowcount).ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    //Background
                    SetBackgroud(JDPlanWorksheet, "G4" + (3 + tabRowcount).ToString(), "#F0F2F5");
                    SetBackgroud(JDPlanWorksheet, "A4:F" + (3 + tabRowcount).ToString(), "#FCFCFC");
                    SetBackgroud(JDPlanWorksheet, "H4:H" + (3 + tabRowcount).ToString(), "#FCFCFC");

                    JDPlanWorksheet.Protection.IsProtected = true; //--------Protect whole sheet
                    JDPlanWorksheet.Cells["G4:H" + (3 + tabRowcount).ToString()].Style.Locked = false; //-------Unlock 3rd column

                    JDPlanWorksheet.Column(8).Width = 100;
                    JDPlanWorksheet.Column(8).Style.WrapText = true;

                    //Add a List validation to the C column
                    var val3 = JDPlanWorksheet.DataValidations.AddIntegerValidation("G4:G" + (3 + tabRowcount).ToString());
                    //For Integer Validation, you have to set error message to true
                    val3.ShowErrorMessage = true;
                    val3.Error = "The value must be a positive integer";
                    //Minimum allowed Value
                    val3.Formula.Value = 0;
                    //Maximum allowed Value
                    val3.Formula2.Value = 1000000;
                    //If the cells are not filled, allow blanks or fill with a valid value, 
                    //otherwise it could generate a error when saving 
                    val3.AllowBlank = true;

                    byte[] bin = pck.GetAsByteArray();


                    HttpContext.Current.Response.Clear();

                    HttpContext.Current.Response.AppendHeader("Content-Length", bin.Length.ToString());
                    HttpContext.Current.Response.AppendHeader("Content-Disposition", String.Format("attachment; filename={0}", fileName)
                        );
                    HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                    HttpContext.Current.Response.BinaryWrite(bin);
                    HttpContext.Current.Response.End();

                    //pck.Save();

                    return "Completed successfully";
                }
                catch (Exception e)
                {

                    return string.Format(@"Failed to generate XLSX file : {0}", e.Message);
                }
            }
        }


        [WebMethod]
        public string CreateJDBaseSpreadsheet(int planId)
        {
            string userName = System.Web.HttpContext.Current.User.Identity.Name;
            
            if (String.IsNullOrEmpty(userName))
            {
                userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            }

            string spreadsheetPath = Path.Combine(@"c:\temp\", string.Concat("JD_Plan_", DateTime.Now.ToString("yyyyMMdd_HHmmss"), ".xlsx"));
            string fileName =  string.Concat("JD_Plan_", DateTime.Now.ToString("yyyyMMdd_HHmmss"), ".xlsx");
            File.Delete(spreadsheetPath);
            FileInfo spreadsheetInfo = new FileInfo(spreadsheetPath);
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;


            using (ExcelPackage pck = new ExcelPackage(spreadsheetInfo))
            {

                try
                {
                    var planDetails = new Tuple<int, int, string>(planId, -1, null);
                    if (planId == -1)
                    {
                        planDetails = CreateNewJDBasePlan(userName);
                    }
                    else
                    {
                        planDetails = GetJDBasePlanHeader(planId);
                    }

                    var JDPlanWorksheet = pck.Workbook.Worksheets.Add("JD_Plan");

                    if (planDetails.Item1 == -1)
                    {
                        throw new Exception("Selected plan does not exists");

                    }

                    var headerDt = CreateJDBaseTopHeaderDataTable();
                    JDPlanWorksheet.Cells["A2"].LoadFromDataTable(headerDt, false);


                    JDPlanWorksheet.Cells["A1"].Value = "Plan ID";
                    JDPlanWorksheet.Cells["B1"].Value = planDetails.Item1;

                    JDPlanWorksheet.Cells["D1"].Value = "Plan Name";

                    JDPlanWorksheet.Cells["E1:X1"].Merge = true; //plan name description
                    
                    JDPlanWorksheet.Cells["D2:H2"].Merge = true;
                    JDPlanWorksheet.Cells["I2:N2"].Merge = true;
                    JDPlanWorksheet.Cells["O2:S2"].Merge = true;

                    JDPlanWorksheet.Cells["E1:X1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    JDPlanWorksheet.Cells["D2:H2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    JDPlanWorksheet.Cells["I2:N2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    JDPlanWorksheet.Cells["O2:S2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    //Borders
                    SetBorders(JDPlanWorksheet, "E1:X1");
                    SetBorders(JDPlanWorksheet, "D2:H2");
                    SetBorders(JDPlanWorksheet, "I2:N2");
                    SetBorders(JDPlanWorksheet, "O2:S2");
                    SetBorders(JDPlanWorksheet, "T2:X2");

                    //Background
                    SetBackgroud(JDPlanWorksheet, "E1:X1", "#D4E6F1");

                    SetBackgroud(JDPlanWorksheet, "D2:H2", "#EAF2F8");
                    SetBackgroud(JDPlanWorksheet, "I2:N2", "#D4E6F1");
                    SetBackgroud(JDPlanWorksheet, "O2:S2", "#A9CCE3");
                    SetBackgroud(JDPlanWorksheet, "T2:X2", "#EAF2F8");

                    JDPlanWorksheet.Cells["A1"].Style.Font.Bold = true;
                    JDPlanWorksheet.Cells["D1"].Style.Font.Bold = true;
                    JDPlanWorksheet.Cells["E1:X1"].Style.Font.Bold = true;
                    JDPlanWorksheet.Cells["A2:X2"].Style.Font.Bold = true;

                    //Secondary header
                    headerDt = CreateJDBaseSecondaryHeaderDataTable(planDetails.Item2);
                    JDPlanWorksheet.Cells["A3"].LoadFromDataTable(headerDt, false);

                    //Align 
                    JDPlanWorksheet.Cells["D3:X3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    JDPlanWorksheet.Cells["A3:C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    //Borders
                    SetBorders(JDPlanWorksheet, "A3:X3");

                    //Background
                    SetBackgroud(JDPlanWorksheet, "A3:X3", "#EAF2F8");
 

                    JDPlanWorksheet.Cells["A3:X3"].Style.Font.Bold = true;

                    //Third header
                    var PeriodsDs = ReadPlanPeriodsFromDB(planDetails.Item1);
                    var PeriodsDt = PeriodsDs.Tables[0];


                    JDPlanWorksheet.Cells["A4"].LoadFromDataTable(PeriodsDt, false);
                    
                    
                    //Align 
                    JDPlanWorksheet.Cells["D4:X4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    JDPlanWorksheet.Cells["A4:C4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    //Borders
                    SetBorders(JDPlanWorksheet, "A4:X4");

                    //Background
                    SetBackgroud(JDPlanWorksheet, "A4:X4", "#EAF2F8");


                    JDPlanWorksheet.Cells["A4:X4"].Style.Font.Bold = true;



                    // populate spreadsheet with data

                    var ds = ReadJDBasePlanFromDB(planDetails.Item1);
                    //set plan name 
                    
                    JDPlanWorksheet.Cells["E1:X1"].Value = ds.Item1;
                    
                    var dt = ds.Item2.Tables[0];
                    var tabRowcount = dt.Rows.Count;

                    JDPlanWorksheet.Cells["A5"].LoadFromDataTable(dt, false);

                    JDPlanWorksheet.Cells[JDPlanWorksheet.Dimension.Address].AutoFitColumns();

                    //Borders
                    SetBorders(JDPlanWorksheet, "A5:X" + (4 + tabRowcount).ToString());


                    //Align 
                    JDPlanWorksheet.Cells["A5:C" + (4 + tabRowcount).ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    JDPlanWorksheet.Cells["D5:X" + (4 + tabRowcount).ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    //Background
                    SetBackgroud(JDPlanWorksheet, "I5:L" + (4 + tabRowcount).ToString(), "#E9F7EF");
                    SetBackgroud(JDPlanWorksheet, "M5:M" + (4 + tabRowcount).ToString(), "#FCFCFC");
                    SetBackgroud(JDPlanWorksheet, "N5:N" + (4 + tabRowcount).ToString(), "#E9F7EF");
                    SetBackgroud(JDPlanWorksheet, "O5:S" + (4 + tabRowcount).ToString(), "#FCFCFC");
                    SetBackgroud(JDPlanWorksheet, "T5:U" + (4 + tabRowcount).ToString(), "#E9F7EF");
                    SetBackgroud(JDPlanWorksheet, "V5:W" + (4 + tabRowcount).ToString(), "#FCFCFC");
                    SetBackgroud(JDPlanWorksheet, "X5:X" + (4 + tabRowcount).ToString(), "#E9F7EF");
                    SetBackgroud(JDPlanWorksheet, "A5:H" + (4 + tabRowcount).ToString(), "#FCFCFC");

                    JDPlanWorksheet.Protection.IsProtected = true; //--------Protect whole sheet
                    JDPlanWorksheet.Cells["I5:L" + (4 + tabRowcount).ToString()].Style.Locked = false; //-------Unlock JD
                    JDPlanWorksheet.Cells["N5:N" + (4 + tabRowcount).ToString()].Style.Locked = false; //-------20 weeks
                    JDPlanWorksheet.Cells["T5:U" + (4 + tabRowcount).ToString()].Style.Locked = false; //-------Ave/ RR
                    JDPlanWorksheet.Cells["X5:X" + (4 + tabRowcount).ToString()].Style.Locked = false; //-------URequested

                    //Unlock columns to the right of the data
                    JDPlanWorksheet.Cells["Y5:AY" + (4 + tabRowcount).ToString()].Style.Locked = false; //-------URequested

                    JDPlanWorksheet.Cells["E1:X1"].Style.Locked = false; //unloack plan name


                    //Add a List validation to the C column
                    var val3 = JDPlanWorksheet.DataValidations.AddIntegerValidation("I5:L" + (4 + tabRowcount).ToString());
                    //For Integer Validation, you have to set error message to true
                    val3.ShowErrorMessage = true;
                    val3.Error = "The value must be a positive integer";
                    //Minimum allowed Value
                    val3.Formula.Value = 0;
                    //Maximum allowed Value
                    val3.Formula2.Value = 1000000;
                    //If the cells are not filled, allow blanks or fill with a valid value, 
                    //otherwise it could generate a error when saving 
                    val3.AllowBlank = true;

                    //Add a List validation to the C column
                    val3 = JDPlanWorksheet.DataValidations.AddIntegerValidation("M5:M" + (4 + tabRowcount).ToString());
                    //For Integer Validation, you have to set error message to true
                    val3.ShowErrorMessage = true;
                    val3.Error = "The value must be a positive integer";
                    //Minimum allowed Value
                    val3.Formula.Value = 0;
                    //Maximum allowed Value
                    val3.Formula2.Value = 1000000;
                    //If the cells are not filled, allow blanks or fill with a valid value, 
                    //otherwise it could generate a error when saving 
                    val3.AllowBlank = true;

                    //Add a List validation to the C column
                    val3 = JDPlanWorksheet.DataValidations.AddIntegerValidation("T5:U" + (4 + tabRowcount).ToString());
                    //For Integer Validation, you have to set error message to true
                    val3.ShowErrorMessage = true;
                    val3.Error = "The value must be a positive integer";
                    //Minimum allowed Value
                    val3.Formula.Value = 0;
                    //Maximum allowed Value
                    val3.Formula2.Value = 1000000;
                    //If the cells are not filled, allow blanks or fill with a valid value, 
                    //otherwise it could generate a error when saving 
                    val3.AllowBlank = true;

                    val3 = JDPlanWorksheet.DataValidations.AddIntegerValidation("X5:X" + (4 + tabRowcount).ToString());
                    //For Integer Validation, you have to set error message to true
                    val3.ShowErrorMessage = true;
                    val3.Error = "The value must be a positive integer";
                    //Minimum allowed Value
                    val3.Formula.Value = 0;
                    //Maximum allowed Value
                    val3.Formula2.Value = 1000000;
                    //If the cells are not filled, allow blanks or fill with a valid value, 
                    //otherwise it could generate a error when saving 
                    val3.AllowBlank = true;


                    byte[] bin = pck.GetAsByteArray();


                    HttpContext.Current.Response.Clear();

                    HttpContext.Current.Response.AppendHeader("Content-Length", bin.Length.ToString());
                    HttpContext.Current.Response.AppendHeader("Content-Disposition", String.Format("attachment; filename={0}", fileName)
                        );
                    HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                    HttpContext.Current.Response.BinaryWrite(bin);
                    HttpContext.Current.Response.End();

                    //pck.Save();

                    return "Completed successfully" ;
                }
                catch (Exception e)
                {

                    return string.Format(@"Failed to generate XLSX file : {0}",e.Message);
                }
            }
        }
        [WebMethod]
        public string UpdateJDPlanStatus(int planId, string status)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString))
                {

                    using (SqlCommand cmd = new SqlCommand("[JD].[UpdatePlanStatus_SP]", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@PlanId", planId));
                        cmd.Parameters.Add(new SqlParameter("@Status", status));

                        cmd.Parameters.Add(new SqlParameter("@User", System.Web.HttpContext.Current.User.Identity.Name));



                        if (con.State != ConnectionState.Open)
                        {
                            con.Open();
                        }
                        var affected = cmd.ExecuteNonQuery();



                        return "Completed Successfully";

                    }


                }
            }
            catch (Exception e)
            {
                return  e.Message;

            }



        }



        public string ReadJDSpreadsheet()
        {

            //var ul = new Upload();
            //ul.UploadFile();
            var ul = new FileUploadingController();
            var res = ul.UploadFile();


            return "";
        }


        private Tuple<int, int, string> GetJDBasePlanHeader(int planId)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString))
                {
                    DataSet ds = new DataSet("PlanHeader");
                    using (SqlCommand cmd = new SqlCommand("[JD].[GenPlanHeader_SP]", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("PlanId", SqlDbType.VarChar, 100).Value = planId;
                        //cmd.Parameters.Add("@PlanId", SqlDbType.Int).Direction = ParameterDirection.Output;
                        //cmd.Parameters.Add("@CurrentWW", SqlDbType.Int).Direction = ParameterDirection.Output;


                        if (con.State != ConnectionState.Open)
                        {
                            con.Open();
                        }
                        //cmd.ExecuteNonQuery();

                        SqlDataAdapter da = new SqlDataAdapter();
                        da.SelectCommand = cmd;

                        da.Fill(ds);

                        var p = (Convert.ToInt32(ds.Tables[0].Rows[0][0]));
                        var ww = (Convert.ToInt32(ds.Tables[0].Rows[0][1]));


                        return new Tuple<int, int, string>((Convert.ToInt32(ds.Tables[0].Rows[0][0]))
                                                        , (Convert.ToInt32(ds.Tables[0].Rows[0][1]))
                                                        , "Completed Successfully");

                    }


                }
            }
            catch (Exception e)
            {
                return new Tuple<int, int, string>(-1, -1, e.Message);

            }



        }

        private Tuple<int, int, string> CreateNewJDBasePlan(string userName)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString))
                {
                    DataSet ds = new DataSet("PlanHeader");
                    using (SqlCommand cmd = new SqlCommand("[JD].[GenerateNewPlan_V2_SP]", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("RequestingUser", SqlDbType.VarChar, 100).Value = userName;
                        //cmd.Parameters.Add("@PlanId", SqlDbType.Int).Direction = ParameterDirection.Output;
                        //cmd.Parameters.Add("@CurrentWW", SqlDbType.Int).Direction = ParameterDirection.Output;


                        if (con.State != ConnectionState.Open)
                        {
                            con.Open();
                        }
                        //cmd.ExecuteNonQuery();
                        SqlDataAdapter da = new SqlDataAdapter();
                        da.SelectCommand = cmd;

                        da.Fill(ds);

                        var p = (Convert.ToInt32(ds.Tables[0].Rows[0][0]));
                        var ww = (Convert.ToInt32(ds.Tables[0].Rows[0][1]));


                        return new Tuple<int, int, string>((Convert.ToInt32(ds.Tables[0].Rows[0][0]))
                                                        , (Convert.ToInt32(ds.Tables[0].Rows[0][1]))
                                                        , "Completed Successfully");



                    }


                }
            }
            catch (Exception e)
            {
                return new Tuple<int, int, string>(-1, -1, e.Message);

            }



        }

        private Tuple<string,DataSet> ReadJDBasePlanFromDB(int planId)
        {
            DataSet JDBaseRows = new DataSet("JDBase");
            string planName="";
            using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString))
            {
                con.Open();
                SqlCommand command = new SqlCommand(string.Format(@"Select planName From JD.JD_Plans where planId = {0}", planId),con);


                using (SqlDataReader reader = command.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        var rd = reader["planName"];
                        planName = (rd == DBNull.Value) ? string.Empty : rd.ToString();
            
                    }
                }


                SqlDataAdapter JDBase = new SqlDataAdapter(string.Format(@"Select      

                                                                          SKU
                                                                        , Product_Name
                                                                        , assembly_level
                                                                        , Forecast_Q1
                                                                        , Forecast_Q2
                                                                        , Forecast_Q3
                                                                        , Forecast_Q4

                                                                        , Forecast_Weeks20

                                                                        , JD_Q1
                                                                        , JD_Q2
                                                                        , JD_Q3
                                                                        , JD_Q4
                                                                        , JD_Weeks20_Calculated
                                                                        , JD_Weeks20

                                                                        , BnB_Q1
                                                                        , BnB_Q2
                                                                        , BnB_Q3
                                                                        , BnB_Q4
                                                                        , Backlog_Weeks20
                                                                        , Avg_RR
                                                                        , Avg_RR_3Weeks
                                                                        , InvQty
                                                                        , WareHouseGoal
                                                                        , Requested

                          
                                                                        
                                                                
                                                              From [JD].[JD_Plan_Pivoted_4_XLS_V2]
                                                              where planId = {0}
                                                              order by assembly_level, Product_Name ", planId), con);

                JDBase.FillSchema(JDBaseRows, SchemaType.Source, "JDBaseRows");
                JDBase.Fill(JDBaseRows, "JDBaseRows");

                return new Tuple<string, DataSet> (planName, JDBaseRows);

            }
        }

        private DataSet ReadPlanPeriodsFromDB(int planId)
        {
            DataSet JDBaseRows = new DataSet("JDBase");
            using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString))
            {
                con.Open();
              

                SqlDataAdapter JDBase = new SqlDataAdapter(string.Format(@"Select      

                                                                          null
                                                                        , null
                                                                        , null
                                                                        , Q1
                                                                        , Q2
                                                                        , Q3
                                                                        , Q4

                                                                        , null

                                                                        , Q1
                                                                        , Q2
                                                                        , Q3
                                                                        , Q4
                                                                        , null
                                                                        , null

                                                                        , Q1
                                                                        , Q2
                                                                        , Q3
                                                                        , Q4
                                                                        , null
                                                                        , null
                                                                        , null
                                                                        , null
                                                                        , null
                                                                        , null

                          
                                                                        
                                                                
                                                              From JD.JD_Plan_Periods_V2 

                                                              
                                                              where planId = {0} ", planId), con);

                JDBase.FillSchema(JDBaseRows, SchemaType.Source, "JDBaseRows");
                JDBase.Fill(JDBaseRows, "JDBaseRows");

                return   JDBaseRows;

            }
        }


        private DataTable CreateJDBaseTopHeaderDataTable()
        {
            DataTable dataTable = new DataTable();

            //add three colums to the datatable
            dataTable.Columns.Add("Sku", typeof(string));
            dataTable.Columns.Add("Product_Name", typeof(string));
            dataTable.Columns.Add("assembly_level", typeof(string));
            dataTable.Columns.Add("Forecast1", typeof(string));
            dataTable.Columns.Add("Forecast2", typeof(string));
            dataTable.Columns.Add("Forecast3", typeof(string));
            dataTable.Columns.Add("Forecast4", typeof(string));
            dataTable.Columns.Add("Forecast5", typeof(string));
            dataTable.Columns.Add("JD1", typeof(string));
            dataTable.Columns.Add("JD2", typeof(string));
            dataTable.Columns.Add("JD3", typeof(string));
            dataTable.Columns.Add("JD4", typeof(string));
            dataTable.Columns.Add("JD5", typeof(string));
            dataTable.Columns.Add("JD6", typeof(string));
            dataTable.Columns.Add("BnB1", typeof(string));
            dataTable.Columns.Add("BnB2", typeof(string));
            dataTable.Columns.Add("BnB3", typeof(string));
            dataTable.Columns.Add("BnB4", typeof(string));
            dataTable.Columns.Add("BnB5", typeof(string));
            dataTable.Columns.Add("Prior QTR RR", typeof(string));
            dataTable.Columns.Add("Current QTR RR", typeof(string));


            dataTable.Columns.Add("BOH", typeof(string));
            dataTable.Columns.Add("Inv Goal", typeof(string));
            dataTable.Columns.Add("Requested", typeof(string));
            //dataTable.Columns.Add("Approved", typeof(string));


            dataTable.Rows.Add( null, null, null
                                , "Forecast", "Forecast", "Forecast", "Forecast", "Forecast"
                                , "JD", "JD", "JD", "JD", "JD", "JD"
                                , "BnB", "BnB", "BnB", "BnB", "BnB"
                                , "Ave RR"
                                , "Ave RR 3Weeks"
                                , "BOH"
                                , "Inv Goal"
                                , "Requested"
                                //, "Approved"
                                );



            return dataTable;


        }

        private DataTable CreateJDBaseSecondaryHeaderDataTable(int ww)
        {
            DataTable dataTable = new DataTable();

            //add three colums to the datatable

            dataTable.Columns.Add("Sku", typeof(string));
            dataTable.Columns.Add("Product_Name", typeof(string));
            dataTable.Columns.Add("assembly_level", typeof(string));
            dataTable.Columns.Add("Forecast_Q1", typeof(string));
            dataTable.Columns.Add("Forecast_Q2", typeof(string));
            dataTable.Columns.Add("Forecast_Q3", typeof(string));
            dataTable.Columns.Add("Forecast_Q4", typeof(string));
            dataTable.Columns.Add("Forecast_20weeks", typeof(string));
            dataTable.Columns.Add("JD_Q1", typeof(string));
            dataTable.Columns.Add("JD_Q2", typeof(string));
            dataTable.Columns.Add("JD_Q3", typeof(string));
            dataTable.Columns.Add("JD_Q4", typeof(string));
            dataTable.Columns.Add("JD_20weeks_Calculated", typeof(string));
            dataTable.Columns.Add("JD_20weeks", typeof(string));
            dataTable.Columns.Add("BnB_Q1", typeof(string));
            dataTable.Columns.Add("BnB_Q2", typeof(string));
            dataTable.Columns.Add("BnB_Q3", typeof(string));
            dataTable.Columns.Add("BnB_Q4", typeof(string));
            dataTable.Columns.Add("BnB_20weeks", typeof(string));
            dataTable.Columns.Add("Ave RR", typeof(string));
            dataTable.Columns.Add("Ave RR 3Weeks", typeof(string));

            dataTable.Columns.Add("BOH", typeof(string));
            dataTable.Columns.Add("Inv Goal", typeof(string));
            dataTable.Columns.Add("Requested", typeof(string));
            //dataTable.Columns.Add("Approved", typeof(string));


            dataTable.Rows.Add( "Sku", "Product_Name", "assembly_level"
                                , "Q1", "Q2", "Q3", "Q4", "20 weeks"
                                , "Q1", "Q2", "Q3", "Q4", "20 weeks Calculated", "20 weeks"
                                , "Q1", "Q2", "Q3", "Q4", "20 weeks"
                                , ""
                                , ""
                                , ww.ToString()
                                , ""
                                , ""
                                );



            return dataTable;


        }

 


        private Tuple<int, int, string> CreateNewDemandCoverageData(string userName)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString))
                {
                    DataSet ds = new DataSet("DemandCoverage");
                    using (SqlCommand cmd = new SqlCommand("[JD].[GenerateNewDemandCoverage_SP]", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("RequestingUser", SqlDbType.VarChar, 100).Value = userName;
                        //cmd.Parameters.Add("@PlanId", SqlDbType.Int).Direction = ParameterDirection.Output;
                        //cmd.Parameters.Add("@CurrentWW", SqlDbType.Int).Direction = ParameterDirection.Output;


                        if (con.State != ConnectionState.Open)
                        {
                            con.Open();
                        }
                        //cmd.ExecuteNonQuery();
                        SqlDataAdapter da = new SqlDataAdapter();
                        da.SelectCommand = cmd;

                        da.Fill(ds);

                        var p = (Convert.ToInt32(ds.Tables[0].Rows[0][0]));
                        var ww = (Convert.ToInt32(ds.Tables[0].Rows[0][1]));


                        return new Tuple<int, int, string>(p
                                                        , ww
                                                        , "Completed Successfully");



                    }


                }
            }
            catch (Exception e)
            {
                return new Tuple<int, int, string>(-1, -1, e.Message);

            }



        }


        private DataSet ReadDemandCoverageDataFromDB(int meetingId)
        {
            DataSet JDBaseRows = new DataSet("DemandCoverage");
            using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString))
            {
                SqlDataAdapter JDBase = new SqlDataAdapter(string.Format(@"Select      
                                                                          Year
                                                                        , SKU
                                                                        , Product_Name
                                                                        , assembly_level
                                                                        , Demand
                                                                        , BPRequested
                                                                        , BPApproved
                                                                        , Remarks
                                                              From JD.JD_Demand_Coverage_V
                                                              where MeetingId = {0}
                                                              order by assembly_level, Product_Name ", meetingId), con);

                JDBase.FillSchema(JDBaseRows, SchemaType.Source, "DemandCoverage");
                JDBase.Fill(JDBaseRows, "DemandCoverage");

                return JDBaseRows;

            }
        }

        private DataTable CreateDemandCoverageTopHeaderDataTable()
        {
            DataTable dataTable = new DataTable();

            //add three colums to the datatable
            dataTable.Columns.Add("Year", typeof(int));
            dataTable.Columns.Add("Sku", typeof(string));
            dataTable.Columns.Add("Product_Name", typeof(string));
            dataTable.Columns.Add("assembly_level", typeof(string));
            dataTable.Columns.Add("Demand", typeof(string));
            dataTable.Columns.Add("BP Requested", typeof(string));
            dataTable.Columns.Add("BP Approved", typeof(string));
            dataTable.Columns.Add("Remarks", typeof(string));



            dataTable.Rows.Add(null, null, null, null
                                , "Demand", "BP Requested", "BP Approved","Remarks");



            return dataTable;


        }

        private DataTable CreateDemandCoverageSecondaryHeaderDataTable(int ww)
        {
            DataTable dataTable = new DataTable();

            //add three colums to the datatable
            dataTable.Columns.Add("Year", typeof(string));
            dataTable.Columns.Add("Sku", typeof(string));
            dataTable.Columns.Add("Product_Name", typeof(string));
            dataTable.Columns.Add("assembly_level", typeof(string));
            dataTable.Columns.Add("Demand", typeof(string));
            dataTable.Columns.Add("BP Requested", typeof(string));
            dataTable.Columns.Add("BP Approved", typeof(string));
            dataTable.Columns.Add("Remarks", typeof(string));

            dataTable.Rows.Add("Year", "Sku", "Product_Name", "assembly_level"
                                , "20 weeks", "20 weeks", "20 weeks","");



            return dataTable;


        }


        //Sheet reader
        public static DataTable GetDataTableFromExcel(string path, int startCol, int endCol , int startRow, int endRow, DataTable tbl)
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets.First();
                
                
                
                for (int rowNum = startRow; rowNum <= endRow; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, startCol, rowNum, endCol];
                    DataRow row = tbl.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
                return tbl;
            }
        }

        //Sheet styling
        private void SetBackgroud(ExcelWorksheet ws, string addr, string color)
        {
            Color colFromHex = System.Drawing.ColorTranslator.FromHtml(color);

            ws.Cells[addr].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[addr].Style.Fill.BackgroundColor.SetColor(colFromHex);

        }

        private void SetBorders(ExcelWorksheet ws, string addr)
        {
            ws.Cells[addr].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells[addr].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells[addr].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws.Cells[addr].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

        }


        //--------------------------------------------------------------------------------------------------------------------
        //SO data ------------------------------------------------------------------------------------------------------------

        [WebMethod]

        public string CreateSODataSpreadsheet()
        {
            string userName = System.Web.HttpContext.Current.User.Identity.Name;

            if (String.IsNullOrEmpty(userName))
            {
                userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            }

            string spreadsheetPath = Path.Combine(@"c:\temp\", string.Concat("SO_Data_", DateTime.Now.ToString("yyyyMMdd_HHmmss"), ".xlsx"));
            string fileName = string.Concat("SO_Data_", DateTime.Now.ToString("yyyyMMdd_HHmmss"), ".xlsx");
            File.Delete(spreadsheetPath);
            FileInfo spreadsheetInfo = new FileInfo(spreadsheetPath);
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;


            using (ExcelPackage pck = new ExcelPackage(spreadsheetInfo))
            {

                try
                {

                    var SODataWorksheet = pck.Workbook.Worksheets.Add("SO_Data");

                    
                    var headerDt = CreateSODataTopHeaderDataTable();
                    SODataWorksheet.Cells["A1"].LoadFromDataTable(headerDt, false);


                    SODataWorksheet.Cells["A1:AD1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                    //Borders
                    SetBorders(SODataWorksheet, "A1:AD1");

                    //Background
                    SetBackgroud(SODataWorksheet, "A1:AD1", "#D4E6F1");


                    SODataWorksheet.Cells["A1:AD1"].Style.Font.Bold = true;


                    // populate spreadsheet with data

                    var ds = ReadSODataFromDB();
                    //set plan name 

                    //SODataWorksheet.Cells["E1:X1"].Value = ds;

                    var dt = ds.Tables[0];
                    var tabRowcount = dt.Rows.Count;

                    SODataWorksheet.Cells["A2"].LoadFromDataTable(dt, false);

                    SODataWorksheet.Cells[SODataWorksheet.Dimension.Address].AutoFitColumns();

                    //Borders
                    SetBorders(SODataWorksheet, "A2:AD" + (1 + tabRowcount).ToString());


                    //Align 
                    SODataWorksheet.Cells["A2:AD" + (1 + tabRowcount).ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    //Background
                    //SetBackgroud(SODataWorksheet, "A2:AD" + (1 + tabRowcount).ToString(), "#E9F7EF");
                    /*
                    SODataWorksheet.Protection.IsProtected = true; //--------Protect whole sheet
                    SODataWorksheet.Cells["A2:AD" + (1 + tabRowcount).ToString()].Style.Locked = false; //-------Unlock JD

                    
                    //Unlock columns to the right of the data
                    SODataWorksheet.Cells["AC2:BY" + (1 + tabRowcount).ToString()].Style.Locked = false; //-------URequested
                    */
                    /*
                    var dd = SODataWorksheet.Cells["Z2:Z" + (1 + tabRowcount).ToString()].DataValidation.AddListDataValidation() ;
                    dd.AllowBlank = true;
                    dd.Formula.Values.Add("Mark Kotzan");
                    dd.Formula.Values.Add("Nicholas Heywood");
                    dd.Formula.Values.Add("Mike McSweeney");
                    dd.Formula.Values.Add("Kenny He");
                    dd.Formula.Values.Add("Takumi Nakamura");
                    dd.Formula.Values.Add("Miro Mlejnek");
                    dd.Formula.Values.Add("Chris Chih");
                    dd.Formula.Values.Add("Brad Suessmith");
                    */

                    ExcelWorksheet ddList = pck.Workbook.Worksheets.Add("Owners");
                    var ownrs = ReadSOOwnersFromDB();
                    ddList.Cells["A1"].LoadFromDataTable(ownrs.Tables[0], false);
                    var ownrsListRowcount = ownrs.Tables[0].Rows.Count;

                    SODataWorksheet.Names.Add("OwnersList", ddList.Cells["A1:A" + ownrsListRowcount.ToString()]);


                    for (int idx = 1; idx < (1 + ownrsListRowcount); idx++)
                    {
                        var val = SODataWorksheet.DataValidations.AddListValidation("Z"+idx.ToString());
                        val.Formula.ExcelFormula = "=OwnersList";
                        val.AllowBlank = true;
                        val.ShowInputMessage = true;
                        val.ShowErrorMessage = true;

                    }


                    
                    ExcelWorksheet acntList = pck.Workbook.Worksheets.Add("Accounts");
                    var accnts = ReadSOAccountsFromDB();
                    acntList.Cells["A1"].LoadFromDataTable(accnts.Tables[0], false);
                    var acntListRowcount = accnts.Tables[0].Rows.Count;
                    /*
                    ReadSOAccountsFromDB
                    acntList.Cells[1, 1].Value = "";
                    acntList.Cells[2, 1].Value = "Mark Kotzan";
                    acntList.Cells[3, 1].Value = "Nicholas Heywood";
                    acntList.Cells[4, 1].Value = "Mike McSweeney";
                    acntList.Cells[5, 1].Value = "Kenny He";
                    acntList.Cells[6, 1].Value = "Takumi Nakamura";
                    acntList.Cells[7, 1].Value = "Miro Mlejnek";
                    acntList.Cells[8, 1].Value = "Chris Chih";
                    acntList.Cells[8, 1].Value = "Brad Suessmith";
                    */
                    SODataWorksheet.Names.Add("AccountsList", acntList.Cells["A1:A"+ acntListRowcount.ToString()]);

                    for (int idx = 1; idx < (1 + acntListRowcount); idx++)
                    {
                        var val = SODataWorksheet.DataValidations.AddListValidation("Y" + idx.ToString());
                        val.Formula.ExcelFormula = "=AccountsList";
                        val.AllowBlank = true;
                        val.ShowInputMessage = true;
                        val.ShowErrorMessage = true;


                    }





                    /*
                    var address = ddList.Cells[1, 1, 3, 1].Address.ToString();
                    var arr = address.Split(':');
                    var char1 = arr[0][0];
                    var num1 = arr[0].Trim(char1);
                    var char2 = arr[1][0];
                    var num2 = arr[1].Trim(char2);


                    SODataWorksheet.Cells["Z2:Z" + (1 + tabRowcount).ToString()].Formula = string.Format("=DropDownList!${0}${1}:${2}${3}", char1, num1, char2, num2);
                    */

                    byte[] bin = pck.GetAsByteArray();


                    HttpContext.Current.Response.Clear();

                    HttpContext.Current.Response.AppendHeader("Content-Length", bin.Length.ToString());
                    HttpContext.Current.Response.AppendHeader("Content-Disposition", String.Format("attachment; filename={0}", fileName)
                        );
                    HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                    HttpContext.Current.Response.BinaryWrite(bin);
                    HttpContext.Current.Response.End();

                    //pck.Save();

                    return "Completed successfully";
                }
                catch (Exception e)
                {

                    return string.Format(@"Failed to generate XLSX file : {0}", e.Message);
                }
            }
        }



        private DataSet ReadSODataFromDB()
        {
            DataSet SODataRows = new DataSet("SOData");
            using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString))
            {
                con.Open();

                SqlDataAdapter JDBase = new SqlDataAdapter(string.Format(@"select 
		  a.Sales_Geo
		, a.Chnl_Ptnr_Engl_Nm
		, a.Rpt_Disti_Direct_Nm
		, a.Rpt_Disti_Direct_Country_Region_Cd
		, isnull(a.Rpt_Disti_Indirect_Nm,'') as Rpt_Disti_Indirect_Nm
		, a.SO_Account_Name
		, a.Forecast_Nm
		, a.CHM_Old_MM
		, a.Division_Short_Nm
		, a.Operation_Cd
		, a.Trns_Record_Type_Cd
		, a.Trns_Record_Type_Desc
		, a.Trns_Record_Eff_Yyyyqq
		, a.Trns_Record_Eff_Yyyymm
		, a.Trns_Record_Eff_Yyyyww
		, a.Trns_Record__Create_Yyyyqq
		, a.Trns_Record_Create_Yyyymm
		, a.Trns_Record_Create_Yyyyww
		, a.Rsls_Qty
		, isnull(a.Total_Stndr_Disti_Cost_USD_Amt,'') as Total_Stndr_Disti_Cost_USD_Amt
		, isnull(a.Total_Rsls_Price_USD_Amt,'') as  Total_Rsls_Price_USD_Amt
		, a.MM
		, a.assembly_level
		, a.name as Product_Name
		, isnull (a.AccountName,'') as AccountName
		, isnull (a.AccountOwner   , '') as AccountOwner   
		, isnull (a.EffectiveFrom  , '') as EffectiveFrom  
		, isnull (a.EffectiveTo	   , '') as EffectiveTo	   
		, isnull (a.ManualDecision , '') as ManualDecision 
        , a.transactionID
from DISTI.SalesOut_V a
where a.Trns_Record_Create_Yyyyww>='202001'

 "), con);

                JDBase.FillSchema(SODataRows, SchemaType.Source, "SOData");
                JDBase.Fill(SODataRows, "SOData");

                return SODataRows;

            }
        }


        private DataSet ReadSOAccountsFromDB()
        {
            DataSet SOAccounts = new DataSet("SOAccounts");
            using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString))
            {
                con.Open();

                SqlDataAdapter sAdptr = new SqlDataAdapter(string.Format(@"select 
		                                                                              a.AccountName
		                                                                          
                                                                            from [DISTI].[Accounts] a
                                                                        union 
                                                                        select ''


 "), con);

                sAdptr.FillSchema(SOAccounts, SchemaType.Source, "SOAccounts");
                sAdptr.Fill(SOAccounts, "SOAccounts");

                return SOAccounts;

            }
        }

        private DataSet ReadSOOwnersFromDB()
        {
            DataSet SOOwnrs = new DataSet("SOOwners");
            using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString))
            {
                con.Open();

                SqlDataAdapter sAdptr = new SqlDataAdapter(string.Format(@"select 
		                                                                              a.[AccountOwner]
		                                                                          
                                                                            from [DISTI].[Owners] a
                                                                        union 
                                                                        select ''


 "), con);

                sAdptr.FillSchema(SOOwnrs, SchemaType.Source, "SOOwners");
                sAdptr.Fill(SOOwnrs, "SOOwners");

                return SOOwnrs;

            }
        }


        private DataTable CreateSODataTopHeaderDataTable()
        {
            DataTable dataTable = new DataTable();

            //add three colums to the datatable
            dataTable.Columns.Add("Geo", typeof(string));
            dataTable.Columns.Add("Channel", typeof(string));
            dataTable.Columns.Add("Direct", typeof(string));
            dataTable.Columns.Add("Direct_Country_Region_Cd", typeof(string));
            
            dataTable.Columns.Add("Inderect", typeof(string));
            dataTable.Columns.Add("SO Name", typeof(string));
            dataTable.Columns.Add("Forecast Nm", typeof(string));
            dataTable.Columns.Add("CHM old MM", typeof(string));
            dataTable.Columns.Add("Division", typeof(string));
            dataTable.Columns.Add("Operation CD", typeof(string));
            dataTable.Columns.Add("Trns Record Type CD", typeof(string));
            dataTable.Columns.Add("Trns Record Type Desc", typeof(string));
            dataTable.Columns.Add("Trns_Record_Eff_Yyyyqq", typeof(string));
            dataTable.Columns.Add("Trns_Record_Eff_Yyyymm", typeof(string));
            dataTable.Columns.Add("Trns_Record_Eff_Yyyyww", typeof(string));
            dataTable.Columns.Add("Trns_Record__Create_Yyyyqq", typeof(string));
            dataTable.Columns.Add("Trns_Record_Create_Yyyymm", typeof(string));
            dataTable.Columns.Add("Trns_Record_Create_Yyyyww", typeof(string));
            dataTable.Columns.Add("Rsls_Qty", typeof(string));
            dataTable.Columns.Add("Total_Stndr_Disti_Cost_USD_Amt", typeof(string));
            dataTable.Columns.Add("Total_Rsls_Price_USD_Amt", typeof(string));

            dataTable.Columns.Add("MM", typeof(string));
            dataTable.Columns.Add("assembly_level", typeof(string));
            dataTable.Columns.Add("Product_Name", typeof(string));
            dataTable.Columns.Add("AccountName", typeof(string));
            dataTable.Columns.Add("AccountOwner", typeof(string));
            dataTable.Columns.Add("EffectiveFrom", typeof(string));
            dataTable.Columns.Add("EffectiveTo", typeof(string));

            dataTable.Columns.Add("ManualDecision", typeof(string));
            dataTable.Columns.Add("transactionID", typeof(string));

            
            dataTable.Rows.Add("Geo", "Channel", "Direct"
                                , "Direct_Country_Region_Cd"
                                , "Indirect"
                                , "SO Name", "Forecast Nm", "CHM Old MM"
                                , "Division", "Operation Cd", "Trns Record Type Cd"
                                , "Trns Record Type Desc"
                                , "Trns_Record_Eff_Yyyyqq"
                                , "Trns_Record_Eff_Yyyymm"
                                , "Trns_Record_Eff_Yyyyww"
                                , "Trns_Record__Create_Yyyyqq"
                                , "Trns_Record_Create_Yyyymm"
                                , "Trns_Record_Create_Yyyyww"
                                , "Rsls_Qty"
                                , "Total_Stndr_Disti_Cost_USD_Amt"
                                , "Total_Rsls_Price_USD_Amt"
                                , "MM"
                                , "assembly_level"
                                , "Product_Name"
                                , "AccountName"
                                , "AccountOwner"
                                , "EffectiveFrom"
                                , "EffectiveTo"
                                , "Manual Decision"
                                , "transactionID"
                                );



            return dataTable;


        }


    }




}
   
