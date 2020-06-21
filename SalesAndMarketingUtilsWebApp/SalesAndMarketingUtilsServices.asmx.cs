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
        
        [WebMethod]
        public string  CreateSpreadsheet(int planId)
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
                        planDetails = CreateNewPlan(userName);
                    }
                    else
                    {
                        planDetails = GetPlanHeader(planId);
                    }

                    var JDPlanWorksheet = pck.Workbook.Worksheets.Add("JD_Plan");

                    if (planDetails.Item1 == -1)
                    {
                        throw new Exception("Selected plan does not exists");

                    }

                    var headerDt = CreateTopHeaderDataTable();
                    JDPlanWorksheet.Cells["A2"].LoadFromDataTable(headerDt, false);


                    JDPlanWorksheet.Cells["A1"].Value = "Plan ID";
                    JDPlanWorksheet.Cells["B1"].Value = planDetails.Item1;

                    JDPlanWorksheet.Cells["E2:I2"].Merge = true;
                    JDPlanWorksheet.Cells["J2:N2"].Merge = true;
                    JDPlanWorksheet.Cells["O2:S2"].Merge = true;

                    JDPlanWorksheet.Cells["E2:I2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    JDPlanWorksheet.Cells["J2:N2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    JDPlanWorksheet.Cells["O2:S2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    //Borders
                    SetBorders(JDPlanWorksheet, "E2:I2");
                    SetBorders(JDPlanWorksheet, "J2:N2");
                    SetBorders(JDPlanWorksheet, "O2:S2");
                    SetBorders(JDPlanWorksheet, "T2:U2");

                    //Background
                    SetBackgroud(JDPlanWorksheet, "E2:I2", "#E4F1F6");
                    SetBackgroud(JDPlanWorksheet, "J2:N2", "#91C5D9");
                    SetBackgroud(JDPlanWorksheet, "O2:S2", "#84B1C2");
                    SetBackgroud(JDPlanWorksheet, "T2:U2", "#E4F1F6");


                    JDPlanWorksheet.Cells["A2:V2"].Style.Font.Bold = true;

                    //Secondary header
                    headerDt = CreateSecondaryHeaderDataTable(planDetails.Item2);
                    JDPlanWorksheet.Cells["A3"].LoadFromDataTable(headerDt, false);

                    //Align 
                    JDPlanWorksheet.Cells["E3:U3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    JDPlanWorksheet.Cells["A3:D3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                    //Borders
                    SetBorders(JDPlanWorksheet, "A3:U3");

                    //Background
                    SetBackgroud(JDPlanWorksheet, "A3:U3", "#E4F1F6");
                    SetBackgroud(JDPlanWorksheet, "J3:N3", "#91C5D9");
                    SetBackgroud(JDPlanWorksheet, "O3:S3", "#84B1C2");
                    SetBackgroud(JDPlanWorksheet, "T3:U3", "#E4F1F6");


                    JDPlanWorksheet.Cells["A3:T3"].Style.Font.Bold = true;


                    // populate spreadsheet with data

                    var ds = ReadJDBasePlanFromDB(planDetails.Item1);
                    var dt = ds.Tables[0];
                    var tabRowcount = dt.Rows.Count;

                    JDPlanWorksheet.Cells["A4"].LoadFromDataTable(dt, false);

                    JDPlanWorksheet.Cells[JDPlanWorksheet.Dimension.Address].AutoFitColumns();

                    //Borders
                    SetBorders(JDPlanWorksheet, "A4:U" + (3 + tabRowcount).ToString());


                    //Align 
                    JDPlanWorksheet.Cells["A4:D" + (3 + tabRowcount).ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    JDPlanWorksheet.Cells["E4:U" + (3 + tabRowcount).ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    //Background
                    SetBackgroud(JDPlanWorksheet, "J4:N" + (3 + tabRowcount).ToString(), "#F0F2F5");
                    SetBackgroud(JDPlanWorksheet, "O4:S" + (3 + tabRowcount).ToString(), "#FCFCFC");
                    SetBackgroud(JDPlanWorksheet, "T4:U" + (3 + tabRowcount).ToString(), "#FCFCFC");
                    SetBackgroud(JDPlanWorksheet, "A4:I" + (3 + tabRowcount).ToString(), "#FCFCFC");

                    JDPlanWorksheet.Protection.IsProtected = true; //--------Protect whole sheet
                    JDPlanWorksheet.Cells["J4:N" + (3 + tabRowcount).ToString()].Style.Locked = false; //-------Unlock 3rd column

                    //Add a List validation to the C column
                    var val3 = JDPlanWorksheet.DataValidations.AddIntegerValidation("J4:N" + (3 + tabRowcount).ToString());
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


        private DataSet ReadJDBasePlanFromDB(int planId)
        {
            DataSet JDBaseRows = new DataSet("JDBase");
            using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString))
            {
                SqlDataAdapter JDBase = new SqlDataAdapter(string.Format(@"Select      
                                                                          Year
                                                                        , SKU
                                                                        , Product_Name
                                                                        , assembly_level
                                                                        , Forecast_Q1
                                                                        , Forecast_Q2
                                                                        , Forecast_Q3
                                                                        , Forecast_Q4

                                                                        , Forecast_Weeks20

                                                                        , Cast(null as int) as JD_Q1
                                                                        , Cast(null as int) as JD_Q2
                                                                        , Cast(null as int) as JD_Q3
                                                                        , Cast(null as int) as JD_Q4
                                                                        , Cast(null as int) as JD_Weeks20

                                                                        , BnB_Q1
                                                                        , BnB_Q2
                                                                        , BnB_Q3
                                                                        , BnB_Q4
                                                                        , Backlog_Weeks20
                                                                        , InvQty
                                                                        , WareHouseGoal
    
                                                              From JD.JD_Plan_Details_V
                                                              where planId = {0}",planId), con);

                JDBase.FillSchema(JDBaseRows, SchemaType.Source, "JDBaseRows");
                JDBase.Fill(JDBaseRows, "JDBaseRows");

                return JDBaseRows;

            }
        }

        private DataTable CreateTopHeaderDataTable()
        {
            DataTable dataTable = new DataTable();

            //add three colums to the datatable
            dataTable.Columns.Add("Year", typeof(int));
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
            dataTable.Columns.Add("BnB1", typeof(string));
            dataTable.Columns.Add("BnB2", typeof(string));
            dataTable.Columns.Add("BnB3", typeof(string));
            dataTable.Columns.Add("BnB4", typeof(string));
            dataTable.Columns.Add("BnB5", typeof(string));

            dataTable.Columns.Add("BOH", typeof(string));
            dataTable.Columns.Add("Inv Goal", typeof(string));


            dataTable.Rows.Add(null, null, null, null
                                , "Forecast", "Forecast", "Forecast", "Forecast", "Forecast"
                                , "JD", "JD", "JD", "JD", "JD"
                                , "BnB", "BnB", "BnB", "BnB", "BnB"
                                , "BOH"
                                , "Inv Goal");



            return dataTable;


        }

        private DataTable CreateSecondaryHeaderDataTable(int ww)
        {
            DataTable dataTable = new DataTable();

            //add three colums to the datatable
            dataTable.Columns.Add("Year", typeof(string));
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
            dataTable.Columns.Add("JD_20weeks", typeof(string));
            dataTable.Columns.Add("BnB_Q1", typeof(string));
            dataTable.Columns.Add("BnB_Q2", typeof(string));
            dataTable.Columns.Add("BnB_Q3", typeof(string));
            dataTable.Columns.Add("BnB_Q4", typeof(string));
            dataTable.Columns.Add("BnB_20weeks", typeof(string));

            dataTable.Columns.Add("BOH", typeof(string));
            dataTable.Columns.Add("Inv Goal", typeof(string));


            dataTable.Rows.Add("Year", "Sku", "Product_Name", "assembly_level"
                                , "Q1", "Q2", "Q3", "Q4", "20 weeks"
                                , "Q1", "Q2", "Q3", "Q4", "20 weeks"
                                , "Q1", "Q2", "Q3", "Q4", "20 weeks"
                                , ww.ToString()
                                , "");



            return dataTable;


        }

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

        private Tuple<int, int,string> CreateNewPlan(string userName)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString))
                {
                    DataSet ds = new DataSet("PlanHeader");
                    using (SqlCommand cmd = new SqlCommand("[JD].[GenerateNewPlan_SP]", con))
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

        private Tuple<int, int, string> GetPlanHeader(int planId)
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

    }
}
