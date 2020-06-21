using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Configuration;
using System.Web.Services;

namespace OpsInvoices
{
    /// <summary>
    /// Summary description for OpsInvoicesServices
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class OpsInvoicesServices : System.Web.Services.WebService
    {

        [WebMethod]
        public string UpdatePOStatus(string poNumber, string userName, string status, string remarks)
        {
            int replyRec = 0;
            try
            {
                using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString))
                {

                    using (SqlCommand cmd = new SqlCommand("[Invoices].[UPDATE_PO_STATUS_SP_V2]", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("PO", SqlDbType.VarChar, 100).Value = poNumber;
                        cmd.Parameters.Add("user", SqlDbType.VarChar, 100).Value = userName;
                        cmd.Parameters.Add("Status", SqlDbType.VarChar, 100).Value = status;
                        cmd.Parameters.Add("Remarks", SqlDbType.VarChar, 100).Value = remarks;

                        if (con.State != ConnectionState.Open)
                        {
                            con.Open();
                        }
                        replyRec = cmd.ExecuteNonQuery();

                    }
                    if (replyRec > 0)
                        return String.Concat(poNumber, ": Status updated successfuly to ", status);
                    else
                        return "Failed";
                }
            }
            catch (Exception e)
            {
                return String.Format("PO  update failed with: {0}", e.Message);

            }



        }

        [WebMethod]

        public DataSet GetAllStatuses()
        {
            

                using (SqlConnection conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString))
                {
                    SqlDataAdapter adp = new SqlDataAdapter("Select POStatus from Invoices.POStatuses where IsActive=1", conn);
                    DataSet ds = new DataSet();
                    adp.Fill(ds, "PO Statuses");
                    return ds;
                }




        }



    }
}
