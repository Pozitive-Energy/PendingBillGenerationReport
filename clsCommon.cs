using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace PendingBillGenerationReport
{
    class clsCommon
    {
        public string strfromid { get; set; }
        public string strfrompassword { get; set; }
        public string strfromsmtp { get; set; }
        public int intportno { get; set; }
        public bool IsSSLallow { get; set; }
        public int intEmailTrackId { get; set; }
        public clsCommon Getemailcredentail(string modulename)
        {
            clsCommon cls = new clsCommon();
            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ReportConnectionString"].ConnectionString);
            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }
                SqlCommand cmd = new SqlCommand("sp_getEmailCredential", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@strkey", modulename);

                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
                if (dt != null && dt.Rows.Count > 0)
                {
                    cls.strfromid = dt.Rows[0]["strFromId"] == null ? "" : dt.Rows[0]["strFromId"].ToString();
                    cls.strfrompassword = dt.Rows[0]["strFromPassword"] == null ? "" : dt.Rows[0]["strFromPassword"].ToString();
                    cls.strfromsmtp = dt.Rows[0]["strFromSmtp"] == null ? "" : dt.Rows[0]["strFromSmtp"].ToString();
                    cls.intportno = dt.Rows[0]["intPortNo"] == null ? 0 : Convert.ToInt32(dt.Rows[0]["intPortNo"].ToString());
                    cls.IsSSLallow = dt.Rows[0]["IsSSLAllow"] == null ? true : Convert.ToBoolean(dt.Rows[0]["IsSSLAllow"].ToString());
                }
                da.Dispose();
                dt.Dispose();
            }
            catch (Exception ex)
            {
                throw new Exception("Get email credentials" + "(" + ex.Message + ")");
            }
            return cls;
        }
    }
}
