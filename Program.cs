using ClosedXML.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace PendingBillGenerationReport
{
    class Program
    {
      

        static void Main(string[] args)
        {
            try
            {
                clsCommon cls = new clsCommon();
                Console.WriteLine("===================== Get Pending Bill Generation Report  Start ===============");

                DataSet BillingSummary = new DataSet();
                DataSet CategorySummary = new DataSet();
                DataSet BillingDetail = new DataSet();
                BillingSummary = GetPendingBillGenerationReport(5);
                CategorySummary = GetPendingBillGenerationReport(6);
                BillingDetail = GetPendingBillGenerationReport(4);

                var GetPendingBillReportExcel = GenerateExcelGetPendingBillReport(BillingSummary, CategorySummary, BillingDetail);
                
                if (GetPendingBillReportExcel != null)
                {
                    getmailinformation(GetPendingBillReportExcel);
                    Console.WriteLine("=====================Get Pending Bill Generation Report  Details Console End ===============");
                }
                else
                {
                    Console.WriteLine("===================== Get Pending Bill Generation Console End ===============");
                }

            }
            catch (Exception ex)
            {

                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex, true);
                var stackFrame = trace.GetFrame(0);
                var LineNumber = stackFrame.GetFileLineNumber();
                ErrorHistoryLog("PendingBillGeneration Report ", "Main Function", ex.GetType().ToString(), ex.Message, ex.StackTrace, -4, " Error Line No Is :" + LineNumber);

                string EmailForError = "jignisha.mohaniya@pe.soultions";

                string Body = "<div dir='ltr'>Dear Team,<br><br></div>";
                Body = Body + "<div>Error occured in PendingBillGeneration Report Details Console Read.<br>";
                Body = Body + "<br>Please find below details : <br>";
                Body = Body + "<br><b>Exception : </b>" + ex.Message;
                Body = Body + "<br><b>Stack Trace : </b>" + ex.StackTrace + "</div>";
                Body = Body + "<br><br><div>Thanks,<br>Support Team.</div> ";

                SendMail(EmailForError, "Error in PendingBillGeneration Report  Details Report Path Console Read" + "_" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss"), Body, true, "", "", "", false);
                Console.WriteLine("Exception.....");
            }

        }

        static DataSet GetPendingBillGenerationReport(int type)
        {
            DataSet ds = new DataSet();

            try
            {
                SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ReportConnectionString"].ConnectionString);
                conn.Open();

                SqlCommand cmd = new SqlCommand("Rpt_PendingBillCustomerReport", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 0;

                cmd.Parameters.Add("@Date", SqlDbType.DateTime).Value = null;
                cmd.Parameters.Add("@Flag", SqlDbType.Int).Value = null;
                cmd.Parameters.Add("@Type", SqlDbType.VarChar).Value = type;
                SqlDataAdapter da = new SqlDataAdapter(cmd);

                DataTable dt = new DataTable();
                da.Fill(ds);
                conn.Close();

                return ds;

            }
            catch (Exception ex)
            {
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex, true);
                var stackFrame = trace.GetFrame(0);
                var LineNumber = stackFrame.GetFileLineNumber();
                ErrorHistoryLog("PendingBillGeneration", "GetPendingBillGeneration", ex.GetType().ToString(), ex.Message, ex.StackTrace, -4, " Error Line No Is :" + LineNumber);

                throw new Exception("GetPendingBillGeneration" + "(" + ex.Message + ")");
            }

        }
        static string GenerateExcelGetPendingBillReport(DataSet BillingSummary, DataSet CategorySummary, DataSet BillingDetail)
        {
            string fileNameWithExt = "";
            try
            {
                if (BillingSummary != null && BillingSummary != null)
                {
                    Console.WriteLine("===================== Start Generate excel For Pending Bill Generation Report =====================");
                    clsCommon common = new clsCommon();
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    string PendingBillReportFilePath = getconfigSetting("PendingBillReportFilePath");
                    string body = string.Empty;
                    string FileSavePath = "", newfname = "", EXT = ".xlsx";
                    string Filepath = Path.Combine(PendingBillReportFilePath + "\\Content\\");
                    if (!System.IO.Directory.Exists(Filepath))
                    {
                        System.IO.Directory.CreateDirectory(Filepath);
                    }

                    string[] files = Directory.GetFiles(Filepath);
                    if (files != null)
                    {
                        foreach (string n in files)
                        {
                            EXT = Path.GetExtension(n.Trim());
                            if (EXT.ToLower() == ".xlsx")
                            {
                                var FileName = System.IO.Path.GetFileName(n);
                                if (FileName == "Demo.xlsx")
                                {
                                    try
                                    {
                                        var strFilePath = FileName.Split('.');
                                        newfname = "PendingBillGeneration_" + DateTime.Now.ToString("ddMMyyyy");
                                        newfname = newfname.Replace("-", "");
                                        fileNameWithExt = newfname + EXT;

                                        string pathdesktop = Path.Combine(PendingBillReportFilePath, "PendingBillGenerationReport");
                                        FileSavePath = Path.Combine(pathdesktop, fileNameWithExt);

                                        if (!System.IO.Directory.Exists(pathdesktop))
                                        {
                                            System.IO.Directory.CreateDirectory(pathdesktop);
                                        }
                                        if (File.Exists(Path.Combine(pathdesktop, fileNameWithExt)))
                                        {
                                            File.SetAttributes(pathdesktop, FileAttributes.Normal);
                                            File.Delete(Path.Combine(pathdesktop, fileNameWithExt));
                                        }

                                        System.IO.File.Copy(n, FileSavePath);
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex, true);
                                        var stackFrame = trace.GetFrame(0);
                                        var LineNumber = stackFrame.GetFileLineNumber();
                                        ErrorHistoryLog("PendingBillReportFilePath Report ", "Create excel file on location", ex.GetType().ToString(), ex.Message, ex.StackTrace, -4, " Error Line No Is :" + LineNumber);

                                        throw new Exception("Create excel file on location" + "(" + ex.Message + ")");
                                    }
                                }
                            }
                        }
                    }

                    FileInfo fileNew = new FileInfo(FileSavePath);
                    using (IXLWorkbook wb = new XLWorkbook())
                    {
                        DataSet[] dataSets = { BillingSummary, CategorySummary, BillingDetail };

                        IXLWorksheet summaryWs = null;
                        IXLWorksheet dtlWs = null;

                        int summaryRow = 1;
                        int nhhRow = 1;

                        for (int i = 0; i < dataSets.Length; i++)
                        {
                            DataTable dt = dataSets[i].Tables[0];
                            if (i == 0 || i == 1)
                            {
                                if (summaryWs == null)
                                {
                                    summaryWs = wb.Worksheets.Add("Summary");
                                }
                                else
                                {
                                    summaryRow++;
                                }
                                int col = 1;
                                foreach (DataColumn column in dt.Columns)
                                {
                                    summaryWs.Cell(summaryRow, col).Value = column.ColumnName;
                                    summaryWs.Cell(summaryRow, col).Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent5);
                                    summaryWs.Cell(summaryRow, col).Style.Font.Bold = true;
                                    summaryWs.Cell(summaryRow, col).Style.Font.FontColor = XLColor.White;
                                    summaryWs.Cell(summaryRow, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                    col++;
                                }

                                summaryRow++;
                                summaryWs.Cell(summaryRow, 1).InsertData(dt.AsEnumerable());
                                summaryRow += dt.Rows.Count;
                            }
                            else
                            {
                                if (dtlWs == null)
                                {
                                    dtlWs = wb.Worksheets.Add("Detail");
                                    int col = 1;
                                    foreach (DataColumn column in dt.Columns)
                                    {
                                        dtlWs.Cell(nhhRow, col).Value = column.ColumnName;
                                        dtlWs.Cell(nhhRow, col).Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent5);
                                        dtlWs.Cell(nhhRow, col).Style.Font.Bold = true;
                                        dtlWs.Cell(nhhRow, col).Style.Font.FontColor = XLColor.White;
                                        dtlWs.Cell(nhhRow, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                        col++;
                                    }
                                    nhhRow++;
                                }

                                dtlWs.Cell(nhhRow, 1).InsertData(dt.AsEnumerable());
                                nhhRow += dt.Rows.Count;
                            }
                        }
                        summaryWs?.Columns().AdjustToContents();
                        dtlWs?.Columns().AdjustToContents();
                        wb.SaveAs(FileSavePath);
                    }
                    return fileNameWithExt;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex, true);
                var stackFrame = trace.GetFrame(0);
                var LineNumber = stackFrame.GetFileLineNumber();
                ErrorHistoryLog("PendingBillGeneration Report ", "Generate Excel", ex.GetType().ToString(), ex.Message, ex.StackTrace, -4, " Error Line No Is :" + LineNumber);

                throw new Exception("Generate Excel" + "(" + ex.Message + ")");
            }

            return fileNameWithExt;
        }
        static void getmailinformation(string GetPendingBillReportExcel)
        {

            try
            {
                string PendingBillReportFilePath = getconfigSetting("PendingBillReportFilePath");
                string body = string.Empty;
                string FileSavePath = "";
                string folderPath = Path.Combine(PendingBillReportFilePath, "PendingBillReportFilePath");

                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }

                FileSavePath = Path.Combine(folderPath, GetPendingBillReportExcel);
                var MailTemplatePath = PendingBillReportFilePath + @"\Content\MailTemplate.html";
                using (StreamReader reader = new StreamReader(MailTemplatePath))
                {
                    body = reader.ReadToEnd();
                }

                body = body.Replace("[date]", DateTime.Now.ToString("dd/MM/yyyy"));
                var strMailSubject = "Pending Bill Generation - " + DateTime.Today.ToString("dd/MM/yyyy");

                var ToEmailId = getconfigSetting("PendingBillGenerateToEmailId");
                var CcEmailId = getconfigSetting("PendingBillGenerateCCEmailId");

                var bccEmailId = getconfigSetting("");
                SendMail(ToEmailId, strMailSubject, body, true, FileSavePath, CcEmailId, bccEmailId, true);

            }
            catch (Exception ex)
            {
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex, true);
                var stackFrame = trace.GetFrame(0);
                var LineNumber = stackFrame.GetFileLineNumber();
                ErrorHistoryLog("PendingBillGeneration", "getmailinformation", ex.GetType().ToString(), ex.Message, ex.StackTrace, -4, " Error Line No Is :" + LineNumber);

                throw new Exception("getmailinformation" + "(" + ex.Message + ")");
            }
        }

        static void SendMail(string ToMail, string strSubject, string strBody, bool bitIsBodyHtml, string GetPendingBillReportExcel, string CcEmailId, string bccEmailId, bool bIsSent)
        {
            try
            {
                clsCommon objCommon = new clsCommon();
                clsCommon emailcls = new clsCommon();
                emailcls = objCommon.Getemailcredentail("General");

                string MailEmailId = ""; string MailEMailPassword = ""; string MailEmailSMTP = ""; int PortNo = 0;
                bool IsSSLAllow = true;

                if (emailcls != null)
                {
                    MailEmailId = emailcls.strfromid;
                    MailEMailPassword = emailcls.strfrompassword;
                    MailEmailSMTP = emailcls.strfromsmtp;
                    PortNo = emailcls.intportno;
                    IsSSLAllow = emailcls.IsSSLallow;
                }
                MailMessage mail = new MailMessage();
                if (ToMail != "")
                {
                    string[] OwnerEmailIdList = ToMail.Split(',');
                    foreach (var item in OwnerEmailIdList)
                    {
                        if (item != null && item != "")
                        {
                            mail.To.Add(item);
                        }
                    }
                }
                if (CcEmailId != "")
                {
                    string[] OwnerEmailIdList = CcEmailId.Split(',');
                    foreach (var item in OwnerEmailIdList)
                    {
                        if (item != null && item != "")
                        {
                            mail.CC.Add(item);
                        }
                    }
                }
                if (bccEmailId != "")
                {
                    string[] OwnerEmailIdList = bccEmailId.Split(',');
                    foreach (var item in OwnerEmailIdList)
                    {
                        if (item != null && item != "")
                        {
                            mail.Bcc.Add(item);
                        }
                    }
                }


                mail.From = new MailAddress(MailEmailId);
                mail.Subject = strSubject;
                mail.IsBodyHtml = bitIsBodyHtml;
                SmtpClient smtp1 = new SmtpClient(MailEmailSMTP, PortNo);
                smtp1.Host = MailEmailSMTP;
                smtp1.Credentials = new System.Net.NetworkCredential(MailEmailId, MailEMailPassword);

                if (GetPendingBillReportExcel != null && GetPendingBillReportExcel.Trim() != "")
                {
                    if (GetPendingBillReportExcel != null && GetPendingBillReportExcel != "" && System.IO.File.Exists(GetPendingBillReportExcel))
                    {
                        mail.Attachments.Add(new Attachment(GetPendingBillReportExcel));
                    }
                }

                mail.Body = strBody;
                smtp1.EnableSsl = IsSSLAllow;
                string strInvoiceType = "Pending Bill Generation Report ";

                System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                smtp1.Send(mail);
                string strFilePath = GetPendingBillReportExcel;
                int intEmailTrackId = InsertEmailTracking("Pending Bill Generation Report", "Pending Bill Generation Report", MailEmailId, ToMail, CcEmailId, bccEmailId,
                             mail.Subject, mail.Body, strFilePath, 0, 2, 0, 0, -4, 0, bIsSent, strInvoiceType);

                var result = System.Net.Mail.DeliveryNotificationOptions.OnSuccess;
                if (result.ToString() == "OnSuccess")
                {
                    int id = InsertEmailTracking("Pending Bill Generation Report", "Pending Bill Generation Report", MailEmailId, ToMail, CcEmailId, bccEmailId,
                             mail.Subject, mail.Body, strFilePath, 1, 2, 0, 0, -4, intEmailTrackId, bIsSent, strInvoiceType);
                }
                else
                {

                    int id = InsertEmailTracking("Pending Bill Generation Report", "Pending Bill Generation Report", MailEmailId, ToMail, CcEmailId, bccEmailId,
                             mail.Subject, mail.Body, strFilePath, 2, 2, 0, 0, -4, intEmailTrackId, bIsSent, strInvoiceType);
                }

            }
            catch (Exception ex)
            {
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex, true);
                var stackFrame = trace.GetFrame(0);
                var LineNumber = stackFrame.GetFileLineNumber();
                ErrorHistoryLog("Pending Bill Generation Report", "Exception in sending error mail.", ex.GetType().ToString(), ex.Message, ex.StackTrace, -4, " Error Line No Is :" + LineNumber);

                throw new Exception("Exception in sending error mail." + "(" + ex.Message + ")");
            }
        }
        public static string getconfigSetting(string strAppKey)
        {
            SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["ReportConnectionString"].ConnectionString);
            try
            {
                string strAppKeyVal = "";
                if (con.State == ConnectionState.Closed)
                {
                    con.Open();
                }
                SqlCommand cmd = new SqlCommand("sp_GetConfigKeyValue", con);
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("@strKey", SqlDbType.NVarChar).Value = strAppKey;

                SqlDataAdapter da = new SqlDataAdapter(cmd);
                cmd.CommandTimeout = 0;
                DataTable dt = new DataTable();
                da.Fill(dt);
                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }

                if (dt.Rows.Count > 0)
                {
                    strAppKeyVal = dt.Rows[0]["strValue"] != null && dt.Rows[0]["strValue"].ToString() != "" ? dt.Rows[0]["strValue"].ToString() : "";
                }
                return strAppKeyVal;
            }
            catch (Exception ex)
            {
                System.Diagnostics.StackTrace trace = new System.Diagnostics.StackTrace(ex, true);
                var stackFrame = trace.GetFrame(0);
                var LineNumber = stackFrame.GetFileLineNumber();
                ErrorHistoryLog("Pending Bill Generation Report", "Get config setting data", ex.GetType().ToString(), ex.Message, ex.StackTrace, -4, " Error Line No Is :" + LineNumber);

                throw new Exception("Get config setting data" + "(" + ex.Message + ")");
            }
        }

        public static void ErrorHistoryLog(string StrControllerName, string StrMethodName, string StrExceptionType, string StrMessege, string StrStackTrace, int intWorkingUserId, string strScenario)
        {
            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ReportConnectionString"].ConnectionString);
            try
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("sp_ExceptionDtl_Insert", conn);
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("@StrControllerName", SqlDbType.VarChar).Value = "PendingBillGeneration Report";
                cmd.Parameters.AddWithValue("@StrMethodName", SqlDbType.VarChar).Value = StrMethodName;
                cmd.Parameters.AddWithValue("@StrExceptionType", SqlDbType.VarChar).Value = StrExceptionType;
                cmd.Parameters.AddWithValue("@StrMessege", SqlDbType.VarChar).Value = StrMessege;
                cmd.Parameters.AddWithValue("@StrStackTrace", SqlDbType.VarChar).Value = StrStackTrace;
                cmd.Parameters.AddWithValue("@intWorkingUserId", SqlDbType.BigInt).Value = intWorkingUserId;
                cmd.Parameters.AddWithValue("@strScenario", SqlDbType.VarChar).Value = strScenario;

                int id = Convert.ToInt32(cmd.ExecuteNonQuery());
                conn.Close();
            }
            catch (Exception ex)
            {

            }
        }

        public static int InsertEmailTracking(string strController, string strAction, string strEmailFrom, string strEmailTo,
        string strEmailCC, string strEmailBCC, string strSubject, string strContent, string strAttachmentPath,
        int intStatus, int intReceivedSent, int intVendorId, int intCompanyId, int intCreatedBy, int intEmailTrackId, bool bIsSent, string strInvoiceType)
        {
            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ReportConnectionString"].ConnectionString);
            clsCommon cls = new clsCommon();
            try
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("Sp_InsertEmailTracking", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandTimeout = 0;

                cmd.Parameters.AddWithValue("@intId", intEmailTrackId);
                cmd.Parameters.AddWithValue("@strController", strController);
                cmd.Parameters.AddWithValue("@strAction", strAction);
                cmd.Parameters.AddWithValue("@strEmailFrom", strEmailFrom);
                cmd.Parameters.AddWithValue("@strEmailTo", strEmailTo);
                cmd.Parameters.AddWithValue("@strEmailCC", strEmailCC);
                cmd.Parameters.AddWithValue("@strEmailBCC", strEmailBCC);
                cmd.Parameters.AddWithValue("@strSubject", strSubject);
                cmd.Parameters.AddWithValue("@strContent", strContent);
                cmd.Parameters.AddWithValue("@strAttachmentPath", strAttachmentPath);
                cmd.Parameters.AddWithValue("@intStatus", intStatus); /// 1 Success 2 Fail
                cmd.Parameters.AddWithValue("@intReceivedSent", intReceivedSent); /// 1 Received 2 Sent
                cmd.Parameters.AddWithValue("@intVendorId", intVendorId);
                cmd.Parameters.AddWithValue("@intCompanyId", intCompanyId);
                cmd.Parameters.AddWithValue("@intCreatedBy", intCreatedBy);
                cmd.Parameters.AddWithValue("@dtCreateDate", DateTime.Now);
                cmd.Parameters.AddWithValue("@strInvoiceType", strInvoiceType);
                cmd.Parameters.AddWithValue("@bIsSent", bIsSent); // 1 success message, 0 error message
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                cmd.CommandTimeout = 0;
                da.ReturnProviderSpecificTypes = true;
                DataTable dt = new DataTable();
                da.Fill(dt);
                conn.Close();
                if (dt.Rows.Count > 0)
                {
                    cls.intEmailTrackId = Convert.ToInt32(dt.Rows[0][0].ToString());
                }
                return cls.intEmailTrackId;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
