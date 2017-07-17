using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Net;
using System.IO;
using Ionic.Zip;
using System.Data.SqlClient;
using System.Configuration;

using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;

namespace MailMerger
{
    public partial class CreateTemplate : System.Web.UI.Page
    {

        protected static DatabaseInfo Dbase = new DatabaseInfo();
        Microsoft.Office.Interop.Word.Application wrdApp;
        Microsoft.Office.Interop.Word.Document wrdDoc;
        Object oMissing = System.Reflection.Missing.Value;
        Object oFalse = false;
        //static string tmpFileLocation;
        //static string formatPath;
        static string dataSourcePath;
        static string tempSourcePath;
        bool isFormatFile = false;
        bool isSourceFile = false;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {

                //var currentDbase = Dbase.dbase(Request.QueryString["database"].Trim());
                DateTime startTime = DateTime.Now;
                DateTime endTime;
                Guid objGuid = Guid.NewGuid();
                try
                {
                    tempSourcePath = Convert.ToString(Request.QueryString["path"]);

                 
                    #region . Fetching datasource file .
                    // If path is http, then download file to local temporary directory.
                    if (tempSourcePath.StartsWith("http"))
                    {
                        WebClient webClient = new WebClient();
                        webClient.DownloadFile(new Uri(tempSourcePath), Server.MapPath("MailMergeDocs") + "\\" + Path.GetFileName(tempSourcePath));
                        dataSourcePath = Server.MapPath("MailMergeDocs") + "\\" + Path.GetFileName(tempSourcePath);
                        isSourceFile = true;
                    }
                    else if (Regex.IsMatch(tempSourcePath.Replace("/", "\\"), @"^(?:[a-zA-Z]\:|\\\\[\w\.]+\\[\w.]+)\\(?:[\w]+\\)*\w([\w.])+$"))
                    {
                        dataSourcePath = tempSourcePath.Replace("/", "\\");
                        if (!File.Exists(dataSourcePath)) throw new Exception("Datasource file does not exists. Please provide correct path");
                    }
                    else if (Regex.IsMatch(tempSourcePath.Replace("/", "\\"), @"^(?:[a-zA-Z]\:|\\[\w\.]+\\[\w.]+)\\(?:[\w]+\\)*\w([\w.])+$"))
                    {
                        dataSourcePath = tempSourcePath.Replace("\\", "\\\\");
                        if (!File.Exists(dataSourcePath)) throw new Exception("Datasource file does not exists. Please provide correct path");
                       
                    }
                    else if (Regex.IsMatch(tempSourcePath, @"^([a-zA-Z]\:|\\\\[^\/\\:*?<>|]+\\[^\/\\:*?<>|]+)(\\[^\/\\:*?<>|]+)+(\.[^\/\\:*?<>|]+)$"))
                    {
                                                          // @"^([a-zA-Z]\:|\\\\[^\/\\:*?<>|]+\\[^\/\\:*?<>|]+)(\\[^\/\\:*?<>|]+)+(\.[^\/\\:*?<>|]+)$"

                        // Check if file is accessable or present on the remote/network drive.
                        if (!File.Exists(tempSourcePath)) throw new Exception("File not found or accessible on " + tempSourcePath);
                     
                        File.Copy(tempSourcePath, Server.MapPath("MailMergeDocs") + "\\" + Path.GetFileName(tempSourcePath), true);
                        dataSourcePath = Server.MapPath("MailMergeDocs") + "\\" + Path.GetFileName(tempSourcePath);
                        isSourceFile = true;
                    }

                    #endregion

                    if (!string.IsNullOrEmpty(dataSourcePath))
                    {
                        try
                        {
                            MergeMail();
                            endTime = DateTime.Now;
                            TimeSpan tSpend = endTime - startTime;
                        }
                        catch (Exception ex)
                        {
                            Response.Write("MergeMail : " + ex.Message);
                        }
                      
                    }
                    else Response.Write("Incorrect file locations.");
                }
                catch (WebException wEx)
                {
                    #region . Exception Checks .

                    // Link for StatusCode : http://msdn.microsoft.com/en-us/library/system.net.httpstatuscode.aspx
                    if (((HttpWebResponse)wEx.Response).StatusCode.ToString().Equals("NotFound"))
                    {
                        Response.Write("One of the file was not found on the web. Please verify the link and try again.");
                    }
                    else if (((HttpWebResponse)wEx.Response).StatusCode.ToString().Equals("Unauthorized"))
                    {
                        Response.Write("Authentication is required to download the file.");
                    }
                    else if (((HttpWebResponse)wEx.Response).StatusCode.ToString().Equals("ProxyAuthenticationRequired"))
                    {
                        Response.Write("Requested proxy requires authentication to download the file.");
                    }
                    else if (((HttpWebResponse)wEx.Response).StatusCode.ToString().Equals("RequestTimeout"))
                    {
                        Response.Write("The remote server did not respond in time.");
                    }
                    else if (((HttpWebResponse)wEx.Response).StatusCode.ToString().Equals("Gone"))
                    {
                        Response.Write("The file is no longer available for download.");
                    }
                    else if (((HttpWebResponse)wEx.Response).StatusCode.ToString().Equals("InternalServerError"))
                    {
                        Response.Write("Internal server error occured at the remote server.");
                    }
                    else if (((HttpWebResponse)wEx.Response).StatusCode.ToString().Equals("ServiceUnavailable"))
                    {
                        Response.Write("The server is temporarily unavailable.");
                    }
                    else
                    {
                        Response.Write("Unable to download the file from the remote server.");
                    }

                    #endregion
                }
                catch (UnauthorizedAccessException uaeEx)
                {
                    Response.Write(uaeEx.Message);
                }
                catch (Exception ex)
                {
                    Response.Write("File Creation Error: "+ ex.Message);
                }
               
                if (isSourceFile && File.Exists(dataSourcePath)) File.Delete(dataSourcePath);
            }
        }

        private void MergeMail()
        {
            var currentDbase = Dbase.dbase(Request.QueryString["database"].Trim());
            try
            {
            string[] SchemeInfo=new string[2];
           
                // Get Scheme Name and Report ID.
                SchemeInfo = DatabaseInfo.GetSchemeReportID(tempSourcePath, currentDbase); // { scheme, ReportID } only two items

                if (SchemeInfo[0].Equals("-") || SchemeInfo[1].Equals("-"))
                    throw new Exception("Unable to trace Scheme Name and/or Report ID");
           

            wrdApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Selection wrdSelection;
            Microsoft.Office.Interop.Word.MailMerge wrdMailMerge;
            object SaveChanges = false;
            try
            {
                // Create an instance of Word and make it visible.
                wrdApp.Visible = false;

                string path = dataSourcePath;// Server.MapPath("MailMergeDocs") + "\\abc.docx";
                string templatePath = Server.MapPath("MailMergeDocs") + "\\DataSourceFile.html";
                wrdDoc = wrdApp.Documents.Open(path);

                if (wrdDoc != null)
                {
                    wrdDoc.Select();

                    wrdSelection = wrdApp.Selection;
                    wrdMailMerge = wrdDoc.MailMerge;
                    var oHeader = DatabaseInfo.GetCVSValues(SchemeInfo[0], SchemeInfo[1], currentDbase);
                    CreateExcelFile(oHeader, templatePath);

                    CreateMailMergeDataFile(templatePath);
                    //CreateMailMergeDataFile("http://localhost/MailMerger/MailMergeDocs/DataSourceFile.html");
                    wrdDoc.Save();
                    //wrdDoc.SaveAs(path, Word.WdSaveFormat.wdFormatTemplate);
                    wrdDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdSaveChanges, ref oMissing, ref oMissing);
                    wrdDoc = null;

                    Response.Clear();

                    System.Web.HttpContext c = System.Web.HttpContext.Current;

                    string archiveName = String.Format("archive-{0}.zip", DateTime.Now.ToString("yyyy-MMM-dd-HHmmss"));
                    Response.ContentType = "application/zip";
                    Response.AddHeader("content-disposition", "attachment; filename=" + archiveName);

                    using (ZipFile zip = new ZipFile())
                    {
                        // filesToInclude is a string[] or List<string>
                        zip.AddFile(path, "");
                        zip.AddFile(templatePath, "");

                        zip.Save(Response.OutputStream);
                    }
                    Response.Flush();


                    // Release References.
                    wrdSelection = null;
                    wrdMailMerge = null;
                    wrdDoc = null;
                    wrdApp.Quit(ref SaveChanges, false, null);
                    wrdApp = null;
                }
                else
                {
                    // Release References.
                    wrdSelection = null;
                    wrdMailMerge = null;
                    wrdDoc = null;
                    wrdApp.Quit(ref SaveChanges, false, null);
                    wrdApp = null;
                    throw new Exception("Could not open the document");
                }
            }
            catch (COMException comex)
            {
                // Release References.
                wrdSelection = null;
                wrdMailMerge = null;
                wrdDoc = null;
                wrdApp.Quit(ref SaveChanges, false, null);
                wrdApp = null;
                throw new Exception(comex.Message);
            }
            catch (Exception)
            {
                // Release References.
                wrdSelection = null;
                wrdMailMerge = null;
                wrdDoc = null;
                wrdApp.Quit(ref SaveChanges, false, null);
                wrdApp = null;
                throw new Exception("Could not open the document");
            }
            }
            catch (Exception ex)
            {

                Response.Write(ex.StackTrace.ToString() + "==>" + Convert.ToString(ex.Message));
            }


        }

        private void CreateMailMergeDataFile(string fileName)
        {
            //wrdDoc.MailMerge.OpenHeaderSource(Name: fileName);
            wrdDoc.MailMerge.OpenDataSource(Name: fileName, SQLStatement: "SELECT * FROM `Table`", ConfirmConversions: true, Format: Microsoft.Office.Interop.Word.WdOpenFormat.wdOpenFormatAllWord);
            
            //wrdDoc.MailMerge.OpenDataSource(fileName,Word.WdOpenFormat.wdOpenFormatWebPages,true,oMissing,oMissing,oMissing,oMissing,oMissing,true,oMissing,oMissing, oMissing  ,"select * from Table ", oMissing,oMissing,oMissing);

        }

        public static bool CreateExcelFile(DataTable dt, string filename)
        {
            try
            {
                string sTableStart = @"<HTML><BODY><TABLE>";
                string sTableEnd = @"</TABLE></BODY></HTML>";
                string sTHead = "<TR>";
                StringBuilder sTableData = new StringBuilder();
                foreach (DataRow row in dt.Rows)
                {
                    sTHead += @"<TH>" + row[0].ToString().Replace("(", string.Empty).Replace(")", string.Empty).Replace(",", string.Empty) + @"</TH>";
                }
                sTHead += @"</TR>";

                string sTable = sTableStart + sTHead + sTableData.ToString() + sTableEnd;
                System.IO.StreamWriter oDatasourceWriter = System.IO.File.CreateText(filename);
                oDatasourceWriter.WriteLine(sTable);
                oDatasourceWriter.Close();
                oDatasourceWriter = null;
                return true;
            }
            catch
            {
                return false;
            }
        }

        #region . Private functions //TODO: Move to a library Project and use LPS library to fetch data from the database .

        #endregion
    }
}