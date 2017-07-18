using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;
using Word = Microsoft.Office.Interop.Word;
using System.Net;
using System.Text.RegularExpressions;
using System.IO;
using System.Net.Mail;
using System.Collections;
using Ionic.Zip;
using System.Messaging;
using Microsoft.Office.Interop.Word;
using System.Web.Configuration;
using iTextSharp.text.pdf;
using System.Configuration;
using System.Text;
using System.Threading;
using System.Diagnostics;
using System.Web.Services.Description;
using Org.BouncyCastle.Asn1.Microsoft;

namespace MailMerger
{
    public partial class MailMerge : System.Web.UI.Page
    {
        #region . Global Variables
        protected static DatabaseInfo Dbase = new DatabaseInfo();

        private const string QueueName = @".\private$\lpmerge";


        Application _wrdApp;
        Document _wrdDoc;
        Application _wrdApp2;
        Document _wrdDoc2;
        Object _oMissing = System.Reflection.Missing.Value;
        Object _oFalse = false;
        static string _formatPath;
        static string _dataSourcePath;
        bool _isFormatFile;
        bool _isSourceFile;
        static string _datasourceTxt;

        int _totalRecords;


        string _finalDocs = Convert.ToString(WebConfigurationManager.AppSettings["zipFilePath"]);
        string _zipFilePath = Convert.ToString(WebConfigurationManager.AppSettings["MailMergeDocs"]);
        int _recoredSize = Convert.ToInt32(WebConfigurationManager.AppSettings["RecordSize"]);
        DateTime _startTime;
        DateTime _endTime;

        string _email;
        private const string PdfDocument = "PDF document";
        private const string WordDocument = "Word document";

        private string _documentType;
        private bool _cbChecked = true;
        private string _prefix;
        private string _field1;
        private string _field2;
        private string _field3;
        private string _field4;
        private string _field5;
        private string _suffix;
        private string _delimiter;

        #endregion

        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                Response.Clear();
                // FetchFiles();
                _dataSourcePath = "";
                OldReceiveData();
            }
            catch (Exception ex)
            {
                WriteError(ex.Message);
            }
            
        }

        private void FetchFiles()
        {
            try
            {
                var currentDbase = Dbase.dbase(Request.QueryString["database"].Trim());
                //var currentDbase = Dbase.dbase("pension");
                if (currentDbase.zip_path != "")
                    _zipFilePath = currentDbase.merge_working_path;
                if (currentDbase.merge_working_path != "")
                    _finalDocs = currentDbase.zip_path;


                string tempTemplatePath = Convert.ToString(Request.QueryString["format"]);
                string tempSourcePath = Convert.ToString(Request.QueryString["source"]);
               

                #region . Fetching Template file .
                string time = DateTime.Now.ToString("dd-MM-yy-HH-mm-ss");
                Guid objGuid = Guid.NewGuid();
                FetchTemplateFile(tempTemplatePath, objGuid, time);

                #endregion
                WriteError("Template File Fetched. - " + _formatPath);
                #region . Fetching datasource file .

                // If path is http, then download file to local temporary directory.
                WriteError("tempSourcePath. - " + tempSourcePath);
                FetchSourceFiles(tempSourcePath, objGuid, time);

                #endregion
                WriteError("Data Source File Fetched." + _dataSourcePath);
            }
            catch (Exception ex)
            {
                WriteError(ex.Message);
            }
        }

        private void FetchSourceFiles(string tempSourcePath, Guid objGuid, string time)
        {
            if (tempSourcePath.StartsWith("http"))
            {
                var webClient = new WebClient();
                webClient.DownloadFile(new Uri(tempSourcePath), _zipFilePath + "\\datasource_" + objGuid.ToString() + ".txt");
                _dataSourcePath = _zipFilePath + "datasource_" + objGuid.ToString() + ".txt";
                _isSourceFile = true;
            }
            else if (Regex.IsMatch(tempSourcePath.Replace("/", "\\"),
                @"^(?:[a-zA-Z]\:|\\\\[\w\.]+\\[\w.]+)\\(?:[\w]+\\)*\w([\w.])+$"))
            {
                _dataSourcePath = tempSourcePath.Replace("/", "\\");
                if (!File.Exists(_dataSourcePath))
                    throw new Exception("Datasource file does not exists. Please provide correct path");
                if (!Directory.Exists(_zipFilePath))
                    Directory.CreateDirectory(_zipFilePath);
                File.Copy(_dataSourcePath, _zipFilePath + "\\datasource_" + objGuid.ToString() + time + ".doc", true);
                _dataSourcePath = _zipFilePath + "datasource_" + objGuid.ToString() + time + ".doc";
                _isSourceFile = true;
            }
            else if (Regex.IsMatch(tempSourcePath.Replace("/", "\\"),
                @"^(?:[a-zA-Z]\:|\\[\w\.]+\\[\w.]+)\\(?:[\w]+\\)*\w([\w.])+$"))
            {
                _dataSourcePath = tempSourcePath.Replace("\\", "\\\\");
                if (!File.Exists(_dataSourcePath))
                    throw new Exception("Datasource file does not exists. Please provide correct path");
                File.Copy(_dataSourcePath, _zipFilePath + "\\datasource_" + objGuid.ToString() + ".doc",
                    true);
                _dataSourcePath = _zipFilePath + "datasource_" + objGuid.ToString() + ".doc";
                _isSourceFile = true;
                //isSourceFile = true;
            }
            else if (Regex.IsMatch(tempSourcePath,
                @"^([a-zA-Z]\:|\\\\[^\/\\:*?<>|]+\\[^\/\\:*?<>|]+)(\\[^\/\\:*?<>|]+)+(\.[^\/\\:*?<>|]+)$"))
            {
                // Check if file is accessable or present on the remote/network drive.
                if (!File.Exists(tempSourcePath))
                    throw new Exception("File not found or accessable on " + tempSourcePath);
                if (!Directory.Exists(_zipFilePath))
                    Directory.CreateDirectory(_zipFilePath);
                File.Copy(Convert.ToString(Request.QueryString["source"]),
                    _zipFilePath + "datasource_" + objGuid.ToString() + time + ".txt", true);
                _dataSourcePath = _zipFilePath + "datasource_" + objGuid.ToString() + time + ".txt";
                
                _isSourceFile = true;
            }
            _datasourceTxt = _dataSourcePath;
        }
        
        private void FetchTemplateFile(string tempTemplatePath, Guid objGuid, string time)
        {
// If path is http, then download file to local temporary directory.
            if (tempTemplatePath.StartsWith("http"))
            {
                var webClient = new WebClient();
                webClient.DownloadFile(new Uri(tempTemplatePath),
                    _zipFilePath + "\format_" + objGuid.ToString() + ".doc");
                _formatPath = _zipFilePath + "format_" + objGuid.ToString() + time + ".doc";
                _isFormatFile = true;
            }
            else if (Regex.IsMatch(tempTemplatePath.Replace("/", "\\"),
                @"^(?:[a-zA-Z]\:|\\\\[\w\.]+\\[\w.]+)\\(?:[\w]+\\)*\w([\w.])+$"))
            {
                _formatPath = tempTemplatePath.Replace("/", "\\");
                if (!File.Exists(_formatPath))
                    throw new Exception("Template file does not exists. Please provide correct path");
            }
            else if (Regex.IsMatch(tempTemplatePath.Replace("/", "\\"),
                @"^(?:[a-zA-Z]\:|\\[\w\.]+\\[\w.]+)\\(?:[\w]+\\)*\w([\w.])+$"))
            {
                _formatPath = tempTemplatePath.Replace("\\", "\\\\");
                if (!File.Exists(_formatPath))
                    throw new Exception("Template file does not exists. Please provide correct path");
            }
            else if (Regex.IsMatch(tempTemplatePath,
                @"^([a-zA-Z]\:|\\\\[^\/\\:*?<>|]+\\[^\/\\:*?<>|]+)(\\[^\/\\:*?<>|]+)+(\.[^\/\\:*?<>|]+)$"))
            {
                // Check if file is accessable or present on the remote/network drive.
                if (!File.Exists(tempTemplatePath))
                    throw new Exception("File not found or accessable on " + tempTemplatePath);
                try
                {
                    File.Copy(tempTemplatePath, _zipFilePath + "format_" + objGuid.ToString() + time + ".doc", true);

                    _formatPath = _zipFilePath + "format_" + objGuid.ToString() + time + ".doc";
                    WriteError("temp template created - " + _formatPath);
                    _isFormatFile = true;
                    if (!File.Exists(_formatPath))
                        throw new Exception("Error opening source file - " + _formatPath);
                }
                catch (Exception ex)
                {
                    Response.Write("There was a problem copying the Template file; please ensure it is not locked");
                }
            }
        }

        void OldReceiveData()
        {

            //todo this needs to be updated, maybe a global variable
            //var currentDbase = Dbase.dbase("pension");
           
            try
            {
                var currentDbase = Dbase.dbase(Request.QueryString["database"].Trim());
                if (!File.Exists(_dataSourcePath) || !File.Exists(_formatPath))
                {
                    FetchFiles();
                }

                _email = Request.QueryString["email"].Trim();

                if (Request.QueryString["file_type"].Trim() != WordDocument && Request.QueryString["file_type"].Trim() != PdfDocument)
                    _documentType = WordDocument;
                else
                    _documentType = Request.QueryString["file_type"].Trim();

                _cbChecked = Request.QueryString["split"].Trim() == "Yes";
                WriteError(_cbChecked.ToString());
                _prefix = Request.QueryString["prefix"].Trim();
                _field1 = Request.QueryString["field1"].Trim();
                _field2 = Request.QueryString["field2"].Trim();
                _field3 = Request.QueryString["field3"].Trim();
                _field4 = Request.QueryString["field4"].Trim();
                _field5 = Request.QueryString["field5"].Trim();
                _suffix = Request.QueryString["suffix"].Trim();
                _delimiter = Request.QueryString["delimiter"].Trim();

                if (!string.IsNullOrEmpty(_dataSourcePath) && !string.IsNullOrEmpty(_formatPath))
                {
                    #region . Count number of rows .
                    try
                    {
                        using (StreamReader r = new StreamReader(_dataSourcePath))
                        {
                            while (r.ReadLine() != null)
                            {
                                _totalRecords++;
                            }
                            WriteError("Data Source File Read");
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteError(ex.Message);
                    }
                    _totalRecords -= 1;

                    #endregion


                    if (_totalRecords > _recoredSize)
                    {
                        if (!string.IsNullOrEmpty(_email))
                        {
                            SendtoQueue(currentDbase);
                            WriteError("Your request is added to the queue. You will be notified when merge process is finished.");
                        }
                        else
                        {
                            WriteError("Please provide email address to receive the mail merge process completion notification.");
                            Response.Write("Please provide email address to receive the mail merge process completion notification.");
                        }
                    }
                    else
                    {
                        MergeMail();
                        try
                        {
                            if (_isFormatFile && File.Exists(_formatPath)) File.Delete(_formatPath);
                            if (_isSourceFile && File.Exists(_dataSourcePath)) File.Delete(_dataSourcePath);
                            if (_isSourceFile && File.Exists(_datasourceTxt)) File.Delete(_datasourceTxt);
                        }
                        catch(Exception ex)
                        {
                            WriteError(ex.Message);
                        }
                    }
                    //endTime = DateTime.Now;LPS-318 aborted exception
                    //TimeSpan tSpend = endTime - startTime;
                    //  this.sendemail(tSpend.TotalSeconds);
                }
                else Response.Write("Incorrect file locations.");
            }
            catch (WebException)
            {
                #region . Exception Checks .
                
                Response.Write("Unable to download the file from the remote server.");


                #endregion
            }
            catch (UnauthorizedAccessException uaeEx)
            {
                Response.Write(uaeEx.StackTrace);
            }
            catch (Exception ex)
            {
                Response.Write(ex.StackTrace);
            }
        }

        private void SendtoQueue(DatabaseInfo currentDbase)
        {
            try
            {
             DatabaseInfo.AddtoMailMergeQueue(0,
                    _formatPath + ";" + _dataSourcePath + ";" + _totalRecords + ";" +
                    _email + ";" + _documentType + ";"
                    + _cbChecked + ";" + _prefix + ";" + _field1 +
                    ";" + _field2 + ";" + _field3 + ";"
                    + _field4 + ";" + _field5 + ";" + _suffix + ";" + _delimiter + ";" +
                    currentDbase.database_name,
                    "",
                    "",
                    currentDbase, _totalRecords);

                SendMessageToQueue(_formatPath + ";" + _dataSourcePath + ";" + _totalRecords + ";" +
                                   _email + ";" + _documentType + ";"
                                   + _cbChecked + ";" + _prefix + ";" + _field1 +
                                   ";" + _field2 + ";" + _field3 + ";"
                                   + _field4 + ";" + _field5 + ";" + _suffix + ";" + _delimiter + ";" +
                                   currentDbase.database_name);
            }
            catch (Exception ex)
            {
                WriteError(ex.Message);
            }
            Response.Write(
                "Your request is added to the queue. You will be notified when merge process is finished.");
        }

        private void MergeMail()
        {
            _startTime = DateTime.Now;
            var array = new List<string>();
            var headerRows = new List<string>();
            Word.MailMerge wrdMailMerge;

            string finalFilePath = "";
            // string time = "";
            int counter = 1;
            _recoredSize = _cbChecked ? 1 : Convert.ToInt32(ConfigurationManager.AppSettings["RecordSize"]);

            //if (_documentType == pdfDocument)
            //    fileNamePrefix = "Pdf-";
            //else if (_documentType == wordDocument)
            //    fileNamePrefix = "Doc-";            

            try
            {
                // Create an instance of Word and make it visible.
                // Add a new document.
                string path = _formatPath;
                string newPath = "";
                var delimitersList = new List<string>
                { "|",
                                                          "\r\n",
                                                          "\t",
                                                          ",",
                                                          ".",
                                                          "!",
                                                          "#",
                                                          "$",
                                                          "%",
                                                          "&",
                                                          "(",
                                                          ")",
                                                          "+",
                                                          "*",                                                         
                                                          "/",
                                                          ":",
                                                          ";",
                                                          "<",
                                                          "=",
                                                          ">",
                                                          "?",
                                                          "@",
                                                          "[",
                                                          "]",
                                                          "^",                                                         
                                                          "`",
                                                          "{",
                                                          "}",                                                         
                                                           "-",
                                                            "_",
                                                          "~"
                                                      };

                #region . clean the data .
                WriteError("Open Data Source File to read.");

                string data = File.ReadAllText(_datasourceTxt,Encoding.Default);
                ////int count = data.Split();

                File.WriteAllText(_datasourceTxt, data.Trim(), Encoding.Default);
                WriteError("Close Data Source File.");

                string[] allLines = File.ReadAllLines(_datasourceTxt);
                #endregion

                #region

                if (_cbChecked)
                {
                    string dataSourceCopyPath = _zipFilePath + "datasource_copy.txt";
                    TextWriter writer = new StreamWriter(dataSourceCopyPath);
                    //  string sourceData = "";
                    //   string[] allLines = File.ReadAllLines(dataSourcePath);
                    for (int i = 0; i < allLines.Length; i++)
                    {
                        if (i > 10)
                            break;
                        writer.WriteLine(allLines[i]);
                    }

                    writer.Flush();
                    writer.Close();
                    array.Add(dataSourceCopyPath);
                    WriteError("Data Source Added to array of lines.");
                }

                #endregion

                #region Check datasouce delimiters

                try
                {
                   

                    if (allLines.Length >= 2)
                    {
                        WriteError("Record length greather than 2");
                        string fieldLine = allLines[0];
                        bool delPresentFlag = false;
                        foreach (string t in delimitersList)
                        {
                            if (fieldLine.Contains(t))
                            {
                                delPresentFlag = true;
                                headerRows.AddRange(fieldLine.Split(t.ToCharArray()[0]).ToArray());
                                break;
                            }
                        }

                        if (!delPresentFlag)
                        { 
                            WriteError("Header delimiter missing in datasource.");
                            throw new Exception("Header delimiter missing in datasource.");
                        }

                        fieldLine = allLines[1];
                        delPresentFlag = delimitersList.Any(t => fieldLine.Contains(t));

                        if (!delPresentFlag)
                        {
                            WriteError("Field delimiter missing in datasource.");
                            throw new Exception("Field delimiter missing in datasource.");
                        }
                    }
                }
                catch
                {
                    WriteError("Error in checking delimiter.");
                }
                WriteError("Checking delimiter complete.");
                #endregion

                const int startingRecord = 1;
                bool fieldsVerified = false;
                int endingRecord = _totalRecords > _recoredSize ? _recoredSize : _totalRecords;
                object oMissing = System.Reflection.Missing.Value;

                WriteError("Number of Records" + _totalRecords);
                while (endingRecord <= _totalRecords && _totalRecords > 0)
                {
                    string fileNamePrefix = "";

                   
                    oMissing = OMissing(oMissing);

                    #region Initiate New Word Application

                    _wrdApp = new Application();
                   _wrdApp.Application.Visible = true;
                    _wrdApp.Options.SaveNormalPrompt = false;
                    _wrdApp.Options.SavePropertiesPrompt = false;
                    _wrdApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;

                    _wrdDoc = _wrdApp.Documents.Open(path);

                    WriteError("File Opened.");
                    if (_wrdDoc != null)
                    {
                        try
                        {
                            

                            #region . Generate file with RecoredSize memebers size .

                            _wrdDoc.Select();
                            WriteError("Doc Selected.");
                            wrdMailMerge = _wrdDoc.MailMerge;

                            _wrdDoc.Application.DisplayAlerts = WdAlertLevel.wdAlertsNone;


                            if (!fieldsVerified)
                            {
                              //  MailMergeFieldNames source = wrdMailMerge.DataSource.FieldNames;
                                MailMergeFields template = wrdMailMerge.Fields;
                                var sourceList = new List<string>();
                                var templateList = new List<string>();
                                foreach (string head in headerRows)
                                {
                                    string headerName;
                                    if (head.Contains('#') || head.Contains('~'))
                                        headerName = head.Replace("#", "").Replace("~", "");
                                    else
                                        headerName = head.ToLower();

                                    sourceList.Add(headerName.ToLower());
                                }
                                
                               
                                WriteError(sourceList.Count + " items in source");
                                foreach (MailMergeField mailMergeField in template)
                                {
                                    string item = mailMergeField.Code.Text.Replace("MERGEFIELD ", string.Empty).Trim();
                                    //item = item.Replace("mergefield", string.Empty).Trim();
                                    if (!item.StartsWith("IF"))
                                    {
                                        if (item.Contains("\\"))
                                        {
                                            string[] parts = item.Split(new string[] { "\\" },
                                                                        StringSplitOptions.RemoveEmptyEntries);
                                            if (parts.Length > 0)
                                                item = parts[0].TrimEnd(' ');
                                        }
                                        if (item.Contains("\""))
                                            item = item.Replace("\"", "");
                                        templateList.Add(item.ToLower());
                                    }
                                }
                                WriteError(templateList.Count + " fields in template");
                                bool fieldException = false;
                                var fieldExceptions = new List<string>();

                                for (int i = 0; i < templateList.Count; i++)
                                {
                                   
                                    if (!sourceList.Contains(templateList[i]))
                                    {
                                        fieldExceptions.Add(templateList[i]);
                                        fieldException = true;
                                    }
                                    
                                }
                                if (fieldException)
                                {
                                    _wrdDoc.Close(ref _oFalse, ref _oMissing, ref _oMissing);
                                    _wrdDoc = null;
                                    WriteError("Merge fields used in the template document does not exist in the data source. Your file could not be created. ");
                                    var exceptionlist = "<br/>";
                                    foreach (string field in fieldExceptions)
                                    {
                                        exceptionlist = exceptionlist + field + "<br/>";
                                    }

                                    throw new Exception("Merge fields used in the template document listed do not exist in the data source." + exceptionlist + "Your file couldn't be created");
                                }

                                fieldsVerified = true;
                            }

                            //_dataSourcePath =
                                //"G:\\Dropbox\\Project Folder\\Together-lps-576c7ec042f6\\MailMerger\\MailMerger\\MailMergeDocs\\1.doc";
                            //wrdMailMerge.OpenDataSource(_dataSourcePath, WdOpenFormat.wdOpenFormatAuto);


                            wrdMailMerge.OpenDataSource(_dataSourcePath);

                            _wrdDoc.MailMerge.SuppressBlankLines = true;
                            wrdMailMerge.DataSource.FirstRecord = startingRecord;
                            wrdMailMerge.DataSource.LastRecord = endingRecord;
                            _wrdApp.Options.CheckSpellingAsYouType = false;
                            _wrdApp.Options.CheckGrammarAsYouType = false;
                            _wrdApp.Options.SuggestSpellingCorrections = false;
                            _wrdApp.Options.CheckGrammarWithSpelling = false;
                            _wrdApp.ActiveDocument.ShowSpellingErrors = false;
                            _wrdDoc.ShowSpellingErrors = false;
                            _wrdDoc.ShowGrammaticalErrors = false;
                            _wrdDoc.SpellingChecked = true;
                            _wrdDoc.GrammarChecked = true;
                            // Perform mail merge.
                            try
                            {
                                WriteError("Merge Ready to execute");
                                wrdMailMerge.Destination = WdMailMergeDestination.wdSendToNewDocument;
                                wrdMailMerge.Execute();
                                WriteError("Merge Executed.");
                            

                            }
                            catch(Exception ex)
                            {
                                WriteError(ex.StackTrace);
                                throw new Exception("There was a problem Merging the document");
                                CleanUp();
                                return;
                            }

                            var fields = new Hashtable();

                            foreach (MailMergeDataField item in _wrdDoc.MailMerge.DataSource.DataFields)
                            {
                                if (item.Name == _field1 || item.Name == _field2 || item.Name == _field3 ||
                                    item.Name == _field4 || item.Name == _field5)
                                    fields.Add(item.Name, item.Value);
                            }
                            WriteError("Fields Checked.");
                            if (_cbChecked)
                            {
                                if (!String.IsNullOrEmpty(_prefix))
                                    fileNamePrefix += _prefix;

                                if (!string.IsNullOrEmpty(_field1))
                                    fileNamePrefix += fields.ContainsKey(_field1)
                                                          ? _delimiter + fields[_field1]
                                                          : string.Empty;

                                if (!string.IsNullOrEmpty(_field2))
                                    fileNamePrefix += fields.ContainsKey(_field2)
                                                          ? _delimiter + fields[_field2]
                                                          : string.Empty;
                                if (!string.IsNullOrEmpty(_field3))
                                    fileNamePrefix += fields.ContainsKey(_field3)
                                                          ? _delimiter + fields[_field3]
                                                          : string.Empty;
                                if (!string.IsNullOrEmpty(_field4))
                                    fileNamePrefix += fields.ContainsKey(_field4)
                                                          ? _delimiter + fields[_field4]
                                                          : string.Empty;
                                if (!string.IsNullOrEmpty(_field5))
                                    fileNamePrefix += fields.ContainsKey(_field5)
                                                          ? _delimiter + fields[_field5]
                                                          : string.Empty;


                                if (!string.IsNullOrEmpty(_suffix))
                                    fileNamePrefix += _delimiter + _suffix;
                                fileNamePrefix += _delimiter;

                                WriteError("Prefix and delimiters checked.");
                            }

                            char[] invalidFileNameChars = Path.GetInvalidFileNameChars();

                            for (int i = 0; i < invalidFileNameChars.Length; i++)
                            {
                                if (fileNamePrefix.Contains(invalidFileNameChars[i].ToString(CultureInfo.InvariantCulture)))
                                {
                                    if (_cbChecked)
                                        fileNamePrefix = fileNamePrefix.Replace(invalidFileNameChars[i].ToString(CultureInfo.InvariantCulture),
                                                                                _delimiter);
                                    else
                                        fileNamePrefix = fileNamePrefix.Replace(invalidFileNameChars[i].ToString(), "_");
                                }

                            }
                            WriteError("Invalid characters checked.");

                            if (_documentType == PdfDocument)
                            {
                                if (fileNamePrefix == "")
                                    fileNamePrefix = "PDF";
                                newPath = _zipFilePath + fileNamePrefix + counter + ".pdf";
                            }
                            else if (_documentType == WordDocument)
                            {
                                if (fileNamePrefix == "")
                                    fileNamePrefix = "Word";
                                newPath = _zipFilePath + fileNamePrefix + counter + ".doc";
                            }

                            array.Add(newPath);
                            WriteError("New File path created.");
                            _wrdDoc.Close(ref _oFalse, ref _oMissing, ref _oMissing);
                            WriteError("Doc Closed.");
                            foreach (_Document item in _wrdApp.Documents)
                            {
                                item.Activate();
                                _wrdApp.ActiveDocument.SpellingChecked = true;
                                _wrdApp.ActiveDocument.GrammarChecked = true;
                                _wrdApp.ActiveDocument.ShowSpellingErrors = false;

                                if (_documentType == PdfDocument)
                                    item.ExportAsFixedFormat(newPath, Word.WdExportFormat.wdExportFormatPDF, false,
                                                             WdExportOptimizeFor.wdExportOptimizeForPrint);
                                else if (_documentType == WordDocument)
                                    item.SaveAs(newPath, Word.WdSaveFormat.wdFormatDocument);
                                WriteError("new Path - " + newPath);
                                item.Close(ref _oFalse, ref _oMissing, ref _oMissing);
                            }

                            #endregion


                            #region

                            if (_documentType == PdfDocument)
                            {
                                byte[] b = File.ReadAllBytes(newPath);
                                var reader = new PdfReader(b);

                                using (var m = new MemoryStream())
                                {
                                    PdfStamper stamper = new PdfStamper(
                                        reader, m, PdfWriter.VERSION_1_5
                                        );

                                    stamper.Writer.CompressionLevel = PdfStream.BEST_COMPRESSION;

                                    int total = reader.NumberOfPages + 1;
                                    for (int i = 1; i < total; i++)
                                    {
                                        reader.SetPageContent(i, reader.GetPageContent(i));
                                    }

                                    stamper.SetFullCompression();
                                    stamper.Close();

                                    File.WriteAllBytes(newPath, m.ToArray());
                                }
                                WriteError("PDF format created.");
                            }

                            #endregion

                            //if (endingRecord == totalRecords) break;

                            #region . Remove Rows .
                            WriteError("datasource_text -" + _datasourceTxt);
                            List<string> quotelist = File.ReadAllLines(_datasourceTxt, Encoding.Default).ToList();
                            string firstItem = quotelist[0];
                            //  quotelist.RemoveRange(1, 1);
                            if (_recoredSize < quotelist.Count)
                                quotelist.RemoveRange(1, _recoredSize);
                            else
                                quotelist.RemoveRange(1, quotelist.Count - 1);


                            _totalRecords = quotelist.Count - 1;

                            File.WriteAllLines(_datasourceTxt, quotelist.ToArray(), Encoding.Default);
                           
                            WriteError("Source Line Removed.");
                            #endregion
                            if (_totalRecords > _recoredSize)
                                endingRecord = _recoredSize;
                            else endingRecord = _totalRecords;
                            //if ((endingRecord + RecoredSize) >= totalRecords) endingRecord = totalRecords;
                        }
                        catch (Exception docEx)
                        {
                            WriteError(docEx.Message);
                            //sendEmail(email, docEx.Message + "\n\n" + docEx.StackTrace.ToString());
                            SendEmail(_email, docEx.Message);
                            throw docEx;
                        }
                    }
                    //two
                    else
                    {
                        WriteError("Document cannot be opened");
                        throw new Exception("Document cannot be opened");
                    }

                    #region Close the word Applicatoin

                    CleanUp();

                    #endregion

                    counter++;
                }

                #region . Throw file to the client .

                if (!_cbChecked)
                {
                    WriteError("Download File Called.");
                    FileStream fStream = new FileStream(newPath, FileMode.Open, FileAccess.Read);
                    long fileSize = fStream.Length;

                    byte[] buffer = new byte[(int)fileSize];
                    fStream.Read(buffer, 0, (int)fileSize);
                    fStream.Close();

                    Response.Buffer = true;
                    Response.Clear();
                    if (_documentType == PdfDocument)
                    {
                        Response.ContentType = "application/pdf";
                        Response.AddHeader("content-disposition", "attachment; filename=Report.pdf");
                    }
                    else if (_documentType == WordDocument)
                    {
                        Response.ContentType = "application/msword";
                        Response.AddHeader("content-disposition", "attachment; filename=Report.doc");
                    }
                    WriteError("File Download Start.");
                    Response.BinaryWrite(buffer);
                    WriteError("Response.BinaryWrite(buffer)");
                    HttpContext.Current.Response.Flush(); // Sends all currently buffered output to the client.
                    HttpContext.Current.Response.SuppressContent = true;  // Gets or sets a value indicating whether to send HTTP content to the client.
                    HttpContext.Current.ApplicationInstance.CompleteRequest(); // Causes ASP.NET to bypass all events and filtering in the HTTP pipeline chain of execution and directly execute the EndRequest event.
                    //HttpContext.Current.Response.End();
                    WriteError("File Downloaded.");
                }
                else if (_cbChecked)
                {
                    using (var zip = new ZipFile())
                    {
                        Guid objGuid = Guid.NewGuid();
                        WriteError("Zip File Creation Called.");
                        zip.AddFiles(array, "");
                        string time = DateTime.Now.ToString("dd-MM-yy-HH-mm-ss");
                        time = time.Replace("-", _delimiter);
                        finalFilePath = _finalDocs + "MailMerge_" + _email  + "_" + time + ".zip";
                        zip.UseZip64WhenSaving = Zip64Option.Always;
                        zip.Save(_zipFilePath + "MailMergedZip_" + objGuid + ".zip");
                        File.Move(_zipFilePath + "MailMergedZip_" + objGuid + ".zip", finalFilePath);
                        
                        WriteError("Zip File Created and Saved");
                    }
                    foreach (var item in array)
                    {
                        try
                        {
                            File.Delete(item);
                        }
                        catch
                        {
                            WriteError("Delete file error");
                            // Do not catch
                        }
                    }
                }

                #endregion


                _endTime = DateTime.Now;
                var tSpend = _endTime - _startTime;
                if (_cbChecked)
                {
                    if (!string.IsNullOrEmpty(_email))
                        sendemail(_email, finalFilePath, tSpend.TotalSeconds);
                }
                Response.Write("Mail Merge has completed, please check email for link to file");
                WriteError("Please check the provided email inbox for merge results");
            }
            //fix for LPS-318
            catch (System.Threading.ThreadAbortException)
            {
                WriteError("LPS-318 aborted exception");
                //do nothing
                //catch Thread was aborted exception
            }
            catch (System.Runtime.InteropServices.COMException comException)
            {
                WriteError("Com Exception");
                Response.Write(comException.Message);
                //do nothing
                //Catch command failed exception while opening a word document
            }
            catch (Exception oEx)
            {
               
                Response.Write(oEx.Message);
                SendEmail(_email, oEx.Message + "\n\n" + oEx.StackTrace);
                
            }
            finally
            {
                try
                {
                    // Release References.
                    wrdMailMerge = null;
                    CleanUp();
                    _wrdDoc = null;
                    _wrdApp.Quit(ref _oMissing, ref _oMissing, ref _oMissing);
                    _wrdApp = null;
                    GC.Collect();
                }
                catch
                {
                    WriteError("Close off session");
                }
            }
        }

        private object OMissing(object oMissing)
        {
            try
            {
                string AppId = _formatPath;


                _wrdApp2 = new Application();
                _wrdApp2.Application.Caption = AppId;

                _wrdApp2.Application.Visible = true;

                ///Get the pid by for word application
                //var WordPid = GetProcessIdByWindowTitle(AppId);

                //while (GetProcessIdByWindowTitle(AppId) == Int32.MaxValue) //Loop till u get
                //{
                //    Thread.Sleep(5);
                //}

                //WordPid = GetProcessIdByWindowTitle(AppId);
                //WriteError("PiD - " + WordPid);

                ///You can hide the application afterward            
                //_wrdApp.Application.Visible = false;


                _wrdApp2.Options.SaveNormalPrompt = false;
                _wrdApp2.Options.SavePropertiesPrompt = false;

                _wrdApp2.DisplayAlerts = WdAlertLevel.wdAlertsNone;

                // Create an instance of Word and make it visible.
                //_wrdApp.Visible = true;
                WriteError("Initiate New Word Application.");
            }
            catch
            {
                WriteError("Failed to initialize MSWord. Check the permission on server.");
                throw new Exception("Failed to initialize MSWord. Check the permission on server.");
            }

            #endregion

            //path = 
            //"G:\\Dropbox\\Project Folder\\Together-lps-576c7ec042f6\\MailMerger\\MailMerger\\MailMergeDocs\\1.doc";


            _wrdDoc2 = _wrdApp2.Documents.Open(_datasourceTxt, false, true, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, WdOpenFormat.wdOpenFormatText, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            

            _dataSourcePath = _datasourceTxt.Substring(0, _datasourceTxt.Length - 3) + "doc";
            try
            {
                if (File.Exists(_dataSourcePath)) File.Delete(_dataSourcePath);
            }
            catch (Exception e)
            {
                WriteError("Error Trying to Delete Source File Doc");
            }

            WriteError(_dataSourcePath);
            _wrdDoc2.SaveAs(_dataSourcePath, WdSaveFormat.wdFormatDocument, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            _wrdDoc2.Close();
            try
            {
                if (_wrdDoc2 != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_wrdDoc2);
                _wrdDoc2 = null;
                if (_wrdApp2 != null)
                {
                    _wrdApp2.Quit(ref _oMissing, ref _oMissing, ref _oMissing);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_wrdApp2);
                }
                _wrdApp2 = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return oMissing;
        }

        public void CleanUp()
        {
            try
            {
                if (_wrdDoc != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_wrdDoc);
                _wrdDoc = null;
                if (_wrdApp != null)
                {
                    _wrdApp.Quit(ref _oMissing, ref _oMissing, ref _oMissing);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_wrdApp);
                }
                _wrdApp = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void SendEmail(string email, string errMessage)
        {
            WriteError("Error email called" + errMessage);
            if (ConfigurationManager.AppSettings["SendEmail"] == "1")
            {

                var message = new System.Net.Mail.MailMessage();
                string emailFrom = WebConfigurationManager.AppSettings["emailFrom"];
                message.From = new MailAddress(emailFrom);
                message.To.Add(email);
                string emails = WebConfigurationManager.AppSettings["emailTo"];
                foreach (var item in emails.Split(",".ToCharArray()))
                {
                    message.Bcc.Add(item);
                }
                message.IsBodyHtml = true;
                string pathFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ErrorHeader.txt");
                StreamReader objStreamReader = File.OpenText(pathFile);
                string subject = objStreamReader.ReadLine();
                message.Subject = subject;
                string content = ReadFile("ErrorTemplate.htm");
                message.Body = content.Replace("#errorMessage#", errMessage);
                var client = new SmtpClient();
                client.Send(message);
                Response.Write(errMessage);
                WriteError("Error Email send");
            }
        }

        private void sendemail(string email, string path, double time)
        {
            WriteError("File Email called" + path);
            if (ConfigurationManager.AppSettings["SendEmail"] == "1")
            {

                var message = new System.Net.Mail.MailMessage();
                string emailFrom = WebConfigurationManager.AppSettings["emailFrom"];
                message.From = new MailAddress(emailFrom);
                string emails = WebConfigurationManager.AppSettings["emailTo"];
                message.To.Add(email);
                foreach (var item in emails.Split(",".ToCharArray()))
                {
                    message.Bcc.Add(item);
                }

                message.IsBodyHtml = true;
                string pathFile = Path.Combine(HttpContext.Current.Server.MapPath("EmailHeader.txt"));
                StreamReader objStreamReader = File.OpenText(pathFile);
                string subject = objStreamReader.ReadLine();
                message.Subject = subject;
                string content = ReadFile("EmailTemplate.htm");
                message.Body = content.Replace("#path#", path).Replace("#ExecutionTime#", time.ToString());
                var client = new SmtpClient();
                client.Send(message);
                WriteError("Email Send");
            }
        }

        private void SendMessageToQueue(string message)
        {
            WriteError("Send Message To Queue");
            // if (ConfigurationManager.AppSettings["SendEmail"].ToString() == "1")
            //   {
            // check if queue exists, if not create it
            MessageQueue msMq;
            if (!MessageQueue.Exists(QueueName))
            {
                msMq = MessageQueue.Create(QueueName);
                msMq.SetPermissions("Everyone", MessageQueueAccessRights.FullControl);
            }
            else
            {
                msMq = new MessageQueue(QueueName);
                msMq.SetPermissions("Everyone", MessageQueueAccessRights.FullControl);
            }

            msMq.Send(message);
            WriteError("Send Message To Queue complete");
            //   }
        }

        public static string ReadFile(string fileName)
        {
            try
            {
                String filename = HttpContext.Current.Server.MapPath(fileName);
                StreamReader objStreamReader = File.OpenText(filename);
                String contents = objStreamReader.ReadToEnd();
                return contents;
            }
// ReSharper disable once EmptyGeneralCatchClause
            catch (Exception)
            {

            }
            return "";
        }


        public static int GetProcessIdByWindowTitle(string AppId)
        {
            Process[] P_CESSES = Process.GetProcesses();
            for (int p_count = 0; p_count < P_CESSES.Length; p_count++)
            {
                if (P_CESSES[p_count].MainWindowTitle.Equals(AppId))
                {
                    return P_CESSES[p_count].Id;
                }
            }

            return Int32.MaxValue;
        }


    }
}