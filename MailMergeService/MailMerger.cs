using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using System.IO;
using Ionic.Zip;
using System.Messaging;
using System.Net.Mail;
using System.Configuration;
using Word = Microsoft.Office.Interop.Word;
using iTextSharp.text.pdf;
using System.Collections;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;

using System.Net;
using System.Text.RegularExpressions;
using System.IO;
using System.Net.Mail;
using System.Collections;
using Ionic.Zip;
using System.Messaging;
using Microsoft.Office.Interop.Word;

using iTextSharp.text.pdf;
using System.Configuration;
using System.Text;
using System.Threading;
using System.Diagnostics;

using Org.BouncyCastle.Asn1.Microsoft;


namespace MailMergeService
{
    public partial class MailMerger
    {
        public bool Running { get; set; }
        string queueName = @".\private$\lpmerge";

        protected static DatabaseInfo Dbase = new DatabaseInfo();

        Microsoft.Office.Interop.Word.Application wrdApp;
        Microsoft.Office.Interop.Word.Document wrdDoc;

        Microsoft.Office.Interop.Word.Application wrdApp2;
        Microsoft.Office.Interop.Word.Document wrdDoc2;
        Object oMissing = System.Reflection.Missing.Value;
        Object oFalse = false;
        private Object oTrue = true;
        string FinalDocs = Convert.ToString(ConfigurationManager.AppSettings["zipFilePath"]);
        string zipFilePath = Convert.ToString(ConfigurationManager.AppSettings["MailMergeDocs"]);
        int RecoredSize = Convert.ToInt32(ConfigurationManager.AppSettings["RecordSize"]);
        static string _datasourceTxt;

        DateTime startTime;
        DateTime endTime;

        string tempTemplatePath;
        string tempSourcePath;
        string pdfDocument = "PDF document";
        string wordDocument = "Word document";

        #region Singleton

        private static MailMerger _merger;
        public static MailMerger Merger
        {
            get
            {
                if (_merger == null)
                {
                    _merger = new MailMerger();

                } return _merger;

            }
        }
        private void MergeMail()
        {
        }
        #endregion


        public void MergeMailWithMultipleRecords(string formatPath, string dataSourcePath, int totalRecords, string email, string _documentType, bool _cbChecked, string _prefix
            , string _field1, string _field2, string _field3, string _field4, string _field5, string _suffix, string _delimiter,string dbase)
        {

            var currentDbase = Dbase.dbase(dbase);
            string FinalDocs = currentDbase.zip_path;
            string zipFilePath = currentDbase.merge_working_path;
            _datasourceTxt = dataSourcePath;
            tempSourcePath = dataSourcePath;
            this.Log("Started:" + formatPath + ";" + dataSourcePath + ";" + totalRecords + ";" + email + ";" + _documentType, EventLogEntryType.Information);
            Running = true;
            startTime = DateTime.Now;
            List<string> array = new List<string>();
            List<string> HeaderRows = new List<string>();
            Selection wrdSelection;
            Word.MailMerge wrdMailMerge;

            object SaveChanges = false;
            string finalFilePath = "";
            int startingRecord = 1, endingRecord = 1;
           // string time = "";
            int counter = 1;

            if (_cbChecked)
                RecoredSize = 1;
            else
                RecoredSize = Convert.ToInt32(ConfigurationManager.AppSettings["RecordSize"]);

            string fileNamePrefix = "";
    

            try
            {
                // Create an instance of Word and make it visible.
                // Add a new document.
                string path = formatPath;
                string newPath = "";


                #region . clean the data .
                writeError("Open Data Source File to read.");
                string data = File.ReadAllText(_datasourceTxt ,Encoding.Default );
                //int count = data.Split();

                File.WriteAllText(dataSourcePath, data.Trim(), Encoding.Default);
                writeError("Close Data Source File.");

                string[] allLines = File.ReadAllLines(_datasourceTxt);
                #endregion

                #region

            //    if (_cbChecked)
                {
                    string dataSourceCopyPath = zipFilePath + "datasource_copy.txt";
                    TextWriter writer = new StreamWriter(dataSourceCopyPath);
                    //string sourceData = "";
                    //string[] allLines = File.ReadAllLines(dataSourcePath);
                    for (int i = 0; i < allLines.Length; i++)
                    {
                        if (i > 10)
                            break;
                        writer.WriteLine(allLines[i]);
                    }

                    writer.Flush();
                    writer.Close();
                    array.Add(dataSourceCopyPath);
                }

                #endregion

                #region Check datasouce delimiters

                try
                {
                    List<string> delimitersList = new List<string>()
                                                      {"|",
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
                
                    if (allLines.Length >= 2)
                    {
                        string fieldLine = allLines[0];
                        bool delPresentFlag = false;
                        for (int i = 0; i < delimitersList.Count; i++)
                        {
                            if (fieldLine.Contains(delimitersList[i]))
                            {
                                delPresentFlag = true;
                                HeaderRows.AddRange(fieldLine.Split(delimitersList[i].ToCharArray()[0]).ToArray());
                                break;
                            }
                        }

                        if (!delPresentFlag)
                        {
                            writeError("Header delimiter missing in datasource.");
                            throw new Exception("Header delimiter missing in datasource.");
                        }

                        fieldLine = allLines[1];
                        delPresentFlag = false;
                        for (int i = 0; i < delimitersList.Count; i++)
                        {
                            if (fieldLine.Contains(delimitersList[i]))
                            {
                                delPresentFlag = true;
                                break;
                            }
                        }

                        if (!delPresentFlag)
                        {
                            writeError("Field delimiter missing in datasource.");
                            throw new Exception("Field delimiter missing in datasource.");
                        }
                    }
                }
                catch
                {
                    writeError("Error in checking delimiter.");
                }

                #endregion

                startingRecord = 1;
                bool fieldsVerified = false;
                if (totalRecords > RecoredSize) 
                    endingRecord = RecoredSize;
                else 
                    endingRecord = totalRecords;

                while (endingRecord <= totalRecords && totalRecords>0)
                {
                    fileNamePrefix = "";

                   

                    var oMissing = OMissing(ref dataSourcePath);

                    try
                    {
                        wrdApp = new Application();
                        wrdApp.Options.SaveNormalPrompt = false;
                        wrdApp.Options.SavePropertiesPrompt = false;
                        wrdApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                        // Create an instance of Word and make it visible.
                        wrdApp.Visible = false;
                        writeError("Initiate New Word Application.");
                    }
                    catch
                    {
                        writeError("Failed to initialize MSWord. Check the permision on server.");
                        throw new Exception("Failed to initialize MSWord. Check the permision on server.");
                    }

                    wrdDoc = wrdApp.Documents.Open(path, ReadOnly: true, OpenAndRepair: true);

                    writeError("File Opened.");
                    if (wrdDoc != null)
                    {
                        try
                        {
                            #region . Generate file with RecoredSize memebers size .

                            wrdDoc.Select();
                            writeError("Doc Selected.");
                            wrdMailMerge = wrdDoc.MailMerge;
                            wrdDoc.Application.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                            // Create a MailMerge Data file.

                            if (!fieldsVerified)
                            {
                                //  MailMergeFieldNames source = wrdMailMerge.DataSource.FieldNames;
                                MailMergeFields template = wrdMailMerge.Fields;
                                List<string> sourceList = new List<string>();
                                List<string> templateList = new List<string>();

                                string HeaderName;
                                foreach (string head in HeaderRows)
                                {
                                    if (head.Contains('#') || head.Contains('~'))
                                        HeaderName = head.Replace("#", "").Replace("~", "");
                                    else
                                        HeaderName = head.ToLower();

                                    sourceList.Add(HeaderName.ToLower());
                                }


                                writeError(sourceList.Count + " items in source");
                                foreach (MailMergeField mailMergeField in template)
                                {
                                    string item =
                                        mailMergeField.Code.Text.Replace("MERGEFIELD ", string.Empty).Trim();
                                    // item = item.TrimStart(' ');
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
                                writeError(templateList.Count + " fields in template");
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
                                    wrdDoc.Close(ref oFalse, ref oMissing, ref oMissing);
                                    wrdDoc = null;
                                    writeError("Merge fields used in the template document does not exist in the data source. Your file could not be created. ");
                                    var exceptionlist = "<br/>";
                                    foreach (string field in fieldExceptions)
                                    {
                                        exceptionlist = exceptionlist + field + "<br/>";
                                    }

                                    throw new Exception("Merge fields used in the template document listed do not exist in the data source." + exceptionlist + "Your file couldn't be created");
                                }

                                fieldsVerified = true;
                            }


                            wrdMailMerge.OpenDataSource(dataSourcePath);


                            if (!fieldsVerified)
                            {
                                MailMergeFieldNames source = wrdMailMerge.DataSource.FieldNames;
                                MailMergeFields template = wrdMailMerge.Fields;
                                List<string> sourceList = new List<string>();
                                List<string> templateList = new List<string>();
                                
                                foreach (MailMergeFieldName mailMergeFieldName in source)
                                {
                                    sourceList.Add(mailMergeFieldName.Name.ToLower());
                                }
                                writeError(sourceList.Count + " items in source");
                                foreach (MailMergeField mailMergeField in template)
                                {
                                    string item =
                                        mailMergeField.Code.Text.Replace("MERGEFIELD ", string.Empty).TrimEnd(' ');
                                    item = item.TrimStart(' ');
                                    if (!item.StartsWith("IF"))
                                    {
                                        if (item.Contains("\\"))
                                        {
                                            string[] parts = item.Split(new string[] {"\\"},StringSplitOptions.RemoveEmptyEntries);
                                            if (parts.Length > 0)
                                                item = parts[0].TrimEnd(' ');
                                        }
                                        if (item.Contains("\""))
                                            item = item.Replace("\"", "");
                                        templateList.Add(item.ToLower());
                                    }
                                }
                                writeError(templateList.Count + " fields in template");
                                for (int i = 0; i < templateList.Count; i++)
                                {
                                    var fieldException = false;
                                    var fieldExceptions = new List<string>();
                                    if (!sourceList.Contains(templateList[i]))
                                    {
                                        fieldExceptions.Add(templateList[i]);
                                        fieldException = true;
                                    }
                                    if (!fieldException) continue;

                                    wrdDoc.Close(ref oFalse, ref oMissing, ref oMissing);
                                    wrdDoc = null;
                                    var exceptionlist = "<br/>";

                                    foreach (string field in fieldExceptions)
                                    {
                                        exceptionlist = exceptionlist + field + "<br/>";
                                    }
                                    writeError("Merge fields used in the template document listed do not exist in the data source." + exceptionlist + "Your file couldn't be created");
                                    throw new Exception("Merge fields used in the template document listed do not exist in the data source." + exceptionlist + "Your file couldn't be created");
                                }

                                fieldsVerified = true;
                            }

                            wrdDoc.MailMerge.SuppressBlankLines = true;
                            wrdMailMerge.DataSource.FirstRecord = startingRecord;
                            wrdMailMerge.DataSource.LastRecord = endingRecord;
                            wrdApp.Options.CheckSpellingAsYouType = false;
                            wrdApp.Options.CheckGrammarAsYouType = false;
                            wrdApp.Options.SuggestSpellingCorrections = false;
                            wrdApp.Options.CheckGrammarWithSpelling = false;
                            wrdApp.ActiveDocument.ShowSpellingErrors = false;
                            wrdDoc.ShowSpellingErrors = false;
                            wrdDoc.ShowGrammaticalErrors = false;
                            wrdDoc.SpellingChecked = true;
                            wrdDoc.GrammarChecked = true;

                            // Perform mail merge.
                            try
                            {
                                writeError("Merge Ready to execute");
                                wrdMailMerge.Destination = WdMailMergeDestination.wdSendToNewDocument;
                                wrdMailMerge.Execute(ref oFalse);
                                writeError("Merge Executed.");

                            }
                            catch (Exception ex)
                            {
                                writeError(ex.StackTrace);
                                throw new Exception("There was a problem Merging the document");
                                CleanUp();
                                return;
                            }
                            Hashtable fields = new Hashtable();

                            foreach (MailMergeDataField item in wrdDoc.MailMerge.DataSource.DataFields)
                            {
                                if (item.Name == _field1 || item.Name == _field2 || item.Name == _field3 ||
                                    item.Name == _field4 || item.Name == _field5)
                                    fields.Add(item.Name, item.Value);
                            }
                            writeError("Fields Checked.");
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


                                if (!String.IsNullOrEmpty(_suffix))
                                    fileNamePrefix += _delimiter + _suffix;
                                fileNamePrefix += _delimiter;
                                writeError("Prefix and delimiters checked.");
                            }

                            char[] invalidFileNameChars = Path.GetInvalidFileNameChars();

                            for (int i = 0; i < invalidFileNameChars.Length; i++)
                            {
                                if (fileNamePrefix.Contains(invalidFileNameChars[i].ToString()))
                                {
                                    if (_cbChecked)
                                        fileNamePrefix = fileNamePrefix.Replace(invalidFileNameChars[i].ToString(),
                                                                                _delimiter);
                                    else
                                        fileNamePrefix = fileNamePrefix.Replace(invalidFileNameChars[i].ToString(), "_");
                                }
                            }
                            writeError("Invalid characters checked.");

                            if (_documentType == pdfDocument)
                            {
                                if (fileNamePrefix == "")
                                    fileNamePrefix = "PDF";
            
                                newPath = zipFilePath + fileNamePrefix + counter + ".pdf";
                            }
                            else if (_documentType == wordDocument)
                            {
                                if (fileNamePrefix == "")
                                    fileNamePrefix = "Word";
        
                                newPath = zipFilePath + fileNamePrefix + counter + ".doc";
                            }

                            array.Add(newPath);
                            writeError("New File path created.");
                            wrdDoc.Close(ref oFalse, ref oMissing, ref oMissing);
                            writeError("Doc Closed.");
                            foreach (_Document item in wrdApp.Documents)
                            {
                                item.Activate();
                                wrdApp.ActiveDocument.SpellingChecked = true;
                                wrdApp.ActiveDocument.GrammarChecked = true;
                                wrdApp.ActiveDocument.ShowSpellingErrors = false;

                                if (_documentType == pdfDocument)
                                    item.ExportAsFixedFormat(newPath, Word.WdExportFormat.wdExportFormatPDF, false,
                                                             WdExportOptimizeFor.wdExportOptimizeForPrint);
                                else if (_documentType == wordDocument)
                                    item.SaveAs(newPath, Word.WdSaveFormat.wdFormatDocument);

                                item.Close(ref oFalse, ref oMissing, ref oMissing);
                            }

                            #endregion

                            // PDF File Size Reducer

                            #region

                            if (_documentType == pdfDocument)
                            {
                                byte[] b = File.ReadAllBytes(newPath);
                                PdfReader reader = new PdfReader(b);

                                using (MemoryStream m = new MemoryStream())
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
                                writeError("PDF format created.");
                            }

                            #endregion

                            //if (endingRecord == totalRecords) break;

                            #region . Remove Rows .

                            List<string> quotelist = File.ReadAllLines(_datasourceTxt,Encoding.Default ).ToList();
                            string firstItem = quotelist[0];
                          //  quotelist.RemoveRange(1, 1);
                            if (RecoredSize < quotelist.Count)
                                quotelist.RemoveRange(1, RecoredSize);
                            else
                                quotelist.RemoveRange(1,quotelist.Count-1);

                          
                                totalRecords = quotelist.Count - 1;
                          

                            File.WriteAllLines(_datasourceTxt, quotelist.ToArray(), Encoding.Default);
                            writeError("Source Line Removed.");
                            #endregion
                            if (totalRecords > RecoredSize)
                                endingRecord = RecoredSize;
                            else endingRecord = totalRecords;
                           // if ((endingRecord + RecoredSize) >= totalRecords) endingRecord = totalRecords;
                        }
                        catch (Exception docEx)
                        {
                            writeError(docEx.Message);
                         //   sendEmail(email, docEx.Message + "\n\n" + docEx.StackTrace.ToString());
                            sendEmail(email, docEx.Message);
                            throw docEx;
                        }
                    }
                    else
                    {
                        writeError("Document cannot be opened");
                        throw new Exception("Document cannot be opened");
                    }


                    CleanUp();

                    counter++;
                }

                #region . Throw file to the client .

                using (ZipFile zip = new ZipFile())
                {
                    Guid objGuid = Guid.NewGuid();
                    writeError("Zip File Creation Called.");
                    zip.AddFiles(array, "");
                    string time = DateTime.Now.ToString("dd-MM-yy-HH-mm-ss");
                    time = time.Replace("-", _delimiter);
                    finalFilePath = FinalDocs + "MailMerge_" + email + "_" + time + ".zip";
                    zip.UseZip64WhenSaving = Zip64Option.Always;
                    // UseZip64WhenSaving
                    zip.Save(zipFilePath + "MailMergedZip_" + objGuid + ".zip");
                    File.Move(zipFilePath + "MailMergedZip_" + objGuid + ".zip", finalFilePath);

                    // File.Move(zipFilePath + "MailMergedZip" + time + ".zip", finalFilePath);
                    writeError("Zip File Created and Saved");
                }
                foreach (var item in array)
                {
                    try
                    {
                        File.Delete(item);
                    }
                    catch
                    {
                        // Do not catch
                    }
                }

                #endregion

                endTime = DateTime.Now;
                var tSpend = endTime - startTime;
                if (!string.IsNullOrEmpty(email))
                    sendemail(email, finalFilePath, tSpend.TotalSeconds);
                writeError("Please check the provided email inbox for merge results");
            }
            //fix for LPS-318
            catch (System.Threading.ThreadAbortException ex)
            {
                writeError("LPS-318 aborted exception");
                //do nothing
                //catch Thread was aborted exception
            }
            catch (System.Runtime.InteropServices.COMException comException)
            {
                writeError("Com Exception");
                //do nothing
                //Catch command failed exception while opening a word document
            }
            catch (Exception oEx)
            {
               
                sendEmail(email, oEx.Message + "\n\n" + oEx.StackTrace.ToString()); 
                
            }
            finally
            {
                try
                {
                   
                    // this.Log("Ended:" + formatPath + ";" + dataSourcePath + ";" + totalRecords + ";" + email + ";" + _documentType, EventLogEntryType.Information);
                    CleanUp();
                    // Release References.
                    wrdSelection = null;
                    wrdMailMerge = null;
                    wrdDoc = null;
                    wrdApp.Quit(ref oMissing, ref oMissing, ref oMissing);
                    writeError("wrdApp Closed");
                    wrdApp = null;
                    GC.Collect();

                }
                catch
                {
                    writeError("Close off session");
                }
            }

            if (File.Exists(dataSourcePath)) File.Delete(dataSourcePath);
            if (File.Exists(_datasourceTxt)) File.Delete(_datasourceTxt);
        }

        private object OMissing(ref string dataSourcePath)
        {
            try
            {
                wrdApp2 = new Application();
                wrdApp2.Options.SaveNormalPrompt = false;
                wrdApp2.Options.SavePropertiesPrompt = false;
                wrdApp2.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                // Create an instance of Word and make it visible.
                wrdApp2.Visible = false;
                writeError("Initiate New Word Application.");
            }
            catch
            {
                writeError("Failed to initialize MSWord. Check the permision on server.");
                throw new Exception("Failed to initialize MSWord. Check the permision on server.");
            }



            object oMissing = System.Reflection.Missing.Value;

            wrdDoc2 = wrdApp2.Documents.Open(_datasourceTxt, false, true, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, WdOpenFormat.wdOpenFormatText, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            dataSourcePath = _datasourceTxt.Substring(0, _datasourceTxt.Length - 3) + "doc";
            try
            {
                if (File.Exists(dataSourcePath)) File.Delete(dataSourcePath);
            }
            catch (Exception e)
            {
                writeError("Error Trying to Delete Source File Doc");
            }
            wrdDoc2.SaveAs(dataSourcePath, WdSaveFormat.wdFormatDocument, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            wrdDoc2.Close();
            try
            {
                if (wrdDoc2 != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wrdDoc2);
                wrdDoc2 = null;
                if (wrdApp2 != null)
                {
                    wrdApp2.Quit(ref oMissing, ref oMissing, ref oMissing);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wrdApp2);
                }
                wrdApp2 = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return oMissing;
        }


        private void sendemail(string email, string path, double time)
        {
            writeError("File Email called" + path);
            if (ConfigurationManager.AppSettings["SendEmail"].ToString() == "1")
            {
            System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
            string emailFrom = ConfigurationManager.AppSettings["emailFrom"];
            message.From = new MailAddress(emailFrom);
            message.To.Add(email);
            string emails = ConfigurationManager.AppSettings["emailTo"];

            foreach (var item in emails.Split(",".ToCharArray()))
            {
                message.Bcc.Add(item);
            }
            message.IsBodyHtml = true;
            string pathFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "EmailHeader.txt");
            StreamReader objStreamReader = File.OpenText(pathFile);
            string subject = objStreamReader.ReadLine();
            message.Subject = subject;
            string content = ReadFile("EmailTemplate.htm");
            message.Body = content.Replace("#path#", path).Replace("#ExecutionTime#", time.ToString());
            SmtpClient client = new SmtpClient();
            client.Send(message);
            writeError("Error Email sent");
        }
        }

        private void sendEmail(string email, string errMessage)
        {
            writeError("Error email called" + errMessage);
            if (ConfigurationManager.AppSettings["SendEmail"].ToString() == "1")
            {
                System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
                string emailFrom = ConfigurationManager.AppSettings["emailFrom"];
                message.From = new MailAddress(emailFrom);
                message.To.Add(email);
                string emails = ConfigurationManager.AppSettings["emailTo"];

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
                SmtpClient client = new SmtpClient();
                client.Send(message);
                writeError("Error Email sent");
            }
        }

       public static string ReadFile(string FileName)
        {
            try
            {
                String FILENAME = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, FileName);
                StreamReader objStreamReader = File.OpenText(FILENAME);
                String contents = objStreamReader.ReadToEnd();
                return contents;
            }
            catch (Exception ex)
            {

            }
            return "";
        } 

        public void CleanUp()
        {
            try
            {
                if (wrdDoc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wrdDoc);
                    writeError("Com Object wrdDoc Released succesfully");
                }
                wrdDoc = null;
                if (wrdApp != null)
                {
                    wrdApp.Quit(ref oMissing, ref oMissing, ref oMissing);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wrdApp);
                    writeError("Com Object wrdApp Released succesfully");
                }
                wrdApp = null;
                GC.Collect();
                writeError("Garbage Collected Ran");
                Running = false;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
           
        }
        void Log(string sEvent, EventLogEntryType type)
        {
            string sSource = "LPMailMergeService";
            string sLog = "Application";

            if (!EventLog.SourceExists(sSource))
                EventLog.CreateEventSource(sSource, sLog);

            EventLog.WriteEntry(sSource, sEvent,
                type);
        }

        public void Run()
        {
            Running = true;
            MessageQueue msMq;// = msMq = new MessageQueue(queueName);
            if (!MessageQueue.Exists(queueName))
            {
                msMq = MessageQueue.Create(queueName);
            }
            else
            {
                msMq = new MessageQueue(queueName);
            }

            try
            {
                while (msMq.GetAllMessages().Count() > 0)
                {
                    var item = msMq.Receive();
                    item.Formatter = new XmlMessageFormatter(new Type[] { typeof(string) });
                    var message = item.Body.ToString().Split(";".ToCharArray());
                    string emailto = string.Empty;
                    if (message.Length > 3)
                        emailto = message[3];
                    this.MergeMailWithMultipleRecords(message[0], message[1], Convert.ToInt32(message[2]),emailto, message[4], Convert.ToBoolean(message[5]), message[6]
                        , message[7], message[8], message[9], message[10], message[11], message[12], message[13],message[14]);

                    this.Log(string.Join(";", message), EventLogEntryType.Information);

                }

            }
            catch (MessageQueueException ee)
            {
                this.Log(ee.ToString(), EventLogEntryType.Error);
            }
            catch (Exception eee)
            {
                this.Log(eee.ToString(), EventLogEntryType.Error);
            }
            finally
            {
                msMq.Close();
            }
            Running = false;

        }


        void FileCopy(string fileFrom, string fileTo, long bufferSize = 1024)
        {

            try
            {
                File.Copy(fileFrom, fileTo);
                File.Delete(fileFrom);
            }
            catch
            { }
            //byte[] buffer = new byte[bufferSize];
            //using (FileStream inStream = new FileStream(fileFrom, FileMode.Open, FileAccess.Read, FileShare.Read))
            //{
            //  //  File.Copy(@"C:\inetpub\wwwroot\docs\t1.docx", @"\\192.168.3.15\Temp\Imran\LPSDocs\t1.docx");
            //    using (FileStream outStream = new FileStream(fileTo, FileMode.Create, FileAccess.Write, FileShare.Write))
            //    {
            //        while (inStream.Position < inStream.Length - buffer.Length)
            //        {
            //            inStream.Read(buffer, 0, buffer.Length);
            //            outStream.Write(buffer, 0, buffer.Length);
            //        }

            //        // Copy the remaining part.
            //        buffer = new byte[inStream.Length - inStream.Position];
            //        inStream.Read(buffer, 0, buffer.Length);
            //        outStream.Write(buffer, 0, buffer.Length);
            //    }
            //}
        }
    }
}
