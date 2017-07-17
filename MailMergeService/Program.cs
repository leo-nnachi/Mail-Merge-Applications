using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;

namespace MailMergeService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            //Guid objGuid = Guid.NewGuid();
            //string zipFilePath = "C:\\docs1\\";
            //string tempSourcePath = "C:\\docs\\penmast_200.txt";
            //string dataSourcePath = tempSourcePath.Replace("/", "\\");
            //if (!File.Exists(dataSourcePath))
            //    throw new Exception("Datasource file does not exists. Please provide correct path");
            //if (!Directory.Exists(zipFilePath))
            //    Directory.CreateDirectory(zipFilePath);
            //File.Copy(dataSourcePath, zipFilePath + "\\datasource_" + objGuid.ToString() + ".doc", true);
            //dataSourcePath = zipFilePath + "datasource_" + objGuid.ToString() + ".doc";
            //// For Service Testing
            //MailMerger ser = new MailMerger();

            //// more than 500 records
            //ser.MergeMailWithMultipleRecords("C:\\docs\\northerntruststatement.doc", dataSourcePath, 1, "ilyas@itsbettertogether.co.uk", "PDF Document", false, "", "", ""
            //    , "", "", "", "", "");

            //ser = new MailMerger();
            //ser.MergeMailWithMultipleRecords("C:\\docs\\northerntruststatement.doc", dataSourcePath, 1, "ilyas@itsbettertogether.co.uk", "Word Document", false, "", "", ""
            //    , "", "", "", "", "");

            //ser = new MailMerger();
            //ser.MergeMailWithMultipleRecords("C:\\docs\\northerntruststatement.doc", dataSourcePath, 1, "ilyas@itsbettertogether.co.uk", "PDF Document", true, "prefix", "field1", "field2", "field3"
            //    , "field4", "field5", "suffix", "_");

            //ser = new MailMerger();
            //ser.MergeMailWithMultipleRecords("C:\\docs\\northerntruststatement.doc", dataSourcePath, 1, "ilyas@itsbettertogether.co.uk", "Word Document", true, "prefix", "field1", "field2", "field3"
            //    , "field4", "field5", "suffix", "_");

            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[] 
			{ 
				new MergeService()
			};
            ServiceBase.Run(ServicesToRun);
        }
    }
}
