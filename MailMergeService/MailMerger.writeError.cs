using System;
using System.Configuration;
using System.IO;

namespace MailMergeService
{
    public partial class MailMerger
    {
        public static void writeError(string ex)
        {
            if (ConfigurationManager.AppSettings["LogError"] == "1")
            {
                string sLogFormat = DateTime.Now.ToShortDateString() + " " +
                                    DateTime.Now.ToLongTimeString() + " ==> " + ex;
                string sPathName = ConfigurationManager.AppSettings["LogPath"] + "ErrorLog\\";

                if (!Directory.Exists(sPathName))
                {
                    Directory.CreateDirectory(sPathName);
                }

                StreamWriter sw = new StreamWriter(sPathName + "windowsService.txt", true);
                sw.WriteLine(sLogFormat);
                sw.Flush();
                sw.Close();
            }
        }
    }
}