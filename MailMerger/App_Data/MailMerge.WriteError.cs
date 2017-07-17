using System;
using System.Configuration;
using System.IO;

namespace MailMerger
{
    public partial class MailMerge
    {
        public static void WriteError(string ex)
        {
            if (ConfigurationManager.AppSettings["LogError"] == "1")
            {
                string sLogFormat = DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString() + " ==> " + ex;
                string sPathName = ConfigurationManager.AppSettings["LogPath"].ToString() + "ErrorLog\\";

                if (!Directory.Exists(sPathName))
                {
                    Directory.CreateDirectory(sPathName);
                }

                var sw = new StreamWriter(sPathName + "webapplog.txt", true);
                sw.WriteLine(sLogFormat);
                sw.Flush();
                sw.Close();
            }
        }
    }
}