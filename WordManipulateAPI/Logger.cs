using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Web;

namespace WordManipulateAPI
{
    public static class Logger
    {
        public static void WriteLog(string message)
        {
            //String repository = ConfigurationManager.AppSettings["LogPath"];
            string location = ConfigurationManager.AppSettings["LogPath"];
            var path = Path.Combine(location, "logs" );
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            var filepath = Path.Combine(path, DateTime.Today.ToString("yyyy-MM-dd") + ".txt");
            StreamWriter sw = new StreamWriter(filepath, true);
            sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " : " + message);
            sw.Close();

        }
    }
}