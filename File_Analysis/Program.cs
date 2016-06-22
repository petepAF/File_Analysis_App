using System;
using System.Net;
using System.IO;
using System.Collections.Generic;

namespace File_Analysis
{
    class Program
    {

        public static void DownloadFTPFiles()
        {
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://dug/periodic_reports/mga_amp_file_analysis/calendar_day/20160621080000_File%20Analysis_calendar_day.zip");
            request.Method = WebRequestMethods.Ftp.DownloadFile;
            request.Credentials = new NetworkCredential("petep", "tr3ple12");
            FtpWebResponse response = (FtpWebResponse)request.GetResponse();
            Stream responseStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(responseStream);
            Console.WriteLine(reader.ReadToEnd());
            Console.WriteLine("Download Complete, status {0}", response.StatusDescription);
            reader.Close();
            response.Close();
        }

        static void Main()
        {
            DownloadFTPFiles();
        }
    }
}
