using System;
using System.Windows.Forms;
using System.Net;
using System.IO;

namespace File_Analysis
{
    class Program
    {
        public DateTime CurrentDateTime()
        {
            DateTime nDate = DateTime.Now;

            return nDate;
        }

        static void Main()
        {
            try
            {
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://dug/periodic_reports/mga_amp_file_analysis/calendar_day/" + DateTime.Now + "8000*_File Analysis_calendar_day.zip");
                request.Method = WebRequestMethods.Ftp.DownloadFile;
                request.Credentials = new
                NetworkCredential("petep", "tr3ple12");
                FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                Stream responseStream = response.GetResponseStream();
                StreamReader reader = new
                StreamReader(responseStream);
                MessageBox.Show(reader.ReadToEnd().ToString());
                reader.Close();
                response.Close();
                Console.WriteLine(reader.ReadToEnd());
            }
            catch
            (NotSupportedException ne)
            {
                MessageBox.Show(ne.Message);
            }
        }
    }
}
