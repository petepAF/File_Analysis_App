using System;
using System.Net;
using System.IO;
using System.Windows.Forms;

namespace File_Analysis
{
    class Program
    {

        public static void DownloadFTPFiles()
        {
            string ftpAddr = "ftp://dug/periodic_reports/mga_amp_file_analysis/calendar_day/";
            string filename = "20160621080000_File%20Analysis_calendar_day.zip";
            string userName = "petep";
            string password = "tr3ple12";

            try
            {
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(ftpAddr + filename);
                request.Credentials = new NetworkCredential(userName, password);
                request.UseBinary = true; // Use binary to ensure correct dlv!
                request.Method = WebRequestMethods.Ftp.DownloadFile;

                FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                Stream responseStream = response.GetResponseStream();
                FileStream writer = new FileStream(@"c:\temp\" + filename, FileMode.Create);

                long length = response.ContentLength;
                int bufferSize = 2048;
                int readCount;
                byte[] buffer = new byte[2048];

                readCount = responseStream.Read(buffer, 0, bufferSize);
                while (readCount > 0)
                {
                    writer.Write(buffer, 0, readCount);
                    readCount = responseStream.Read(buffer, 0, bufferSize);
                }

                responseStream.Close();
                response.Close();
                writer.Close();
                MessageBox.Show(filename + " is stored at C:\\temp");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        static void Main()
        {
            DownloadFTPFiles();
        }
    }
}
