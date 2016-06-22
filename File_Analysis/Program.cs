using System;
using System.Collections.Generic;
using System.Net;
using System.IO;
using System.Linq;
using System.Diagnostics;
using Ionic.Zip;

namespace File_Analysis
{
    struct DirectoryItem
    {
        public Uri BaseUri;

        public string AbsolutePath
        {
            get
            {
                return string.Format("{0}/{1}", BaseUri, Name);
            }
        }

        public DateTime DateCreated;
        public bool IsDirectory;
        public string Name;
        public List<DirectoryItem> Items;
    }

    class Program
    {
        static void Main(string[] args)
        {
            List<DirectoryItem> listing = GetDirectoryInformation("ftp://dug/periodic_reports/mga_amp_file_analysis/calendar_day", "petep", "tr3ple12");
            Console.ReadKey();
        }

        static List<DirectoryItem> GetDirectoryInformation(string address, string username, string password) {
            FtpWebRequest request = (FtpWebRequest)FtpWebRequest.Create(address);
            request.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
            request.Credentials = new NetworkCredential(username, password);
            request.UsePassive = true;
            request.UseBinary = true;
            request.KeepAlive = false;

            List<DirectoryItem> returnValue = new List<DirectoryItem>();
            string[] list = null;

            using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
            using (StreamReader reader = new StreamReader(response.GetResponseStream()))
            {
                list = reader.ReadToEnd().Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            }

            foreach (string line in list)
            {
                // Windows FTP Server Response Format
                // DateCreated    IsDirectory    Name
                string data = line;

                //// Parse date
                string date = data.Substring(55, 8);
                DateTime dateTime = DateTime.ParseExact(date, "yyyyMMdd", null);
                data = data.Remove(0, 40);

                // Parse <DIR>
                string dir = data.Substring(0, 5);
                bool isDirectory = dir.Equals("<dir>", StringComparison.InvariantCultureIgnoreCase);
                data = data.Remove(0, 5);
                data = data.Remove(0, 10);

                // Parse name
                string name = data;

                // Create directory info
                DirectoryItem item = new DirectoryItem();
                item.BaseUri = new Uri(address);
                item.DateCreated = dateTime;
                item.IsDirectory = isDirectory;
                item.Name = name;

                string nDate = DateTime.Now.ToString("yyyyMMdd");

                FtpWebRequest dRequest = (FtpWebRequest)FtpWebRequest.Create(item.AbsolutePath);
                dRequest.Method = WebRequestMethods.Ftp.DownloadFile;
                dRequest.Credentials = new NetworkCredential("petep", "tr3ple12");
                dRequest.UsePassive = true;
                dRequest.UseBinary = true;
                dRequest.KeepAlive = false;

                FtpWebResponse dResponse = (FtpWebResponse)dRequest.GetResponse();
                Stream dResponseStream = dResponse.GetResponseStream();

                FileStream writer = new FileStream(@"C:/Users/petep/Desktop/FA_Reports/" + name, FileMode.Create);

                long length = dResponse.ContentLength;
                int bufferSize = 2048;
                int readCount;
                byte[] buffer = new byte[2048];

                readCount = dResponseStream.Read(buffer, 0, bufferSize);
                while (readCount > 0)
                {
                    writer.Write(buffer, 0, readCount);
                    readCount = dResponseStream.Read(buffer, 0, bufferSize);
                }

                dResponseStream.Close();
                dResponse.Close();
                writer.Close();

                string zipToUnpack = "C:/Users/petep/Desktop/FA_Reports/" + name;
                string unpackDirectory = "C:/Users/petep/Desktop/FA_Reports/Extracted_Files/";
                using (ZipFile zip = ZipFile.Read(zipToUnpack))
                {
                    foreach (ZipEntry e in zip)
                    {
                        e.Extract(unpackDirectory, ExtractExistingFileAction.OverwriteSilently);
                    }
                }

                //Debug.WriteLine(item.AbsolutePath);
                item.Items = item.IsDirectory ? GetDirectoryInformation(item.AbsolutePath, username, password) : null;

                returnValue.Add(item);
            }
            return returnValue;
        }
    }
}