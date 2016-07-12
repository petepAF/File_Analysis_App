using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Net;
using System.IO;
using System.Net.Mail;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Ionic.Zip;
using System.Threading.Tasks;

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
        static void Main()
        {
            //Time when method needs to be called
            var DailyTime = "08:30:00";
            var timeParts = DailyTime.Split(new char[1] { ':' });

            while (true)
            {

                var dateNow = DateTime.Now;
                var date = new DateTime(dateNow.Year, dateNow.Month, dateNow.Day,
                           int.Parse(timeParts[0]), int.Parse(timeParts[1]), int.Parse(timeParts[2]));

                TimeSpan ts;
                if (date > dateNow)
                    ts = date - dateNow;
                else
                {
                    date = date.AddDays(1);
                    ts = date - dateNow;
                }

                //waits certan time and run the code
                Task.Delay(ts).ContinueWith((x) =>
                {

                    Console.WriteLine("Gathering files from the FTP server...\n\n");

                    List<DirectoryItem> listing = GetDirectoryInformation("ftp://dug/periodic_reports/mga_amp_file_analysis/calendar_day", "petep", "tr3ple12");

                    string[] fileNames = Directory.GetFiles(@"C:\Users\petep\Desktop\FA_Reports\Extracted_Files\", "*Completed*" + DateTime.Today.ToString("yyyyMMdd") + "*");

                    Console.WriteLine("Grabbed " + fileNames[0] + " from directory\n\n");

                    Console.WriteLine("Coverting file to XLS format...\n\n");

                    if (fileNames.Length > 0)
                    {
                        Application xlApp;
                        Workbook sWorkBook;
                        Worksheet sWorkSheet;
                        Workbook dWorkBook;
                        Worksheet dWorkSheet;
                        Workbook fWorkBook;
                        Worksheet fWorkSheet;
                        Workbook faWorkBook;
                        Worksheet faWorkSheet;
                        object misValue = System.Reflection.Missing.Value;

                        xlApp = new Application();
                        xlApp.DisplayAlerts = false;

                        sWorkBook = xlApp.Workbooks.Open(@"C:/Users/petep/Desktop/FA_Reports/File_Analysis_TEMPLATE.xltx");
                        sWorkSheet = (Worksheet)sWorkBook.Worksheets.get_Item(1);
                        dWorkBook = xlApp.Workbooks.Open(fileNames[0], misValue, misValue, XlFileFormat.xlCSV, misValue, misValue, 1, misValue, ",", misValue, misValue, misValue, misValue, misValue, misValue);
                        dWorkSheet = (Worksheet)dWorkBook.Worksheets.get_Item(1);
                        dWorkBook.SaveAs(@"C:\Users\petep\Desktop\FA_Reports\FA", XlFileFormat.xlWorkbookNormal, misValue, misValue, 0, misValue, XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue);
                        dWorkBook.Close(false, misValue, misValue);

                        Console.WriteLine("Opening XLS formatted file to generate report...\n\n");

                        fWorkBook = xlApp.Workbooks.Open("C:/Users/petep/Desktop/FA_Reports/FA.xls");
                        fWorkSheet = (Worksheet)fWorkBook.Worksheets.get_Item(1);

                        Range from = fWorkSheet.get_Range("A1:G1000");
                        Range to = sWorkSheet.get_Range("A1:G1000");

                        from.Copy(to);

                        sWorkBook.Close(true, @"C:\Users\petep\Desktop\FA_Reports\Completed_Daily_Report.xlsx", misValue);
                        faWorkBook = xlApp.Workbooks.Open(@"C:\Users\petep\Desktop\FA_Reports\Completed_Daily_Report.xlsx");
                        faWorkSheet = (Worksheet)faWorkBook.Worksheets.get_Item(2);

                        Console.WriteLine("Report succesffully generated!\n\n");

                        Range nothing = faWorkSheet.get_Range("D12:F12");
                        Range lowPriority = faWorkSheet.get_Range("D12");
                        Range mediumPriority = faWorkSheet.get_Range("E12");
                        Range highPriority = faWorkSheet.get_Range("F12");

                        bool lowZero = lowPriority.Value == 0;
                        bool medZero = mediumPriority.Value == 0;
                        bool higZero = highPriority.Value == 0;
                        bool lowGT = lowPriority.Value > 0;
                        bool medGT = mediumPriority.Value > 0;
                        bool higGT = highPriority.Value > 0;
                        bool lowLT = lowPriority.Value >= 0;
                        bool medLT = mediumPriority.Value >= 0;

                        Console.WriteLine("Preparing email alert...\n");

                        if (lowZero & medZero & higZero)
                        {

                            // Do nothing.
                            faWorkBook.Close(true, misValue, misValue);
                          }
                        else if (lowGT & medZero & higZero)
                        {

                            SmtpClient smtp = new SmtpClient("mail.accidentfund.com", 25);
                            MailMessage message = new MailMessage();

                            message.To.Add("netalerts@accidentfund.com");
                            message.CC.Add("Pete.Peterson@accidentfund.com");
                            message.CC.Add("Greg.Warning@accidentfund.com");
                            message.CC.Add("Nicole2R@accidentfund.com");
                            message.From = new MailAddress("FileAnalysisAlert@accidentfund.com");

                            Console.WriteLine("Building email alert...\n");

                            message.Subject = "File Analysis (Low Priority Alert) - " + DateTime.Today.ToString("MM-dd-yyyy");
                            message.Body = "Threshold limit reached\n\nProcesses > 20 Minutes: " + lowPriority.Value + "\n"
                                + "\nPlease see attached report.\n";

                            faWorkBook.Close(true, misValue, misValue);

                            Attachment attachment;
                            attachment = new Attachment("C:/Users/petep/Desktop/FA_Reports/Completed_Daily_Report.xlsx");
                            attachment.Name = "File Analysis Report - " + DateTime.Today.ToString("MM/dd/yyyy") + ".xlsx";
                            message.Attachments.Add(attachment);

                            smtp.Send(message);
                        }
                        else if (lowLT & medGT & higZero)
                        {

                            SmtpClient smtp = new SmtpClient("mail.accidentfund.com", 25);
                            MailMessage message = new MailMessage();

                            message.To.Add("netalerts@accidentfund.com");
                            message.CC.Add("Pete.Peterson@accidentfund.com");
                            message.CC.Add("Greg.Warning@accidentfund.com");
                            message.CC.Add("Nicole2R@accidentfund.com");
                            message.From = new MailAddress("FileAnalysisAlert@accidentfund.com");

                            Console.WriteLine("Building email alert...\n");
                            message.Subject = "File Analysis (Medium Priority Alert) - " + DateTime.Today.ToString("MM-dd-yyyy");
                            message.Body = "Threshold limit reached\n\nProcesses > 20 Minutes: " + lowPriority.Value + "\n"
                                + "Processes > 40 Minutes: " + mediumPriority.Value + "\n"
                                + "\nPlease see attached report.\n";

                            faWorkBook.Close(true, misValue, misValue);

                            Attachment attachment;
                            attachment = new Attachment("C:/Users/petep/Desktop/FA_Reports/Completed_Daily_Report.xlsx");
                            attachment.Name = "File Analysis Report - " + DateTime.Today.ToString("MM/dd/yyyy") + ".xlsx";
                            message.Attachments.Add(attachment);

                            smtp.Send(message);
                        }
                        else if (lowLT & medLT & higGT)
                        {

                            SmtpClient smtp = new SmtpClient("mail.accidentfund.com", 25);
                            MailMessage message = new MailMessage();

                            message.To.Add("netalerts@accidentfund.com");
                            message.CC.Add("Pete.Peterson@accidentfund.com");
                            message.CC.Add("Greg.Warning@accidentfund.com");
                            message.CC.Add("Nicole2R@accidentfund.com");
                            message.From = new MailAddress("FileAnalysisAlert@accidentfund.com");

                            Console.WriteLine("Building email alert...\n");
                            message.Subject = "File Analysis (High Priority Alert) - " + DateTime.Today.ToString("MM-dd-yyyy");
                            message.Body = "Threshold limit reached\n\nProcesses > 20 Minutes: " + lowPriority.Value + "\n"
                                + "Processes > 40 Minutes: " + mediumPriority.Value + "\n"
                                + "Processes > 60 Minutes: " + highPriority.Value + "\n"
                                + "\nPlease see attached report.\n";

                            faWorkBook.Close(true, misValue, misValue);

                            Attachment attachment;
                            attachment = new Attachment("C:/Users/petep/Desktop/FA_Reports/Completed_Daily_Report.xlsx");
                            attachment.Name = "File Analysis Report - " + DateTime.Today.ToString("MM/dd/yyyy") + ".xlsx";
                            message.Attachments.Add(attachment);

                            smtp.Send(message);
                        }

                        Marshal.ReleaseComObject(faWorkSheet);
                        Marshal.ReleaseComObject(faWorkBook);
                        Marshal.ReleaseComObject(sWorkSheet);
                        Marshal.ReleaseComObject(sWorkBook);
                        Marshal.ReleaseComObject(dWorkSheet);
                        Marshal.ReleaseComObject(dWorkBook);
                        Marshal.ReleaseComObject(fWorkSheet);
                        Marshal.ReleaseComObject(fWorkBook);
                        Marshal.ReleaseComObject(xlApp);

                        foreach (Process excelProcess in Process.GetProcesses())
                        {
                            if (excelProcess.ProcessName.Equals("EXCEL"))
                            {
                                excelProcess.Kill();
                                break;
                            }
                        }

                        DirectoryInfo dFiles = new DirectoryInfo(@"C:\Users\petep\Desktop\FA_Reports\Extracted_Files\");
                        DirectoryInfo zFiles = new DirectoryInfo(@"C:\Users\petep\Desktop\FA_Reports\");

                        foreach (FileInfo file in dFiles.GetFiles())
                        {
                            file.Delete();
                        }

                        foreach (FileInfo zFile in zFiles.GetFiles("*.zip"))
                        {
                            zFile.Delete();
                        }

                        Console.WriteLine("Email sent successfully!");
                    }
                });

                Console.Read();
            } 
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