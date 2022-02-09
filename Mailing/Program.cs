using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.Net.Mail;
using Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using System.IO;
using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;

namespace Mailing
{
    public static class MailingStatus
    {       
        public static int nextAddress = 1;
        public static string patchAddress = "C:\\mails\\mails.xlsx";
        public static string patchLog = "C:\\mails\\Log.txt";
        public static string[] dirs = Directory.GetFiles(@"c:\\mails\\att\\");
        public static void notFiled()
        {
            using (FileStream log = new FileStream(patchLog, FileMode.Append, FileAccess.Write))
            {
                byte[] info = new UTF8Encoding(true).GetBytes("\n" + DateTime.Now + ": Таблица пустая, или первая строка не заполнена.");
                log.Write(info, 0, info.Length);
                return;
            }
        }
        public static void Success()
        {
            using (FileStream log = new FileStream(MailingStatus.patchLog, FileMode.Append, FileAccess.Write))
            {
                byte[] info = new UTF8Encoding(true).GetBytes("\n" + DateTime.Now + ": Рассылка успешно выполнена.");
                log.Write(info, 0, info.Length);
                return;
            }
        }
        public static void fileIsMessing()
        {
            using (FileStream log = new FileStream(MailingStatus.patchLog, FileMode.Append, FileAccess.Write))
            {
                byte[] info = new UTF8Encoding(true).GetBytes("\n" + DateTime.Now + ": Файл с адресами отсутвует или назван другим именем.");
                log.Write(info, 0, info.Length);
                return;
            }
        }
        public static void killProcessXLSX()
        {
            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            foreach (var process in List)
            {
                process.Kill();
            }
            return;
        }
    }

    class Program
    {
        //public static string attachment;
        static void Main(string[] args)
        {            
            try
            {
                Excel.Application xlsApp = new Excel.Application();
                Workbook ObjWorkBook = xlsApp.Workbooks.Open(MailingStatus.patchAddress, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Worksheet ObjWorkSheet;
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];

                Range Rng, CheckingRow;
                Rng = xlsApp.get_Range("A1", "A100");

                var dataArr = (object[,])Rng.Value;

                if (dataArr[1, MailingStatus.nextAddress] == null)
                {
                    MailingStatus.notFiled();
                    MailingStatus.killProcessXLSX();
                    return;
                }

                while (dataArr[MailingStatus.nextAddress, 1] != null)
                {
                    List<object> list = new List<object>();
                    list.Add(dataArr[MailingStatus.nextAddress, 1]);

                    foreach (var result in list)
                    {
                        MailAddress from = new MailAddress("sendtestmessages@gmail.com", "Тест рассылки");
                        MailAddress to = new MailAddress(result.ToString());
                        MailMessage m = new MailMessage(from, to);

                        foreach (var items in MailingStatus.dirs)
                        {
                            // attachment = items;
                            Attachment att = new Attachment(items);
                            m.Attachments.Add(att);
                        }

                        m.Subject = "Тема тестовой рассылки: ";
                        m.Body = ("Тестовая рассылка: ");
                        m.IsBodyHtml = true;
                        SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
                        smtp.Credentials = new NetworkCredential("sendtestmessages@gmail.com", "Password");
                        smtp.EnableSsl = true;
                        smtp.Send(m);
                    }

                    MailingStatus.nextAddress++;
                }

                if (dataArr[MailingStatus.nextAddress, 1] == null)
                {

                    MailingStatus.Success();
                    MailingStatus.killProcessXLSX();

                    return;
                }
            }
            catch
            {
                MailingStatus.fileIsMessing();

                return;   
                //test
            }          
        }
    }
}
