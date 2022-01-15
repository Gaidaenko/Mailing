using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.Net.Mail;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Microsoft.Win32;
using System.IO;
using System;
using System.Collections.Generic;

namespace Mailing
{  
    class Program
    {
        public static string patchAddress;
      //  public static string attachment;
        public static int nextAddress = 1;
        public static string[] dirs = Directory.GetFiles(@"c:\\mails\\att\\");

        static void Main(string[] args)
        {
            patchAddress = "C:\\mails\\mails.xlsx";

            Excel.Application xlsApp = new Excel.Application();
            Workbook ObjWorkBook = xlsApp.Workbooks.Open(patchAddress, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Worksheet ObjWorkSheet;
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            
            Range Rng, CheckingRow;
            Rng = xlsApp.get_Range("A1", "A100");

            var dataArr = (object[,])Rng.Value;

            if(dataArr[1,nextAddress] == null)
            {
                Console.WriteLine("Таблица пустая или начинается не с адреса");
                return;
            }

            while (dataArr[nextAddress, 1] != null)
            {               
                 List<object> list = new List<object>();
                 list.Add(dataArr[nextAddress, 1]);

                 foreach (var result in list)
                 {
                    MailAddress from = new MailAddress("MyMail@gmail.com", "Тест рассылки");
                    MailAddress to = new MailAddress(result.ToString());
                    MailMessage m = new MailMessage(from, to);

                    foreach (var items in dirs)
                    {
                       // attachment = items;
                        Attachment att = new Attachment(items);
                        m.Attachments.Add(att);
                    }
                                        
                    
                    m.Subject = "Тема тестовой рассылки: ";
                    m.Body = ("Тестовая рассылка: ");
                    m.IsBodyHtml = true;
                    SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
                    smtp.Credentials = new NetworkCredential("MyMail@gmail.com", "Password");
                    smtp.EnableSsl = true;
                    smtp.Send(m);
                 }

                nextAddress++;
                
            }            
            if (dataArr[nextAddress, 1] == null)
            {
                Console.WriteLine("Рассылка закончена");
                Console.ReadKey();
                return;
            }

            Console.ReadKey();          
        }
    }
}
