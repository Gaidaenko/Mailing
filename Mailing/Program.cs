using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace Mailing
{
    class Program
    {
        static void Main(string[] args)
        {

            SendMail();
        }


        static void SendMail()
        {

            MailAddress from = new MailAddress("sendtestmessages@gmail.com", "Тест рассылки");
            MailAddress to = new MailAddress("yura@bitservice.ua");
            MailMessage m = new MailMessage(from, to);
            m.Subject = "Тема тестовой рассылки: ";
            m.Body = ("Тестовая рассылка: ");
            m.IsBodyHtml = true;
            SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
            smtp.Credentials = new NetworkCredential("sendtestmessages@gmail.com", "Password");
            smtp.EnableSsl = true;
            smtp.Send(m);


        }
    }
}
