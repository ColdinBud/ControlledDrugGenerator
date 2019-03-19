using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace ControlledDrugReportGenerator.Class
{
    class SendEmail
    {
        public static void SendMail(string pathFile, string mailName)
        {
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress("test@abc.edu.tw");
            mail.Subject = mailName;

            mail.To.Add("raphael.huang@ylhealth.org");
            mail.Attachments.Add(new Attachment(pathFile + ".xlsx"));
            mail.Body = mailName + " ERROR!!!!!";

            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Credentials = new System.Net.NetworkCredential("raphael.huang@ylhealth.org", "!qaz2wsX");
            smtpClient.Host = "smtp.gmail.com";
            smtpClient.Port = 25;
            smtpClient.EnableSsl = true;
            smtpClient.Send(mail);

            smtpClient.Dispose();
            mail.Dispose();
        }
    }
}
