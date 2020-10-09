using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;
using OpenPop;
using OpenPop.Mime;
using System.Net.Configuration;

namespace ExcelJob
{
    class MailFuncs
    {
        public static void SendMail(string emailDe, string emailPara, string assunto, string corpo)
        {
            using (SmtpClient client = new SmtpClient("smtp.gmail.com", 587)
            {
                Credentials = new NetworkCredential("xxxxxxxxx", "xxxxxxxx"),
                EnableSsl = true
            })
            {
                client.Send(emailDe, emailPara, assunto, corpo);
            }
        }

        static StringBuilder builder = new StringBuilder();
        public static void ReadEmail()
        {
            using(OpenPop.Pop3.Pop3Client client = new OpenPop.Pop3.Pop3Client())
            {
                client.Connect("pop.gmail.com", 995, true);
                client.Authenticate("recent:testwill27@gmail.com", "Teste@123");

                Console.WriteLine("Checking inbox");
                var count = client.GetMessageCount();
                OpenPop.Mime.Message message = client.GetMessage(count);
                OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
                builder.Append(plainText.GetBodyAsText());
                Console.WriteLine(builder.ToString());

                var att = message.FindAllAttachments();

                foreach (var ado in att)
                {
                    ado.Save(new System.IO.FileInfo(System.IO.Path.Combine(@"C:\Users\wviegas\Documents\will", ado.FileName)));
                }
            }
        }
            
    }
}
