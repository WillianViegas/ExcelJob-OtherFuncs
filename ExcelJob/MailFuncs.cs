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
using System.Net.Mime;

namespace ExcelJob
{
    class MailFuncs
    {
        public static void SendMail(string emailDe, string emailPara, string assunto, string corpo)
        {
            using (SmtpClient client = new SmtpClient("smtp.gmail.com", 587)
            {
                Credentials = new NetworkCredential("username", "pass"),
                EnableSsl = true
            })
            {
                client.Send(emailDe, emailPara, assunto, corpo);
            }
        }

        public static void SendMailWithAttachment(string emailDe, string emailPara, string assunto, string corpo, string file)
        {

            using (SmtpClient client = new SmtpClient("smtp.gmail.com", 587)
            {
                Credentials = new NetworkCredential("username", "pass"),
                EnableSsl = true
            })
            {
                MailMessage message = new MailMessage(
                emailDe,
                emailPara,
                assunto,
                corpo);

                if (System.IO.File.Exists(file))
                {
                    Attachment data = new Attachment(file, MediaTypeNames.Application.Octet);

                    ContentDisposition disposition = data.ContentDisposition;
                    disposition.CreationDate = System.IO.File.GetCreationTime(file);
                    disposition.ModificationDate = System.IO.File.GetLastWriteTime(file);
                    disposition.ReadDate = System.IO.File.GetLastAccessTime(file);
                    message.Attachments.Add(data);
                }

                try
                {
                    client.Send(message);
                }
                catch(Exception e)
                {
                    Console.WriteLine("Exceção em SendMailWithAttachment: {0}", e.Message);
                }
            }
        }

        static StringBuilder builder = new StringBuilder();
        public static void ReadEmail()
        {
            using(OpenPop.Pop3.Pop3Client client = new OpenPop.Pop3.Pop3Client())
            {
                client.Connect("pop.gmail.com", 995, true);
                client.Authenticate("recent:username", "pass");

                if (client.Connected)
                {
                    Console.WriteLine("Checking inbox");
                    var count = client.GetMessageCount();
                    OpenPop.Mime.Message message = client.GetMessage(count);
                    OpenPop.Mime.MessagePart plainText = message.FindFirstPlainTextVersion();
                    
                    builder.Append("Subject: " + message.Headers.Subject + "\n");
                    builder.Append("Date: " + message.Headers.Date + "\n");
                    builder.Append("Body: " + plainText.GetBodyAsText());
                    Console.WriteLine(builder.ToString());

                    //verifica se existem anexos e caso sim salva-os na pasta
                    var att = message.FindAllAttachments();

                    foreach (var ado in att)
                    {
                        ado.Save(new System.IO.FileInfo(System.IO.Path.Combine(@"C:\Users\wviegas\Documents\will", ado.FileName)));
                    }
                }
            }
        }

        public static void ListEmail()
        {
            using(OpenPop.Pop3.Pop3Client client = new OpenPop.Pop3.Pop3Client())
            {
                client.Connect("pop.gmail.com", 995, true);
                client.Authenticate("recent:username", "pass");

                if (client.Connected)
                {
                    int count = client.GetMessageCount();
                    List<Message> allMessages = new List<Message>(count);
                    for (int i = count; i > 0; i--)
                    {
                        allMessages.Add(client.GetMessage(i));
                    }

                    foreach(Message msg in allMessages)
                    {
                        string subject = msg.Headers.Subject;
                        string date = msg.Headers.Date;

                        OpenPop.Mime.MessagePart plainText = msg.FindFirstPlainTextVersion();
                        Console.WriteLine("Subject: " + subject);
                        Console.WriteLine("Date: " + date);
                        Console.WriteLine("Body: " + plainText.GetBodyAsText());
                    }
                }
            }
        }
            
    }
}
