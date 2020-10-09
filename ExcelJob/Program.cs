using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Configuration;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJob
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelFuncs excel = new ExcelFuncs(@"planilha.xlsx", 1);

            //gravando na planilha
            List<string> listaNomes = new List<string> { "Nomes", "Marcos", "joão", "Milena", "Zeca", "Jon" };
            List<string> listaEmail = new List<string> { "Email", "mm@gmail.com", "jj@gmail.com", "mi@gmail.com", "ze@gmail.com", "jon@gmail.com" };
            List<string> listaPhones = new List<string> { "Phone", "(11) 9 9986-3478", "(11) 9 9945-3728", "(11) 9 9347-8549", "(11) 9 9583-5873", "(11) 9 9334-2294" };
            List<string> listaPaises = new List<string> { "Paises", "EUA", "BRASIL", "DINAMARCA", "RUSSIA", "ESPANHA"  };
            List<string> listaAtivos = new List<string> { "Ativo", "True", "False", "True", "True", "False" };

            excel.WriteExcel(listaNomes, 1);
            excel.WriteExcel(listaEmail, 2);
            excel.WriteExcel(listaPhones, 3);
            excel.WriteExcel(listaPaises, 4);
            excel.WriteExcel(listaAtivos, 5);

            excel.Save();

            //ler e separar planilha
            List<string> colunasExcel = excel.ReadExcel();
            List<string> colunaNomesParaLista = excel.ColunaPraLista(colunasExcel, 0, 5);
            List<string> colunaEmailParaLista = excel.ColunaPraLista(colunasExcel, 1, 5);
            List<string> colunaPhoneParaLista = excel.ColunaPraLista(colunasExcel, 2, 5);
            List<string> colunaPaisParaLista = excel.ColunaPraLista(colunasExcel, 3, 5);
            List<string> colunaAtivoParaLista = excel.ColunaPraLista(colunasExcel, 4, 5);

            excel.Close();

            foreach (string x in colunaNomesParaLista)
                Console.WriteLine(x);

            Console.WriteLine("------------------");

            foreach (string x in colunaEmailParaLista)
                Console.WriteLine(x);

            Console.WriteLine("------------------");

            foreach (string x in colunaPhoneParaLista)
                Console.WriteLine(x);

            Console.WriteLine("------------------");

            foreach (string x in colunaPaisParaLista)
                Console.WriteLine(x);

            Console.WriteLine("------------------");

            foreach (string x in colunaAtivoParaLista)
                Console.WriteLine(x);

            Console.WriteLine("------------------");

           //MailFuncs.ReadEmail();
           //MailFuncs.ListEmail();

            //converter docx para PDF
            PDFFormat.ConverterWordParaPDF(@"C:\Users\wviegas\Documents\comum.docx", @"convertido.pdf");

            //enviar email
            //  MailFuncs.SendEmail("emailFrom", "emailTo", "teste email", "testando email");

            string subject = MailFuncs.FindEmailBySubject("Enviando Attachment");
            string file = @"C:\Users\wviegas\Documents\will\planilha.docx";
            MailFuncs.SendEmailWithAttachment("emailFrom", "emailTo", subject, "segue planilha", file);
        }
    }
}
