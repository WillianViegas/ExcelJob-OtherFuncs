using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelJob
{
    class PDFFormat
    {
        public static Word.Application word = new Word.Application();
        public static Microsoft.Office.Interop.Word.Document wordDocument { get; set; }
       
        public static void ConverterWordParaPDF(string origem, string destino)
        {
            try
            {
                    wordDocument = word.Documents.Open(origem);
                    wordDocument.ExportAsFixedFormat(destino, Word.WdExportFormat.wdExportFormatPDF);
                    Console.WriteLine("Convertido com sucesso");
            }
            catch(Exception e)
            {
                throw new Exception(e.Message);
            }
            finally
            {
                wordDocument.Close();
                word.Quit();
            }
        }
    }
}
