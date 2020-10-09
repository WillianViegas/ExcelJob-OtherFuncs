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
        public static Microsoft.Office.Interop.Word.Document wordDocument { get; set; }
       public static void ConverterWordParaPDF(string origem, string destino)
        {
            try
            {
                if (!System.IO.File.Exists(destino))
                {
                    Word.Application word = new Word.Application();
                    wordDocument = word.Documents.Open(origem);
                    wordDocument.ExportAsFixedFormat(destino, Word.WdExportFormat.wdExportFormatPDF);
                    
                    word.Quit();
                    wordDocument.Close();
                    Console.WriteLine("Convertido com sucesso");
                }
                else
                {
                    Console.WriteLine("Arquivo destino já existe");
                }

            }
            catch(Exception e)
            {
                throw new Exception(e.Message);
            }
        }
    }
}
