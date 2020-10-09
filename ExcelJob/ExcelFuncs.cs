using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelJob
{
    class ExcelFuncs
    {
        private string path = "";
        private Excel._Application excel = new Excel.Application();
        private Workbook wb;
        private Worksheet ws;

        public ExcelFuncs(string path, int sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }

        public void WriteExcel(List<string> lista, int coluna)
        {
            int count = 1;

            for (int i = 0; i < lista.Count; i++)
            {
                ws.Cells[count, coluna] = lista[i];
                count++;
            }
        }

        public void Save()
        {
            wb.Save();
        }

        public List<string> ReadExcel()
        {
            List<string> result = new List<string>();
            Excel.Range xlRange = ws.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        result.Add(xlRange.Cells[i, j].Value2.ToString());
                    }
                }

            }
            Marshal.ReleaseComObject(xlRange);

            return result;
        }

        public List<string> ColunaPraLista(List<string> listaDeColunas, int posicaoColuna, int qntColunas)
        {
            List<string> result = new List<string>();

            for (int i = posicaoColuna; i < listaDeColunas.Count; i += qntColunas)
            {
                result.Add(listaDeColunas[i]);
            }

            return result;
        }

        public void Close()
        {
            GC.Collect();
            // GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(ws);

            wb.Close();
            Marshal.ReleaseComObject(wb);

            excel.Quit();
            Marshal.ReleaseComObject(excel);
        }
    }
}
