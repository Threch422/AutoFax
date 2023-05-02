using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace AutoFax
{
    public class ExcelHandler
    {
        protected Application excelApp;
        protected Workbook excelWorkBook;
        protected Worksheet excelWorkSheet;
        protected Range usedRange;

        protected int Rows => usedRange.Rows.Count;
        protected int Columns => usedRange.Columns.Count;

        public ExcelHandler()
        {
            excelApp = new Application();
            excelApp.Visible = false;
            excelWorkBook = null;
            excelWorkSheet = null;
            usedRange = null;
        }

        public ExcelHandler(string filePath)
        {
            excelApp = new Application();
            excelApp.Visible = false;
            excelWorkBook = excelApp.Workbooks.Open(filePath);
            excelWorkSheet = excelWorkBook.Sheets[1];
            usedRange = excelWorkSheet.UsedRange;
        }

        ~ExcelHandler()
        {
            excelWorkBook.Close(0);
            excelApp.Quit();
        }

        // Returning the FaxNumber and RecipientName
        protected internal virtual IEnumerable<(string, string)> GetRowsInfo()
        {
            for (int row = 2; row <= Rows; row++)
                yield return
                    (usedRange.Cells[row, 1].Value2?.ToString() ?? string.Empty, 
                     usedRange.Cells[row, 2].Value2?.ToString() ?? string.Empty);
        }
    }
}
