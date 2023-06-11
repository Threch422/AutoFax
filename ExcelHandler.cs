﻿using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace AutoFax
{
    public class ExcelHandler
    {
        protected string filePath;

        public ExcelHandler()
        {
            this.filePath = string.Empty;
        }

        public ExcelHandler(string filePath)
        {
            this.filePath = filePath;
        }

        // Returning the FaxNumber and RecipientName
        protected internal virtual List<(string, string)> GetRowsInfo()
        {
            Application excelApp = new Application();
            excelApp.Visible = false;
            Workbook excelWorkbook = excelApp.Workbooks.Open(this.filePath, ReadOnly: true);
            Worksheet excelWorkSheet = excelWorkbook.Sheets[1];
            Range usedRange = excelWorkSheet.UsedRange;

            List<(string, string)> output = new List<(string, string)>();

            for (int row = 2; row <= usedRange.Rows.Count; row++)
            {
                output.Add((usedRange.Cells[row, 1].Value2?.ToString() ?? string.Empty, usedRange.Cells[row, 2].Value2?.ToString() ?? string.Empty));
            }

            excelWorkbook.Close(0);
            excelApp.Quit();

            return output;

            //for (int row = 2; row <= Rows; row++)
            //    yield return
            //        (usedRange.Cells[row, 1].Value2?.ToString() ?? string.Empty,
            //         usedRange.Cells[row, 2].Value2?.ToString() ?? string.Empty);
        }
    }
}
