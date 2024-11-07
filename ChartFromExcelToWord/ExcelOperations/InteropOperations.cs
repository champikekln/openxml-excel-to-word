using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ChartFromExcelToWord.ExcelOperations
{
    public static class InteropOperations
    {
        public static void RecalCulate(string filePath)
        {
            filePath = @"C:\Champike\Personal\Apps\excelFile\ChartFromExcelToWord\openxml-excel-to-word\ChartFromExcelToWord\bin\Debug\net8.0\Book1.xlsx";
            Application excelApp = new Application();
            excelApp.Visible = false;
            Workbook workbook = excelApp.Workbooks.Open(filePath);

            try
            {
                workbook.RefreshAll(); 
                workbook.Application.Calculate(); 
                workbook.Save();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                workbook.Close(false);
                excelApp.Quit();
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }
    }
}
