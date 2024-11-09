using Microsoft.Office.Interop.Excel;

namespace ChartFromExcelToWord.ExcelOperations
{
    public abstract class InteropOperations
    {
        public void RecalCulate(string filePath)
        {
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
