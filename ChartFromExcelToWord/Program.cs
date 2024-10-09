
using ConsoleApp1;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Reflection;

Console.WriteLine("Chart from Excel file to Word file");

try
{
    string excelPath = @"Book1.xlsx";
    string docPath = @"Doc1.docx";
    Sheet selectedSheet = null;

    using (var fileStream = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
    {
        using (SpreadsheetDocument excelDocument = SpreadsheetDocument.Open(fileStream, false))
        {
            try
            {
                WorkbookPart workbookPart = excelDocument.WorkbookPart;
                selectedSheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(k => k.Name == "Sheet1");
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(selectedSheet.Id);
                DrawingsPart drawingPart = worksheetPart.DrawingsPart;

                using (var docx = WordprocessingDocument.Open(docPath, true))
                {
                    try
                    {
                        MainDocumentPart mainPart = docx.MainDocumentPart;
                        ChartOperations objChart = new ChartOperations();
                        objChart.AddChartObjectToWordDoc(ref mainPart, drawingPart);
                    }
                    finally
                    {
                        docx.Save();
                        docx.Dispose();
                        Console.WriteLine("Completed. Please check the file");
                    }
                }
            }
            finally
            {
                excelDocument.Dispose();
            }
        }
    }
}
catch (TargetInvocationException ex)
{
    Console.WriteLine($"Inner Exception: {ex.InnerException?.Message}");
}