
using ChartFromExcelToWord;
using ChartFromExcelToWord.ExcelOperations;
using ConsoleApp1;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Data;
using System.Reflection;

Console.WriteLine("Chart from Excel file to Word file");

try
{
    string excelPath = @"Book1.xlsx";
    string docPath = @"Doc1.docx";
    Sheet selectedSheet = null;

    DataTable dataTable = new DataTable();
    dataTable.Columns.Add("Name", typeof(string));
    dataTable.Columns.Add("Salary", typeof(double));

    dataTable.Rows.Add("Name1", 1000);
    dataTable.Rows.Add("Name2", 60000);
    dataTable.Rows.Add("Name3", 1000);
    dataTable.Rows.Add("Name4", 55000);
    dataTable.Rows.Add("Name5", 2000);
    dataTable.Rows.Add("Name6", 56000);

    string filePath = @"C:\Champike\GitHub\openxml-excel-to-word\ChartFromExcelToWord\bin\Debug\net8.0\Book1.xlsx";
    int sartRow = 1;
    ExcelOperations obj = new ExcelOperations("A", "B", sartRow, dataTable.Rows.Count + sartRow, "Sheet1", "chart1", filePath, dataTable);

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
                        ILabelOperations objLabel = new LabelOperations();
                        MainDocumentPart mainPart = docx.MainDocumentPart;
                        ChartOperations objChart1 = new ChartOperations(ref mainPart, drawingPart, new ChartProperties() { chartName = "chart1", chartCaption = "Chart 1", primaryLabel = "Chart 1 Primary Label", isBold = true, isItalic = true, fontColor = "000000", isUnderlined = true, fontSize = "24" }, objLabel);
                        CommonOperations objCommonOperations = new CommonOperations();
                        objCommonOperations.AddPageBreak(ref mainPart);
                        objCommonOperations.AddNewLine(ref mainPart);
                        objCommonOperations.AddNewLine(ref mainPart);
                        ChartOperations objChart2 = new ChartOperations(ref mainPart, drawingPart, new ChartProperties() { chartName = "chart1", chartCaption = "Chart 2", primaryLabel = "Chart 2 Primary Label", isBold = true, isItalic = true, fontColor = "000000", isUnderlined = true, fontSize = "24" }, objLabel);
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