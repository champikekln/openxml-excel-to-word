
using ChartFromExcelToWord;
using ConsoleApp1;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
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
                        ILabelOperations objLabel = new LabelOperations();
                        MainDocumentPart mainPart = docx.MainDocumentPart;
                        ChartOperations objChart1 = new ChartOperations(ref mainPart, drawingPart, new ChartProperties() { chartName= "chart1", chartCaption="Chart 1", primaryLabel ="Chart 1 Primary Label", isBold =true, isItalic =true, fontColor = "000000", isUnderlined=true, fontSize="24" }, objLabel);
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