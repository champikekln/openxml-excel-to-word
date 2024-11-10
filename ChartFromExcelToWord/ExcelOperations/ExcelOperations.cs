using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DataTable = System.Data.DataTable;
using Formula = DocumentFormat.OpenXml.Drawing.Charts.Formula;
using Values = DocumentFormat.OpenXml.Drawing.Charts.Values;

namespace ChartFromExcelToWord.ExcelOperations
{
    public class ExcelOperations : InteropOperations
    {
        private string xAxisColumn { get; set; }
        private string yAxisColumn { get; set; }
        private int startRow { get; set; }
        private int endRow { get; set; }
        private string sheetName { get; set; }
        private string chartName { get; set; }
        private string filePath { get; set; }
        private DataTable chartTable { get; set; }

        public ExcelOperations(string xAxisColumnVal, string yAxisColumnVal, int startRowVal, int endRowVal, string sheetNameVal, string chartNameVal, string filepathVal, DataTable chartTableVal)
        {
            xAxisColumn = xAxisColumnVal;
            yAxisColumn = yAxisColumnVal;
            startRow = startRowVal;
            endRow = endRowVal;
            sheetName = sheetNameVal;
            chartName = chartNameVal;
            filePath = filepathVal;
            chartTable = chartTableVal;

            WriteDataTableToExcel(filePath, chartTable);
        }

        private void WriteDataTableToExcel(string filePath, DataTable dataTable)
        {
            ReArrangeChartData();
            using (var workbook = new XLWorkbook("Book1.xlsx"))
            {
                try
                {
                    var worksheet = workbook.Worksheet(sheetName);
                    var lastRow = startRow;

                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataTable.Columns.Count; j++)
                        {
                            var value = dataTable.Rows[i][j];

                            if (value is int || value is long || value is short)
                            {
                                worksheet.Cell(lastRow + i + 1, j + 1).Value = Convert.ToInt32(value);
                            }
                            else if (value is float || value is double || value is decimal)
                            {
                                worksheet.Cell(lastRow + i + 1, j + 1).Value = Convert.ToDouble(value);
                            }
                            else if (value is DateTime)
                            {
                                worksheet.Cell(lastRow + i + 1, j + 1).Value = Convert.ToDateTime(value);
                            }
                            else
                            {
                                worksheet.Cell(lastRow + i + 1, j + 1).Value = value?.ToString();
                            }
                        }
                    }
                    workbook.ForceFullCalculation = true;
                    workbook.CalculationOnSave = true;
                    workbook.Save();
                }
                finally
                {
                    workbook.Dispose();
                }
            }
            RecalCulate(filePath);
        }

        private void ReArrangeChartData()
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, true))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                var drawingPart = worksheetPart.DrawingsPart;
                if (drawingPart == null) return;

                foreach (var chartPart in drawingPart.ChartParts)
                {
                    Chart chart = chartPart.ChartSpace.Elements<Chart>().First();

                    string chartNameInExcel = $"/xl/charts/{chartName}.xml";
                    if (chartPart.Uri.ToString() == chartNameInExcel)
                    {
                        var chartSeries = chart.Descendants<PieChartSeries>().FirstOrDefault();
                        if (chartSeries != null)
                        {
                            //Set Y axis data
                            var values = chartSeries.Descendants<Values>().FirstOrDefault();
                            if (values != null)
                            {
                                var formula = values.Descendants<Formula>().FirstOrDefault();
                                if (formula != null)
                                {
                                    formula.Text = $"{sheetName}!${yAxisColumn}${startRow}:${yAxisColumn}${endRow}";
                                }
                            }
                            //Set X axis data
                            var categoryAxisData = chartSeries.Descendants<CategoryAxisData>().FirstOrDefault();
                            if (categoryAxisData != null)
                            {
                                var formula = categoryAxisData.Descendants<Formula>().FirstOrDefault();
                                if (formula != null)
                                {
                                    formula.Text = $"{sheetName}!${xAxisColumn}${startRow}:${xAxisColumn}${endRow}";
                                }
                            }
                        }
                        document.Save();
                        break;
                    }
                }
            }
        }
    }
}
