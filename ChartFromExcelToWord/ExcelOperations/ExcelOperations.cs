using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.IO;
using static ChartFromExcelToWord.DynamicCharts.DynamicCharts;

namespace ChartFromExcelToWord.ExcelOperations
{
    public class ExcelOperations
    {

        public ExcelOperations()
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("Salary", typeof(double));

            dataTable.Rows.Add("John Doe", 1000);
            dataTable.Rows.Add("Jane Smith", 60000);
            dataTable.Rows.Add("Samuel Johnson", 1000);
            dataTable.Rows.Add("Samuel Johnson1", 55000);
            dataTable.Rows.Add("Samuel Johnson2", 2000);
            dataTable.Rows.Add("Samuel Johnson3", 56000);

            string filePath = @"Book1.xlsx";

            WriteDataTableToExcel(filePath, dataTable);

            ChartTableDef objChartDef1 = new ChartTableDef();
            objChartDef1.Id = 1;
            objChartDef1.Name = "Chart1";
            objChartDef1.Title = "Chart 1";
            objChartDef1.XAxisTitle = "X Axis";
            objChartDef1.YAxisTitle = "Y Axis";
            objChartDef1.startingColumnIndex = 1;
            objChartDef1.startingRowIndex = 2;
            objChartDef1.chartType = "Column";
            objChartDef1.Columns = new List<ChartTableColumn>() { new ChartTableColumn() { Id = 1, Name = "Name", Format = "string", columnIndex = 1 }, new ChartTableColumn() { Id = 2, Name = "Salary", Format = "double", columnIndex = 2 } };
        }

        private void WriteDataTableToExcel(string filePath, DataTable dataTable)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                try
                {
                    var worksheet = workbook.Worksheet("Sheet1");
                    var lastRow = worksheet.LastRowUsed().RowNumber();

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
                    //workbook.CalculateMode = ClosedXML.Excel.XLCalculateMode.Auto;
                    workbook.ForceFullCalculation = true;
                    workbook.CalculationOnSave = true;
                    workbook.Save();
                }
                finally
                {
                    workbook.Dispose();
                }
            }
            InteropOperations.RecalCulate(filePath);
        }

        private void WriteDataTableToExcel1(string filePath)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                try
                {
                    workbook.Save();
                }
                finally
                {
                    workbook.Dispose();
                }
            }
        }



        //private void WriteDataTableToExcel(string filePath, DataTable dataTable)
        //{
        //    using (var workbook = new XLWorkbook())
        //    {
        //        var worksheet = workbook.Worksheets.Add("Sheet1");

        //        for (int i = 1; i < dataTable.Rows.Count; i++)
        //        {
        //            for (int j = 0; j < dataTable.Columns.Count; j++)
        //            {
        //                var cellValue = dataTable.Rows[i][j] != null ? dataTable.Rows[i][j].ToString() : string.Empty;
        //                worksheet.Cell(i + 1, j + 1).Value = cellValue; 
        //            }
        //        }
        //        workbook.Save();
        //    }
        //}

        //private static void WriteDataTableToExcel(string filePath, DataTable dataTable)
        //{
        //    using (var workbook = new XLWorkbook())
        //    {
        //        var worksheet = workbook.Worksheets.Add("Sheet1");
        //        worksheet.Cell(1, 1).InsertTable(dataTable);
        //        workbook.SaveAs(filePath);

        //    }
        //}
    }
}
