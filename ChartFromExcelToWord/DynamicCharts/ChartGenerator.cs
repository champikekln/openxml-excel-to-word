using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing;
using System.Collections.Generic;
using System.Linq;
using static ChartFromExcelToWord.DynamicCharts.DynamicCharts;

public class ExcelChartHelper
{
    public void CreateChartExcel(string filePath, ChartTableDef chartDef)
    {
        //using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
        //{
        //    // Create workbook and worksheet
        //    WorkbookPart workbookPart = document.AddWorkbookPart();
        //    workbookPart.Workbook = new Workbook();
        //    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        //    worksheetPart.Worksheet = new Worksheet(new SheetData());

        //    // Add Sheets to workbook
        //    Sheets sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
        //    Sheet sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "ChartData" };
        //    sheets.Append(sheet);

        //    // Fill in data for chart
        //    FillSheetData(worksheetPart, chartDef);

        //    // Create chart part
        //    DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
        //    worksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing() { Id = worksheetPart.GetIdOfPart(drawingsPart) });
        //    drawingsPart.WorksheetDrawing = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();

        //    ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
        //    chartPart.ChartSpace = new ChartSpace();
        //    Chart chart = chartPart.ChartSpace.AppendChild(new Chart());

        //    // Set chart title
        //    chart.Append(new Title(new ChartText(new RichText(new DocumentFormat.OpenXml.Drawing.Run(new Text(chartDef.Title))))));

        //    // Create the plot area
        //    PlotArea plotArea = chart.AppendChild(new PlotArea());

        //    // Add chart type (e.g., BarChart)
        //    BarChart barChart = plotArea.AppendChild(new BarChart(new BarDirection() { Val = BarDirectionValues.Column }));

        //    // Populate chart series from data
        //    for (int i = 0; i < chartDef.Columns.Count; i++)
        //    {
        //        BarChartSeries barChartSeries = barChart.AppendChild(new BarChartSeries(
        //            new Index() { Val = (uint)i },
        //            new Order() { Val = (uint)i },
        //            new SeriesText(new DocumentFormat.OpenXml.Drawing.Charts.NumericValue() { Text = chartDef.Columns[i].Name })));

        //        // Add category axis data
        //        CategoryAxisData catAxisData = barChartSeries.AppendChild(new CategoryAxisData());
        //        StringReference strRef = new StringReference() { Formula = $"ChartData!$A$2:$A${chartDef.Rows.Count + 1}" };
        //        strRef.Append(new StringCache(new PointCount() { Val = (uint)chartDef.Rows.Count }));
        //        catAxisData.Append(strRef);

        //        // Add values for the series
        //        Values values = barChartSeries.AppendChild(new Values());
        //        NumberReference numRef = new NumberReference() { Formula = $"ChartData!${(char)('B' + i)}$2:${(char)('B' + i)}${chartDef.Rows.Count + 1}" };
        //        numRef.Append(new NumberingCache(new PointCount() { Val = (uint)chartDef.Rows.Count }));
        //        values.Append(numRef);
        //    }

        //    // Add chart axis titles if specified
        //    if (!string.IsNullOrEmpty(chartDef.XAxisTitle))
        //    {
        //        plotArea.Append(new CategoryAxis(new AxisTitle(new ChartText(new RichText(new DocumentFormat.OpenXml.Drawing.Run(new Text(chartDef.XAxisTitle)))))));
        //    }
        //    if (!string.IsNullOrEmpty(chartDef.YAxisTitle))
        //    {
        //        plotArea.Append(new ValueAxis(new AxisTitle(new ChartText(new RichText(new DocumentFormat.OpenXml.Drawing.Run(new Text(chartDef.YAxisTitle)))))));
        //    }

        //    workbookPart.Workbook.Save();
        //}
    }

    private void FillSheetData(WorksheetPart worksheetPart, ChartTableDef chartDef)
    {
        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

        // Header row
        Row headerRow = new Row();
        headerRow.AppendChild(new Cell() { CellValue = new CellValue("Category"), DataType = CellValues.String });
        foreach (var column in chartDef.Columns)
        {
            headerRow.AppendChild(new Cell() { CellValue = new CellValue(column.Name), DataType = CellValues.String });
        }
        sheetData.AppendChild(headerRow);

        // Data rows
        foreach (var row in chartDef.Rows)
        {
            Row dataRow = new Row();
            dataRow.AppendChild(new Cell() { CellValue = new CellValue(row.Name), DataType = CellValues.String });

            foreach (var column in chartDef.Columns)
            {
                // You might add logic here to look up actual values, e.g., based on a dictionary of data in each row
                dataRow.AppendChild(new Cell() { CellValue = new CellValue("0"), DataType = CellValues.Number });
            }

            sheetData.AppendChild(dataRow);
        }
    }
}
