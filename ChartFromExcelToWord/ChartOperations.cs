using ChartFromExcelToWord;
using ChartFromExcelToWord.Common.enums;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ConsoleApp1
{
    public class ChartOperations : CommonOperations
    {
        private ILabelOperations _labelOperations { get; set; }
        private string _chartName { get; set; } = "chart1";
        private string _chartCaption { get; set; } = "caption1";
        private string _primaryLabel { get; set; } = "Chart1 - Primary Label";
        private CustomJustification _justification { get; set; } = CustomJustification.Center;
        private bool _isBold { get; set; } = true;
        private bool _isItalic { get; set; } = true;
        private string _fontColor { get; set; } = "000000";
        private bool _isUnderlined { get; set; } = true;
        private string _fontSize { get; set; } = "24";

        public ChartOperations(ref MainDocumentPart mainPart, DrawingsPart drawingPart, ChartProperties chartProps, ILabelOperations labelOperations)
        {
            _chartName = chartProps.chartName;
            _chartCaption = chartProps.chartCaption;
            _primaryLabel = chartProps.primaryLabel;
            _justification = chartProps.justification;
            _isBold = chartProps.isBold;
            _isItalic = chartProps.isItalic;
            _fontColor = chartProps.fontColor;
            _fontSize = chartProps.fontSize;
            _labelOperations = labelOperations;

            AddChartObjectToWordDoc(ref mainPart, drawingPart);
        }

        private void AddChartObjectToWordDoc(ref MainDocumentPart mainPart, DrawingsPart drawingPart)
        {
            string relId;
            string chartNameInExcel = $"/xl/charts/{_chartName}.xml";
            ChartPart selectedChartPart = (ChartPart)drawingPart.ChartParts.FirstOrDefault(x => x.Uri.OriginalString.Equals(chartNameInExcel));

            if (selectedChartPart != null)
            {
                if (!string.IsNullOrEmpty(_primaryLabel))
                {
                    _labelOperations.AddLabel(ref mainPart, new LabelProps() { fontColor = _fontColor, fontSize = _fontSize, isBoldText = _isBold, isItalic = _isItalic, isUnderline = _isUnderlined, labelValue = _primaryLabel });
                }

                ChartPart importedChartPart = mainPart.AddPart<ChartPart>(selectedChartPart);
                relId = string.Format("{0}{1}", "R", Guid.NewGuid().ToString());
                ChartPart chartPart = mainPart.AddNewPart<ChartPart>(relId);
                chartPart.ChartSpace = (ChartSpace)selectedChartPart.ChartSpace.Clone();

                var paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph() { RsidParagraphAddition = "00C75AEB", RsidRunAdditionDefault = "000F3EFF" };

                DocumentFormat.OpenXml.Wordprocessing.Run run = new DocumentFormat.OpenXml.Wordprocessing.Run();
                DocumentFormat.OpenXml.Wordprocessing.Drawing drawing = new DocumentFormat.OpenXml.Wordprocessing.Drawing();
                DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline inline = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline();
                inline.Append(new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = 5274310L, Cy = 3076575L });
                DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties docPros = new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties() { Id = (UInt32Value)1U, Name = _chartName };
                inline.Append(docPros);
                DocumentFormat.OpenXml.Drawing.Graphic g = new DocumentFormat.OpenXml.Drawing.Graphic();
                var graphicData = new DocumentFormat.OpenXml.Drawing.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };
                var chartReference = new ChartReference() { Id = relId };
                graphicData.Append(chartReference);
                g.Append(graphicData);
                inline.Append(g);
                drawing.Append(inline);
                run.Append(drawing);
                paragraph.Append(run);
                mainPart.Document.Body.Append(paragraph);
                AddChartCaption(_chartCaption, "Figure", ref mainPart);
            }
        }
        private void AddChartCaption(string caption, string name, ref MainDocumentPart mainPart)
        {
            if (!string.IsNullOrEmpty(caption))
            {
                string[] captionArray = caption.Split(':');
                string captionValue = string.Empty;
                string objectNumber = string.Empty;
                captionValue = caption;

                Paragraph paragraph = new Paragraph() { RsidParagraphAddition = "00EB2BC7", RsidParagraphProperties = "00566D73", RsidRunAdditionDefault = "00566D73" };
                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Caption" };
                paragraphProperties2.Append(paragraphStyleId1);
                Run run1 = new Run();
                Text text1 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text1.Text = name + " ";
                run1.Append(text1);
                SimpleField simpleField = null;
                simpleField = new SimpleField() { Instruction = " SEQ Figure \\* ARABIC " };

                Run run2 = new Run();
                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();
                runProperties1.Append(noProof1);
                Text text2 = new Text();
                text2.Text = objectNumber;
                run2.Append(runProperties1);
                run2.Append(text2);
                simpleField.Append(run2);

                Run run3 = new Run();
                Text text3 = new Text();
                text3.Text = ":" + captionValue;
                run3.Append(text3);

                BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", Id = "0" };
                BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = "0" };
                paragraph.Append(paragraphProperties2);
                paragraph.Append(run1);
                paragraph.Append(simpleField);
                paragraph.Append(run3);
                paragraph.Append(bookmarkStart1);
                paragraph.Append(bookmarkEnd1);
                mainPart.Document.Body.Append(paragraph);
            }
        }
    }

    public class ChartProperties
    {
        public string chartName { get; set; }
        public string chartCaption { get; set; }
        public string primaryLabel { get; set; }
        public CustomJustification justification { get; set; }
        public bool isBold { get; set; }
        public bool isItalic { get; set; }
        public string fontColor { get; set; }
        public bool isUnderlined { get; set; }
        public string fontSize { get; set; }
    }
}