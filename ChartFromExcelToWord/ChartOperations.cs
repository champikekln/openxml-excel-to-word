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
        private string _chartName = "chart1";
        private string _chartCaption = "caption1";
        private string _primaryLabel = "Chart1 - Primary Label";
        private string _secondaryLabel = "Chart1 - Secondary Label";
        private CustomJustification _justification = CustomJustification.Center;
        private bool _isBold = true;
        private bool _isItalic = true;
        private string _fontColor = "000000";
        private bool _isUnderlined = true;
        private string _fontSize = "24";

        public bool AddChartObjectToWordDoc(ref MainDocumentPart mainPart, DrawingsPart drawingPart)
        {
            string relId;
            string chartNameInExcel = string.Format("{0}{1}{2}", "/xl/charts/", _chartName, ".xml");
            ChartPart selectedChartPart = (ChartPart)drawingPart.ChartParts.FirstOrDefault(x => x.Uri.OriginalString.Equals(chartNameInExcel));

            if (selectedChartPart != null)
            {
                AddLabel(ref mainPart, _primaryLabel, _isBold, _fontColor, _isItalic, _isUnderlined, _fontSize, _justification);
                AddSecondaryLabel(ref mainPart, _secondaryLabel, false);

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
                AddChartCaption("caption1", "Figure", ref mainPart);
            }
            return true;
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

        public static void AddSecondaryLabel(ref MainDocumentPart mainPart, string value, bool isboldText = true)
        {
            Paragraph p = new Paragraph();
            ParagraphProperties pp = new ParagraphProperties();
            pp.Justification = new Justification() { Val = JustificationValues.Center };
            p.Append(pp);
            DocumentFormat.OpenXml.Wordprocessing.Run r = new DocumentFormat.OpenXml.Wordprocessing.Run();

            if (isboldText)
            {
                RunProperties runProperties = r.AppendChild(new RunProperties());
                DocumentFormat.OpenXml.Wordprocessing.Bold bold = new DocumentFormat.OpenXml.Wordprocessing.Bold();
                bold.Val = OnOffValue.FromBoolean(true);
                runProperties.AppendChild(bold);
            }

            Text t = new Text(value) { Space = SpaceProcessingModeValues.Preserve };
            r.Append(t);
            p.Append(r);
            mainPart.Document.Body.Append(p);
        }
    }
}