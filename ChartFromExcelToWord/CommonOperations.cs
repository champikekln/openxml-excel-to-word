using ChartFromExcelToWord.Common.enums;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ChartFromExcelToWord
{
    public class CommonOperations
    {
        public void AddLabel(ref MainDocumentPart mainPart, string value, bool isBoldText = true, string hexColor = "000000", bool isItalic = false, 
            bool isUnderline = false, string fontSize = "24", CustomJustification justification = CustomJustification.Center)
        {
            Paragraph p = new Paragraph();
            ParagraphProperties pp = new ParagraphProperties();
            pp.Justification = new Justification() { Val = GetJustificationValue(justification) };
            p.Append(pp);

            DocumentFormat.OpenXml.Wordprocessing.Run r = new DocumentFormat.OpenXml.Wordprocessing.Run();
            RunProperties runProperties = new RunProperties();

            if (isBoldText)
            {
                DocumentFormat.OpenXml.Wordprocessing.Bold bold = new DocumentFormat.OpenXml.Wordprocessing.Bold();
                bold.Val = OnOffValue.FromBoolean(true);
                runProperties.Append(bold);
            }

            if (isItalic)
            {
                DocumentFormat.OpenXml.Wordprocessing.Italic italic = new DocumentFormat.OpenXml.Wordprocessing.Italic();
                italic.Val = OnOffValue.FromBoolean(true);
                runProperties.Append(italic);
            }

            if (isUnderline)
            {
                DocumentFormat.OpenXml.Wordprocessing.Underline underline = new DocumentFormat.OpenXml.Wordprocessing.Underline() { Val = UnderlineValues.Single };
                runProperties.Append(underline);
            }

            DocumentFormat.OpenXml.Wordprocessing.Color color = new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = hexColor };
            runProperties.Append(color);

            DocumentFormat.OpenXml.Wordprocessing.FontSize fontSizeElement = new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = fontSize };
            runProperties.Append(fontSizeElement);

            r.Append(runProperties);
            Text t = new Text(value) { Space = SpaceProcessingModeValues.Preserve };
            r.Append(t);
            p.Append(r);

            mainPart.Document.Body.Append(p);
        }

        private JustificationValues GetJustificationValue(CustomJustification justification)
        {
            return justification switch
            {
                CustomJustification.Left => JustificationValues.Left,
                CustomJustification.Right => JustificationValues.Right,
                CustomJustification.Center => JustificationValues.Center,
                _ => JustificationValues.Center 
            };
        }
    }
}
