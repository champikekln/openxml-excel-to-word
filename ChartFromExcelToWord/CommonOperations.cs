using ChartFromExcelToWord.Common.enums;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ChartFromExcelToWord
{
    public class CommonOperations
    {
        public void AddPageBreak(ref MainDocumentPart mainPart)
        {
            Paragraph PageBreakParagraph = new Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Break() { Type = BreakValues.Page }));
            mainPart.Document.Body.Append(PageBreakParagraph);
        }

        public static void AddNewLine(ref MainDocumentPart mainPart)
        {
            Paragraph p = new Paragraph();
            ParagraphProperties pp = new ParagraphProperties();
            pp.Justification = new Justification() { Val = JustificationValues.Both };
            p.Append(pp);
            DocumentFormat.OpenXml.Wordprocessing.Run r = new DocumentFormat.OpenXml.Wordprocessing.Run();
            Text t = new Text("") { Space = SpaceProcessingModeValues.Preserve };
            r.Append(t);
            p.Append(r);
            mainPart.Document.Body.Append(p);
        }
    }
}
