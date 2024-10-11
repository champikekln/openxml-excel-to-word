using ChartFromExcelToWord;
using ChartFromExcelToWord.Common.enums;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ChartFromExcelToWord
{
    public class LabelOperations : ILabelOperations
    {
        private string _labelValue { get; set; } = string.Empty;
        private bool _isBoldText { get; set; } = false;
        private string _fontColor { get; set; } = "000000";
        private bool _isItalic { get; set; } = false;
        private bool _isUnderline { get; set; } = false;
        private string _fontSize { get; set; } = "24";
        private CustomJustification _justification = CustomJustification.Center;

        public LabelOperations() { }
        public LabelOperations(ref MainDocumentPart mainPart, LabelProps props)
        {
            _labelValue = props.labelValue;
            _isBoldText = props.isBoldText;
            _fontColor = props.fontColor;
            _isItalic = props.isItalic;
            _isUnderline = props.isUnderline;
            _fontSize = props.fontSize;
            _justification = props.justification;

            AddLabel(ref mainPart);
        }

        public void AddLabel(ref MainDocumentPart mainPart, LabelProps props)
        {
            _labelValue = props.labelValue;
            _isBoldText = props.isBoldText;
            _fontColor = props.fontColor;
            _isItalic = props.isItalic;
            _isUnderline = props.isUnderline;
            _fontSize = props.fontSize;
            _justification = props.justification;

            AddLabel(ref mainPart);
        }

        public void AddLabel(ref MainDocumentPart mainPart)
        {
            Paragraph p = new Paragraph();
            ParagraphProperties pp = new ParagraphProperties();
            pp.Justification = new Justification() { Val = GetJustificationValue(_justification) };
            p.Append(pp);

            DocumentFormat.OpenXml.Wordprocessing.Run r = new DocumentFormat.OpenXml.Wordprocessing.Run();
            RunProperties runProperties = new RunProperties();

            if (_isBoldText)
            {
                DocumentFormat.OpenXml.Wordprocessing.Bold bold = new DocumentFormat.OpenXml.Wordprocessing.Bold();
                bold.Val = OnOffValue.FromBoolean(true);
                runProperties.Append(bold);
            }

            if (_isItalic)
            {
                DocumentFormat.OpenXml.Wordprocessing.Italic italic = new DocumentFormat.OpenXml.Wordprocessing.Italic();
                italic.Val = OnOffValue.FromBoolean(true);
                runProperties.Append(italic);
            }

            if (_isUnderline)
            {
                DocumentFormat.OpenXml.Wordprocessing.Underline underline = new DocumentFormat.OpenXml.Wordprocessing.Underline() { Val = UnderlineValues.Single };
                runProperties.Append(underline);
            }

            DocumentFormat.OpenXml.Wordprocessing.Color color = new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = _fontColor };
            runProperties.Append(color);

            DocumentFormat.OpenXml.Wordprocessing.FontSize fontSizeElement = new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = _fontSize };
            runProperties.Append(fontSizeElement);

            r.Append(runProperties);
            Text t = new Text(_labelValue) { Space = SpaceProcessingModeValues.Preserve };
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

    public class LabelProps
    {
        public string labelValue { get; set; }
        public bool isBoldText { get; set; }
        public string fontColor { get; set; }
        public bool isItalic { get; set; }
        public bool isUnderline { get; set; }
        public string fontSize { get; set; }
        public CustomJustification justification = CustomJustification.Center;
    }
}
