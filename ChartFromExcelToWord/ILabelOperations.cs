using DocumentFormat.OpenXml.Packaging;

namespace ChartFromExcelToWord
{
    public interface ILabelOperations
    {
        public void AddLabel(ref MainDocumentPart mainPart, LabelProps props);
    }
}
