using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace MiniSoftware
{
    public interface IMiniWordComponent
    {
        public void Execute(WordprocessingDocument docx, OpenXmlElement run, IMiniWordComponent value);
    }
}
