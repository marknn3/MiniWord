using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;

namespace MiniSoftware;

public interface IMiniWordComponentList
{
    public void Execute(WordprocessingDocument docx, OpenXmlElement run, IList<IMiniWordComponent> list);
}
