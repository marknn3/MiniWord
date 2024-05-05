using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;

namespace MiniSoftware;

public class MiniWordColorText : IMiniWordComponent, IMiniWordComponentList
{
    public string FontColor { get; set; }
    public string Text { get; set; }
    public string HighlightColor { get; set; }

    public void Execute(WordprocessingDocument docx, OpenXmlElement run, IMiniWordComponent value)
    {
        var list = new[] { value as MiniWordColorText };
        Execute(docx, run, list);
    }

    public void Execute(WordprocessingDocument docx, OpenXmlElement run, IList<IMiniWordComponent> list)
    {
        var colorText = AddColorText(list as IList<MiniWordColorText>);
        run.Append(colorText);
    }

    private static RunProperties AddColorText(IList<MiniWordColorText> miniWordColorTextArray)
    {
        RunProperties runProps = new RunProperties();
        foreach (var miniWordColorText in miniWordColorTextArray)
        {
            Text text = new Text(miniWordColorText.Text);
            Color color = new Color() { Val = miniWordColorText.FontColor?.Replace("#", "") };
            Shading shading = new Shading() { Fill = miniWordColorText.HighlightColor?.Replace("#", "") };
            runProps.Append(shading);
            runProps.Append(color);
            runProps.Append(text);
        }

        return runProps;
    }
}
