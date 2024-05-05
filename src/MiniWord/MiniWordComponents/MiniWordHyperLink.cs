using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MiniSoftware.Utility;
using System;
using System.Collections.Generic;

namespace MiniSoftware;

public class MiniWordHyperLink : IMiniWordComponent, IMiniWordComponentList
{
    public string Url { get; set; }

    public string Text { get; set; }

    public UnderlineValues UnderLineValue { get; set; } = UnderlineValues.Single;

    public TargetFrameType TargetFrame { get; set; } = TargetFrameType.Blank;

    public void Execute(WordprocessingDocument docx, OpenXmlElement run, IMiniWordComponent value)
    {
        var list = new[] { value as MiniWordHyperLink };
        Execute(docx, run, list);
    }

    public void Execute(WordprocessingDocument docx, OpenXmlElement run, IList<IMiniWordComponent> list)
    {
        var links = list as IList<MiniWordHyperLink>;

        foreach (var linkInfo in links)
        {
            var mainPart = docx.MainDocumentPart;
            var hyperlink = GetHyperLink(mainPart, linkInfo);
            run.Append(hyperlink);
            run.Append(new Break());
        }
    }

    private static Hyperlink GetHyperLink(MainDocumentPart mainPart, MiniWordHyperLink linkInfo)
    {
        var hr = mainPart.AddHyperlinkRelationship(new Uri(linkInfo.Url), true);
        Hyperlink xmlHyperLink = new Hyperlink(
            new RunProperties(
                new RunStyle { Val = "Hyperlink", },
                new Underline { Val = linkInfo.UnderLineValue },
                new Color { ThemeColor = ThemeColorValues.Hyperlink }),
                new Text(linkInfo.Text)
            )
        {
            DocLocation = linkInfo.Url,
            Id = hr.Id,
            TargetFrame = linkInfo.GetTargetFrame()
        };
        return xmlHyperLink;
    }

    private string GetTargetFrame()
    {
        return TargetFrame switch
        {
            TargetFrameType.Blank => "_blank",
            TargetFrameType.Top => "_top",
            TargetFrameType.Self => "_self",
            TargetFrameType.Parent => "_parent",
            _ => "_blank"
        };
    }
}
