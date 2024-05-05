using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using A = DocumentFormat.OpenXml.Drawing;
using WP = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace MiniSoftware;

public class MiniWordPicture : IMiniWordComponent
{
    public string Path { get; set; }
    private string _extension;
    public string Extension
    {
        get
        {
            if (Path != null)
                return System.IO.Path.GetExtension(Path).ToUpperInvariant().Replace(".", "");
            else
            {
                return _extension.ToUpper();
            }
        }
        set { _extension = value; }
    }

    internal ImagePartType GetImagePartType
    {
        get
        {
            switch (Extension.ToLower())
            {
                case "bmp": return ImagePartType.Bmp;
                case "emf": return ImagePartType.Emf;
                case "ico": return ImagePartType.Icon;
                case "jpg": return ImagePartType.Jpeg;
                case "jpeg": return ImagePartType.Jpeg;
                case "pcx": return ImagePartType.Pcx;
                case "png": return ImagePartType.Png;
                case "svg": return ImagePartType.Svg;
                case "tiff": return ImagePartType.Tiff;
                case "wmf": return ImagePartType.Wmf;
                default:
                    throw new NotSupportedException($"{_extension} is not supported");
            }
        }
    }

    public byte[] Bytes { get; set; }

    /// <summary>
    /// Unit is Pixel
    /// </summary>
    public Int64Value Width { get; set; } = 400;

    internal Int64Value Cx { get { return Width * 9525; } }

    /// <summary>
    /// Unit is Pixel
    /// </summary>
    public Int64Value Height { get; set; } = 400;

    //format resource from http://openxmltrix.blogspot.com/2011/04/updating-images-in-image-placeholde-and.html
    internal Int64Value Cy { get { return Height * 9525; } }

    public void Execute(WordprocessingDocument docx, OpenXmlElement run, IMiniWordComponent value)
    {
        var pic = value as MiniWordPicture;
        byte[] l_Data = null;
        if (pic.Path != null)
        {
            l_Data = File.ReadAllBytes(pic.Path);
        }
        if (pic.Bytes != null)
        {
            l_Data = pic.Bytes;
        }

        var mainPart = docx.MainDocumentPart;

        var imagePart = mainPart.AddImagePart(pic.GetImagePartType);
        using (var stream = new MemoryStream(l_Data))
        {
            imagePart.FeedData(stream);
            AddPicture(run, mainPart.GetIdOfPart(imagePart), pic);
        }
    }

    private static void AddPicture(OpenXmlElement appendElement, string relationshipId, MiniWordPicture pic)
    {
        // Define the reference of the image.
        var element =
             new Drawing(
                 new WP.Inline(
                     new WP.Extent() { Cx = pic.Cx, Cy = pic.Cy },
                     new WP.EffectExtent()
                     {
                         LeftEdge = 0L,
                         TopEdge = 0L,
                         RightEdge = 0L,
                         BottomEdge = 0L
                     },
                     new WP.DocProperties()
                     {
                         Id = (UInt32Value)1U,
                         Name = $"Picture {Guid.NewGuid().ToString()}"
                     },
                     new WP.NonVisualGraphicFrameDrawingProperties(
                         new A.GraphicFrameLocks() { NoChangeAspect = true }),
                     new A.Graphic(
                         new A.GraphicData(
                             new PIC.Picture(
                                 new PIC.NonVisualPictureProperties(
                                     new PIC.NonVisualDrawingProperties()
                                     {
                                         Id = (UInt32Value)0U,
                                         Name = $"Image {Guid.NewGuid().ToString()}.{pic.Extension}"
                                     },
                                     new PIC.NonVisualPictureDrawingProperties()),
                                 new PIC.BlipFill(
                                     new A.Blip(
                                         new A.BlipExtensionList(
                                             new A.BlipExtension()
                                             {
                                                 Uri =
                                                    $"{{{Guid.NewGuid().ToString("n")}}}"
                                             })
                                     )
                                     {
                                         Embed = relationshipId,
                                         CompressionState =
                                         A.BlipCompressionValues.Print
                                     },
                                     new A.Stretch(
                                         new A.FillRectangle())),
                                 new PIC.ShapeProperties(
                                     new A.Transform2D(
                                         new A.Offset() { X = 0L, Y = 0L },
                                         new A.Extents() { Cx = pic.Cx, Cy = pic.Cy }),
                                     new A.PresetGeometry(
                                         new A.AdjustValueList()
                                     )
                                     { Preset = A.ShapeTypeValues.Rectangle }))
                         )
                         { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                 )
                 {
                     DistanceFromTop = (UInt32Value)0U,
                     DistanceFromBottom = (UInt32Value)0U,
                     DistanceFromLeft = (UInt32Value)0U,
                     DistanceFromRight = (UInt32Value)0U,
                     EditId = "50D07946"
                 });
        appendElement.Append((element));
    }

}
