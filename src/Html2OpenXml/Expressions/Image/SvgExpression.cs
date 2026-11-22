/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
using AngleSharp.Svg.Dom;
using AngleSharp.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using System.Text;

using a = DocumentFormat.OpenXml.Drawing;
using pic = DocumentFormat.OpenXml.Drawing.Pictures;
using wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of a <c>svg</c> element.
/// </summary>
sealed class SvgExpression(ISvgSvgElement node) : ImageExpressionBase(node)
{
    private static readonly Size DefaultSvgSize = new(300, 150);
    private readonly ISvgSvgElement svgNode = node;


    protected override Drawing? CreateDrawing(ParsingContext context)
    {
        var imgPart = AddImagePart(context.HostingPart, ImagePartType.Svg);
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(svgNode.OuterHtml), writable: false);
            imgPart.FeedData(stream);
        var imagePartId = context.HostingPart.GetIdOfPart(imgPart);

        return CreateSvgDrawing(context, svgNode, imagePartId, Size.Empty);
    }

    private static ImagePart AddImagePart(OpenXmlPartContainer hostingPart, PartTypeInfo partType)
    {
        return hostingPart switch {
            MainDocumentPart mp => mp.AddImagePart(partType),
            HeaderPart h => h.AddImagePart(partType),
            FooterPart f => f.AddImagePart(partType),
            _ => throw new NotSupportedException($"HostingPart was not recognized: {hostingPart.GetType().FullName}")
        };
    }

    internal static Drawing? CreateSvgDrawing(ParsingContext context, ISvgSvgElement svgNode, string imagePartId, Size preferredSize)
    {
        preferredSize = GuessSize(svgNode, preferredSize);

        long widthInEmus = new Unit(UnitMetric.Pixel, preferredSize.Width).ValueInEmus;
        long heightInEmus = new Unit(UnitMetric.Pixel, preferredSize.Height).ValueInEmus;

        var (imageObjId, drawingObjId) = IncrementDrawingObjId(context);

        string? title = svgNode.QuerySelector("title")?.TextContent?.CollapseAndStrip() ?? "Picture " + imageObjId;
        string? description = svgNode.QuerySelector("desc")?.TextContent?.CollapseAndStrip() ?? string.Empty;

        var img = new Drawing(
            new wp.Inline(
                new wp.Extent() { Cx = widthInEmus, Cy = heightInEmus },
                new wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new wp.DocProperties() { Id = drawingObjId, Name = title, Description = description },
                new wp.NonVisualGraphicFrameDrawingProperties
                {
                    GraphicFrameLocks = new a.GraphicFrameLocks() { NoChangeAspect = true }
                },
                new a.Graphic(
                    new a.GraphicData(
                        new pic.Picture(
                            new pic.NonVisualPictureProperties
                            {
                                NonVisualDrawingProperties = new pic.NonVisualDrawingProperties()
                                {
                                    Id = imageObjId,
                                    Name = title
                                },
                                NonVisualPictureDrawingProperties = new()
                            },
                            new pic.BlipFill(
                                new a.Blip(
                                    new a.BlipExtensionList(
                                        new a.BlipExtension(new SVGBlip { Embed = imagePartId })
                                        {
                                            Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
                                        })
                                )
                                { Embed = imagePartId /* ideally, that should be a png representation of the svg */ },
                                new a.Stretch(
                                    new a.FillRectangle())
                            ),
                            new pic.ShapeProperties(
                                new a.Transform2D(
                                    new a.Offset() { X = 0L, Y = 0L },
                                    new a.Extents() { Cx = widthInEmus, Cy = heightInEmus }),
                                new a.PresetGeometry(
                                    new a.AdjustValueList()
                                )
                                { Preset = a.ShapeTypeValues.Rectangle })
                        )
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            )
            { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U }
        );

        return img;
    }

    private static Size GuessSize(ISvgSvgElement svgNode, Size preferredSize)
    {
        if (!preferredSize.IsEmpty)
            return preferredSize;

        // % is not supported
        var width = Unit.Parse(svgNode.GetAttribute("width").AsSpan(), UnitMetric.Pixel);
        if (width.IsFixed && width.Value > 0)
        {
            preferredSize.Width = width.ValueInPx;
        }
        var height = Unit.Parse(svgNode.GetAttribute("height").AsSpan(), UnitMetric.Pixel);
        if (height.IsFixed && height.Value > 0)
        {
            preferredSize.Height = height.ValueInPx;
        }

        if (preferredSize.HasInvalidDimension &&
            CssShorthand.TryParse(svgNode.GetAttribute("viewBox").AsSpan(),
                p => Unit.Parse(p, UnitMetric.Pixel), p => p.IsValid,
                out (Unit, Unit, Unit, Unit) viewBox))
        {
            preferredSize.Width = viewBox.Item3.ValueInPx;
            preferredSize.Height = viewBox.Item4.ValueInPx;
        }

        if (preferredSize.HasInvalidDimension)
        {
            // If size is still empty, use default dimensions
            return DefaultSvgSize;
        }

        return preferredSize;
    }
}