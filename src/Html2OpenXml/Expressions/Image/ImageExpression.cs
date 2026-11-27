/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Svg.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml.IO;

using a = DocumentFormat.OpenXml.Drawing;
using pic = DocumentFormat.OpenXml.Drawing.Pictures;
using wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of an image.
/// </summary>
class ImageExpression(IHtmlImageElement node) : ImageExpressionBase(node)
{
    private static readonly Size DefaultImgSize = new(300, 200);
    private readonly IHtmlImageElement imgNode = node;


    protected override Drawing? CreateDrawing(ParsingContext context)
    {
        string? src = imgNode.GetAttribute("src");

        if (src == null)
            return null;

        // Bug reported by Erik2014. Inline 64 bit images can be too big and Uri.TryCreate will fail silently with a SizeLimit error.
        // To circumvent this buffer size, we will work either on the Uri, either on the original src.
        var isDataUri = DataUri.IsDataUri(src);
        if (!isDataUri &&
            !AngleSharpExtensions.TryParseUrl(src, UriKind.RelativeOrAbsolute, out var _))
        {
            return null;
        }

        string alt = imgNode.Title ?? imgNode.AlternativeText ?? string.Empty;
        HtmlImageInfo? iinfo = context.ImageLoader.Download(src, CancellationToken.None)
            .ConfigureAwait(false).GetAwaiter().GetResult();

        if (iinfo == null)
            return null;

        var (widthInEmus, heightInEmus) = CalculateSize(context, iinfo);

        if (iinfo.TypeInfo == ImagePartType.Svg)
        {
            var imagePart = context.HostingPart.GetPartById(iinfo.ImagePartId);
            using var stream = imagePart.GetStream(System.IO.FileMode.Open);
            using var sreader = new System.IO.StreamReader(stream);
            imgNode.Insert(AdjacentPosition.AfterBegin, sreader.ReadToEnd());

            var svgNode = imgNode.FindChild<ISvgSvgElement>();
            if (svgNode is null) return null;
            return SvgExpression.CreateSvgDrawing(context, svgNode, iinfo.ImagePartId,
                widthInEmus, heightInEmus);
        }

        var (imageObjId, drawingObjId) = IncrementDrawingObjId(context);
        var img = new Drawing(
            new wp.Inline(
                new wp.Extent() { Cx = widthInEmus, Cy = heightInEmus },
                new wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new wp.DocProperties() { Id = drawingObjId, Name = "Picture " + imageObjId, Description = string.Empty },
                new wp.NonVisualGraphicFrameDrawingProperties
                {
                    GraphicFrameLocks = new a.GraphicFrameLocks() { NoChangeAspect = true }
                },
                new a.Graphic(
                    new a.GraphicData(
                        new pic.Picture(
                            new pic.NonVisualPictureProperties
                            {
                                NonVisualDrawingProperties = new pic.NonVisualDrawingProperties() {
                                    Id = imageObjId,
                                    Name = isDataUri ? string.Empty : src,
                                    Description = alt
                                },
                                NonVisualPictureDrawingProperties = new pic.NonVisualPictureDrawingProperties(
                                    new a.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true })
                            },
                            new pic.BlipFill(
                                CreateBlip(iinfo),
                                new a.SourceRectangle(),
                                new a.Stretch(
                                    new a.FillRectangle())),
                            new pic.ShapeProperties(
                                new a.Transform2D(
                                    new a.Offset() { X = 0L, Y = 0L },
                                    new a.Extents() { Cx = widthInEmus, Cy = heightInEmus }),
                                new a.PresetGeometry(
                                    new a.AdjustValueList()
                                )
                                { Preset = a.ShapeTypeValues.Rectangle }
                            )
                            { BlackWhiteMode = a.BlackWhiteModeValues.Auto })
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            )
            { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U }
        );

        return img;
    }

    private (long WidthInEmus, long HeightInEmus) CalculateSize(ParsingContext context, HtmlImageInfo imageInfo)
    {
        Unit preferredWidth = Unit.Empty;
        Unit preferredHeight = Unit.Empty;
        Unit maxContentWidth = context.GetAvailableContentWidth();

        // width and height attribute are pixels according to HTML spec.
        if (imgNode.DisplayWidth > 0)
        {
            preferredWidth = new Unit(UnitMetric.Pixel, imgNode.DisplayWidth);
        }
        if (imgNode.DisplayHeight > 0)
        {
            preferredHeight = new Unit(UnitMetric.Pixel, imgNode.DisplayHeight);
        }

        HtmlAttributeCollection? styles = null;
        if (preferredWidth.Value == 0)
        {
            styles = imgNode.GetStyles();
            preferredWidth = GetDimension(styles, "width", "max-width", () => maxContentWidth);
        }
        if (preferredHeight.Value == 0)
        {
            styles ??= imgNode.GetStyles();
            preferredHeight = GetDimension(styles, "height", "max-height", () => context.GetAvailableContentHeight());
        }

        // Intrinsic image size
        Size actualSize = imageInfo.Size;

        // Resolve missing dimension from image aspect ratio
        // Image perspective skewed. Bug fixed by ddeforge on github.com/onizet/html2openxml/discussions/350500
        if (preferredWidth.Value <= 0 || preferredHeight.Value <= 0)
        {
            // Image perspective skewed. Bug fixed by ddeforge on github.com/onizet/html2openxml/discussions/350500
            (preferredWidth, preferredHeight) = ImageHeader.KeepAspectRatio(actualSize, preferredWidth.ValueInEmus, preferredHeight.ValueInEmus);
        }

        // Fallback for unknown image sizes (external image, malformed image...)
        if (preferredWidth.ValueInEmus <= 0 || preferredHeight.ValueInEmus <= 0)
        {
            // Use default size for external images or when size cannot be determined
            // Default to a reasonable size (similar to how browsers handle images with unknown dimensions)
            preferredWidth = new Unit(UnitMetric.Pixel, DefaultImgSize.Width);
            preferredHeight = new Unit(UnitMetric.Pixel, DefaultImgSize.Height);
        }

        long widthInEmus = preferredWidth.ValueInEmus;
        long heightInEmus = preferredHeight.ValueInEmus;

        // Prevent overflow
        if (widthInEmus > maxContentWidth.ValueInEmus)
        {
            double scale = (double) maxContentWidth.ValueInEmus / widthInEmus;

            widthInEmus = maxContentWidth.ValueInEmus;
            heightInEmus = (long) Math.Round(heightInEmus * scale);
        }

        return (widthInEmus, heightInEmus);
    }

    private static Unit GetDimension(HtmlAttributeCollection styles,
        string primaryStyle, string fallbackStyle, Func<Unit> availableContainerWidthInTwips)
    {
        var unit = styles.GetUnit(primaryStyle);
        if (!unit.IsValid)
        {
            unit = styles.GetUnit(fallbackStyle);
        }

        if (!unit.IsValid)
            return Unit.Empty;

        if (unit.Metric == UnitMetric.Percent)
        {
            var maxValue = availableContainerWidthInTwips().Value;
            return new Unit(UnitMetric.Twips, unit.Value * maxValue / 100);
        }

        return unit;
    }

    /// <summary>
    /// Creates a Blip element with either an embedded or external image reference.
    /// </summary>
    private static a.Blip CreateBlip(HtmlImageInfo iinfo)
    {
        if (iinfo.IsExternal)
        {
            // Use Link property for external images
            return new a.Blip() { Link = iinfo.ImagePartId };
        }
        else
        {
            // Use Embed property for embedded images (default behaviour)
            return new a.Blip() { Embed = iinfo.ImagePartId };
        }
    }
}
