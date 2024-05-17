/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml - All Rights Reserved
 * 
 * This source is subject to the Microsoft Permissive License.
 * Please see the License.txt file for more information.
 * All other rights reserved.
 * 
 * THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY 
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
 * PARTICULAR PURPOSE.
 */
using System;
using System.Collections.Generic;
using System.Threading;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml.IO;

using a = DocumentFormat.OpenXml.Drawing;
using pic = DocumentFormat.OpenXml.Drawing.Pictures;
using wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of an image.
/// </summary>
sealed class ImageExpression(IHtmlElement node) : HtmlElementExpression(node)
{
    private readonly IHtmlImageElement imgNode = (IHtmlImageElement) node;


    /// <inheritdoc/>
    public override IEnumerable<OpenXmlCompositeElement> Interpret (ParsingContext context)
    {
        var drawing = CreateDrawing(context);

        if (drawing == null)
            return [];

        Run run = new(drawing);
        Border border = ComposeStyles();
        if (border.Val?.Equals(BorderValues.None) == false)
            run.InsertInProperties(prop => prop.Border = border);
        return [run];
    }

    public override void CascadeStyles(OpenXmlCompositeElement element)
    {
        throw new NotSupportedException();
    }

    private Border ComposeStyles ()
    {
        var styleAttributes = node.GetStyles();
        var border = new Border() { Val = BorderValues.None };

        // OpenXml limits the border to 4-side of the same color and style.
        SideBorder styleBorder = styleAttributes.GetSideBorder("border");
        if (styleBorder.IsValid)
        {
            border.Val = styleBorder.Style;
            border.Color = styleBorder.Color.ToHexString();
            border.Size = (uint) styleBorder.Width.ValueInPx * 4;
        }
        else
        {
            var borderWidth = Unit.Parse(imgNode.GetAttribute("border"));
            if (borderWidth.IsValid)
            {
                border.Val = BorderValues.Single;
                border.Size = (uint) borderWidth.ValueInPx * 4;
            }
        }
        return border;
    }

    private Drawing? CreateDrawing(ParsingContext context)
    {
        string? src = imgNode.GetAttribute("src");

        // Bug reported by Erik2014. Inline 64 bit images can be too big and Uri.TryCreate will fail silently with a SizeLimit error.
        // To circumvent this buffer size, we will work either on the Uri, either on the original src.
        if (src == null || 
            (!DataUri.IsWellFormed(src) && !AngleSharpExtensions.TryParseUrl(src, UriKind.RelativeOrAbsolute, out var _)))
        {
            return null;
        }

        string alt = imgNode.Title ?? imgNode.AlternativeText ?? string.Empty;

        Size preferredSize = Size.Empty;

        // % is not supported
        if (imgNode.DisplayWidth > 0)
        {
            preferredSize.Width = imgNode.DisplayWidth;
        }
        if (imgNode.DisplayHeight > 0)
        {
            // Image perspective skewed. Bug fixed by ddeforge on github.com/onizet/html2openxml/discussions/350500
            preferredSize.Height = imgNode.DisplayHeight;
        }

        var imageObjId = context.Properties<uint?>("imageObjId");
        var drawingObjId = context.Properties<uint?>("drawingObjId");
        if (!imageObjId.HasValue)
        {
            // In order to add images in the document, we need to asisgn an unique id
            // to each Drawing object. So we'll loop through all of the existing <wp:docPr> elements
            // to find the largest Id, then increment it for each new image.

            drawingObjId = 1; // 1 is the minimum ID set by MS Office.
            imageObjId = 1;
            foreach (var d in context.MainPart.Document.Body!.Descendants<Drawing>())
            {
                if (d.Inline == null) continue; // fix some rare issue where Inline is null (reported by scwebgroup)
                if (d.Inline!.DocProperties?.Id?.Value > drawingObjId) drawingObjId = d.Inline.DocProperties.Id;

                var nvPr = d.Inline!.Graphic?.GraphicData?.GetFirstChild<pic.NonVisualPictureProperties>();
                if (nvPr != null && nvPr.NonVisualDrawingProperties?.Id?.Value > imageObjId)
                    imageObjId = nvPr.NonVisualDrawingProperties.Id;
            }
            if (drawingObjId > 1) drawingObjId++;
            if (imageObjId > 1) imageObjId++;
        }

        HtmlImageInfo? iinfo = context.Converter.ImagePrefetcher.Download(src, CancellationToken.None)
            .ConfigureAwait(false).GetAwaiter().GetResult();

        if (iinfo == null)
            return null;

        if (preferredSize.IsEmpty)
        {
            preferredSize = iinfo.Size;
        }
        else if (preferredSize.Width <= 0 || preferredSize.Height <= 0)
        {
            Size actualSize = iinfo.Size;
            preferredSize = ImageHeader.KeepAspectRatio(actualSize, preferredSize);
        }

        long widthInEmus = new Unit(UnitMetric.Pixel, preferredSize.Width).ValueInEmus;
        long heightInEmus = new Unit(UnitMetric.Pixel, preferredSize.Height).ValueInEmus;

        ++drawingObjId;
        ++imageObjId;

        var img = new Drawing(
            new wp.Inline(
                new wp.Extent() { Cx = widthInEmus, Cy = heightInEmus },
                new wp.EffectExtent() { LeftEdge = 19050L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new wp.DocProperties() { Id = drawingObjId, Name = "Picture " + imageObjId, Description = string.Empty },
                new wp.NonVisualGraphicFrameDrawingProperties {
                    GraphicFrameLocks = new a.GraphicFrameLocks() { NoChangeAspect = true }
                },
                new a.Graphic(
                    new a.GraphicData(
                        new pic.Picture(
                            new pic.NonVisualPictureProperties {
                                NonVisualDrawingProperties = new pic.NonVisualDrawingProperties() { 
                                    Id = imageObjId,
                                    Name = DataUri.IsWellFormed(src) ? string.Empty : src,
                                    Description = alt },
                                NonVisualPictureDrawingProperties = new pic.NonVisualPictureDrawingProperties(
                                    new a.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true })
                            },
                            new pic.BlipFill(
                                new a.Blip() { Embed = iinfo.ImagePartId },
                                new a.SourceRectangle(),
                                new a.Stretch(
                                    new a.FillRectangle())),
                            new pic.ShapeProperties(
                                new a.Transform2D(
                                    new a.Offset() { X = 0L, Y = 0L },
                                    new a.Extents() { Cx = widthInEmus, Cy = heightInEmus }),
                                new a.PresetGeometry(
                                    new a.AdjustValueList()
                                ) { Preset = a.ShapeTypeValues.Rectangle }
                            ) { BlackWhiteMode = a.BlackWhiteModeValues.Auto })
                    ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            ) { DistanceFromTop = (UInt32Value) 0U, DistanceFromBottom = (UInt32Value) 0U, DistanceFromLeft = (UInt32Value) 0U, DistanceFromRight = (UInt32Value) 0U }
        );

        context.Properties("imageObjId", imageObjId);
        context.Properties("drawingObjId", drawingObjId!);

        return img;
    }
}