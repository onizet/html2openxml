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
using System.Threading;
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
    private readonly IHtmlImageElement imgNode = node;


    protected override Drawing? CreateDrawing(ParsingContext context)
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

        HtmlImageInfo? iinfo = context.ImageLoader.Download(src, CancellationToken.None)
            .ConfigureAwait(false).GetAwaiter().GetResult();

        if (iinfo == null)
            return null;

        if (iinfo.TypeInfo == ImagePartType.Svg)
        {
            var imagePart = context.HostingPart.GetPartById(iinfo.ImagePartId);
            using var stream = imagePart.GetStream(System.IO.FileMode.Open);
            using var sreader = new System.IO.StreamReader(stream);
            imgNode.Insert(AdjacentPosition.AfterBegin, sreader.ReadToEnd());

            var svgNode = imgNode.FindChild<ISvgSvgElement>();
            if (svgNode is null) return null;
            return SvgExpression.CreateSvgDrawing(context, svgNode, iinfo.ImagePartId, preferredSize);
        }

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

        var (imageObjId, drawingObjId) = IncrementDrawingObjId(context);
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

        return img;
    }
}