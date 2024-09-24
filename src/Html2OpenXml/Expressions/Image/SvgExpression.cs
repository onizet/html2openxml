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
using AngleSharp.Svg.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using System.Text;

using a = DocumentFormat.OpenXml.Drawing;
using pic = DocumentFormat.OpenXml.Drawing.Pictures;
using wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using AngleSharp.Text;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of a <c>svg</c> element.
/// </summary>
sealed class SvgExpression(ISvgSvgElement node) : ImageExpressionBase(node)
{
    private readonly ISvgSvgElement svgNode = node;


    protected override Drawing? CreateDrawing(ParsingContext context)
    {
        var imgPart = context.MainPart.AddImagePart(ImagePartType.Svg);
        using var stream = new System.IO.MemoryStream(Encoding.UTF8.GetBytes(svgNode.OuterHtml), writable: false);
            imgPart.FeedData(stream);
        var imagePartId = context.MainPart.GetIdOfPart(imgPart);
        return CreateSvgDrawing(context, svgNode, imagePartId, Size.Empty);
    }

    internal static Drawing CreateSvgDrawing(ParsingContext context, ISvgSvgElement svgNode, string imagePartId, Size preferredSize)
    {
        var width = Unit.Parse(svgNode.GetAttribute("width"));
        var height = Unit.Parse(svgNode.GetAttribute("height"));
        long widthInEmus, heightInEmus;
        if (width.IsValid && height.IsValid)
        {
            widthInEmus = width.ValueInEmus;
            heightInEmus = height.ValueInEmus;
        }
        else
        {
            widthInEmus = new Unit(UnitMetric.Pixel, preferredSize.Width).ValueInEmus;
            heightInEmus = new Unit(UnitMetric.Pixel, preferredSize.Height).ValueInEmus;
        }

        var (imageObjId, drawingObjId) = IncrementDrawingObjId(context);

        string? title = svgNode.QuerySelector("title")?.TextContent?.CollapseAndStrip() ?? "Picture " + imageObjId;
        string? description = svgNode.QuerySelector("desc")?.TextContent?.CollapseAndStrip() ?? string.Empty;

        var img = new Drawing(
            new wp.Inline(
                new wp.Extent() { Cx = widthInEmus, Cy = heightInEmus },
                new wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new wp.DocProperties() { Id = drawingObjId, Name = title, Description = description },
                new wp.NonVisualGraphicFrameDrawingProperties {
                    GraphicFrameLocks = new a.GraphicFrameLocks() { NoChangeAspect = true }
                },
                new a.Graphic(
                    new a.GraphicData(
                        new pic.Picture(
                            new pic.NonVisualPictureProperties {
                                NonVisualDrawingProperties = new pic.NonVisualDrawingProperties() {
                                    Id = imageObjId, Name = title
                                },
                                NonVisualPictureDrawingProperties = new()
                            },
                            new pic.BlipFill(
                                new a.Blip(
                                    new a.BlipExtensionList(
                                        new a.BlipExtension(new SVGBlip { Embed = imagePartId }) {
                                            Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
                                        })
                                ) { Embed = imagePartId /* ideally, that should be a png representation of the svg */ },
                                new a.Stretch(
                                    new a.FillRectangle())
                            ),
                            new pic.ShapeProperties(
                                new a.Transform2D(
                                    new a.Offset() { X = 0L, Y = 0L },
                                    new a.Extents() { Cx = widthInEmus, Cy = heightInEmus }),
                                new a.PresetGeometry(
                                    new a.AdjustValueList()
                                ) { Preset = a.ShapeTypeValues.Rectangle })
                        )
                    ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            ) { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U }
        );

        return img;
    }
}