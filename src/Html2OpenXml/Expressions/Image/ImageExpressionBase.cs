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
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

using a = DocumentFormat.OpenXml.Drawing;
using pic = DocumentFormat.OpenXml.Drawing.Pictures;
using wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of an image.
/// </summary>
abstract class ImageExpressionBase(AngleSharp.Dom.IElement node)  : HtmlDomExpression
{
    private readonly RunProperties runProperties = new();
    private readonly ParagraphProperties paraProperties = new();
    // some style attributes, such as borders, will convert this node to a framed container
    private bool renderAsFramed;


    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret (ParsingContext context)
    {
        var drawing = CreateDrawing(context);

        if (drawing == null)
            return [];

        Run run = new(drawing);
        ComposeStyles();

        if (runProperties.HasChildren)
            run.RunProperties = runProperties;

        if (renderAsFramed)
            return [new Paragraph(paraProperties, run)];
        return [run];
    }

    private void ComposeStyles ()
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
            var borderWidth = Unit.Parse(node.GetAttribute("border"));
            if (borderWidth.IsValid)
            {
                border.Val = BorderValues.Single;
                border.Size = (uint) borderWidth.ValueInPx * 4;
            }
        }

        if (border.Val?.Equals(BorderValues.None) == false)
        {
            runProperties.Border = border;
        }

        // if the layout is not inline and both left and right are auto, image appears centered
        // https://developer.mozilla.org/en-US/docs/Web/CSS/margin-left
        var margin = styleAttributes.GetMargin("margin");
        if (margin.Left.Type == UnitMetric.Auto 
            && margin.Right.Type == UnitMetric.Auto
            && !AngleSharpExtensions.IsInlineLayout(styleAttributes["display"], "inline-block"))
        {
            paraProperties.Justification = new() { Val = JustificationValues.Center };
            renderAsFramed = true;
        }
    }

    /// <summary>
    /// Create the Drawing model referencing the iamge part.
    /// </summary>
    protected abstract Drawing? CreateDrawing(ParsingContext context);

    /// <summary>
    /// Resolve the next available <see cref="AbstractNum.AbstractNumberId"/> (they must be unique and ordered).
    /// </summary>
    internal static (uint imageObjId, uint drawingObjId) IncrementDrawingObjId(ParsingContext context)
    {
        var imageObjId = context.Properties<uint?>("imageObjId");
        var drawingObjId = context.Properties<uint?>("drawingObjId");
        if (!imageObjId.HasValue || !drawingObjId.HasValue)
        {
            drawingObjId ??= 1; // 1 is the minimum ID set by MS Office.
            imageObjId ??= 1;

            var mainPart = context.MainPart;
            foreach (var part in new[] { 
                mainPart.Document.Body!.Descendants<Drawing>(),
                mainPart.HeaderParts.Where(f => f.Header != null).SelectMany(f => f.Header.Descendants<Drawing>()),
                mainPart.FooterParts.Where(f => f.Footer != null).SelectMany(f => f.Footer.Descendants<Drawing>())
            })
            foreach (Drawing d in part)
            {
                wp.DocProperties? docProperties = null;
                pic.NonVisualPictureProperties? nvPr = null;

                if (d.Anchor != null)
                {
                    docProperties = d.Anchor.GetFirstChild<wp.DocProperties>();
                    nvPr = d.Anchor.GetFirstChild<a.Graphic>()?.GraphicData?.GetFirstChild<pic.Picture>()?.GetFirstChild<pic.NonVisualPictureProperties>();
                }
                else if (d.Inline != null)
                {
                    docProperties = d.Inline!.DocProperties;
                    nvPr = d.Inline!.Graphic?.GraphicData?.GetFirstChild<pic.NonVisualPictureProperties>();
                }

                if (docProperties?.Id != null && docProperties.Id.Value > drawingObjId)
                    drawingObjId = docProperties.Id.Value;

                if (nvPr != null && nvPr.NonVisualDrawingProperties?.Id?.Value > imageObjId)
                    imageObjId = nvPr.NonVisualDrawingProperties.Id;
            }
        }

        // In order to add images in the document, we need to assign an unique id
        // to each Drawing object. So we'll loop through all of the existing <wp:docPr> elements
        // to find the largest Id, then increment it for each new image.
        imageObjId++;
        drawingObjId++;
        context.Properties("drawingObjId", drawingObjId);
        context.Properties("imageObjId", imageObjId);
        return (imageObjId.Value, drawingObjId.Value);
    }
}