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
using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

using a = DocumentFormat.OpenXml.Drawing;
using pic = DocumentFormat.OpenXml.Drawing.Pictures;

/// <summary>
/// Process the parsing of a link element.
/// </summary>
sealed class HyperlinkExpression(IHtmlElement node) : PhrasingElementExpression(node)
{
    private readonly IHtmlAnchorElement linkNode = (IHtmlAnchorElement) node;


    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret (ParsingContext context)
    {
        var h = CreateHyperlink(context);
        var childElements = Interpret(context.CreateChild(this), node.ChildNodes);
        if (h is null)
        {
            return childElements;
        }

        // Let's see whether the link tag include an image inside its body.
        // If so, the Hyperlink OpenXmlElement is lost and we'll keep only the images
        // and applied a HyperlinkOnClick attribute.
        var imagesInLink = childElements.Where(e => e.HasChild<Drawing>());
        if (imagesInLink.Any())
        {
            foreach (var img in imagesInLink)
            {
                // Retrieves the "alt" attribute of the image and apply it as the link's tooltip
                Drawing? d = img.GetFirstChild<Drawing>();
                if (d == null) continue;

                var enDp = d.Descendants<pic.NonVisualDrawingProperties>().GetEnumerator();
                string? alt;
                if (enDp.MoveNext()) alt = enDp.Current.Description;
                else alt = null;

                d.InsertInDocProperties(
                    new a.HyperlinkOnClick() { Id = h.Id ?? h.Anchor, Tooltip = alt });
            }
        }

        // can't use GetFirstChild<Run> or we may find the one containing the image
        foreach (var el in childElements)
        {
            if (el is Run run && !run.HasChild<Drawing>())
            {
                run.InsertInProperties(prop =>
                    prop.RunStyle = context.DocumentStyle.GetRunStyle(
                            context.DocumentStyle.DefaultStyles.HyperlinkStyle)
                );
                break;
            }
        }

        // Append the processed elements and put them to the Run of the Hyperlink
        h.Append(childElements);

        return [new Paragraph(h)];
    }

    private Hyperlink? CreateHyperlink(ParsingContext context)
    {
        string? att = linkNode.GetAttribute("href");
        Hyperlink? h = null;

        if (string.IsNullOrEmpty(att))
            return null;

        // is it an anchor?
        if (att![0] == '#' && att.Length > 1)
        {
            // Always accept _top anchor
            if (!context.Converter.ExcludeLinkAnchor || att == "#_top")
            {
                h = new Hyperlink(
                    ) { History = true, Anchor = att.Substring(1) };
            }
        }
        // ensure the links does not start with javascript:
        else if (AngleSharpExtensions.TryParseUrl(att, UriKind.Absolute, out var uri))
        {
            var extLink = context.MainPart.AddHyperlinkRelationship(uri!, true);

            h = new Hyperlink(
                ) { History = true, Id = extLink.Id };
        }

        if (h == null)
        {
            // link to a broken url, simply process the content of the tag
            return null;
        }

        if (!string.IsNullOrEmpty(node.Title))
            h.Tooltip = node.Title;
        return h;
    }
}