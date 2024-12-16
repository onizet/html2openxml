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
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

using a = DocumentFormat.OpenXml.Drawing;
using pic = DocumentFormat.OpenXml.Drawing.Pictures;

/// <summary>
/// Process the parsing of a link element.
/// </summary>
sealed class HyperlinkExpression(IHtmlAnchorElement node) : PhrasingElementExpression(node)
{
    private readonly IHtmlAnchorElement linkNode = node;


    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret (ParsingContext context)
    {
        var h = CreateHyperlink(context);
        var childElements = Interpret(context.CreateChild(this), linkNode.ChildNodes);
        if (h is null)
        {
            return childElements;
        }

        // Let's see whether the link tag include an image inside its body.
        // If so, the Hyperlink OpenXmlElement is lost and we'll keep only the images
        // and applied a HyperlinkOnClick attribute.
        IEnumerable<OpenXmlElement> imagesInLink;
        // Clickable image is only supported in body but not in header/footer
        if (context.HostingPart is MainDocumentPart &&
            (imagesInLink = childElements.Where(e => e.HasChild<Drawing>())).Any())
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

                d.Inline ??= new a.Wordprocessing.Inline();
                d.Inline.DocProperties ??= new a.Wordprocessing.DocProperties();

                if (h.Anchor == "_top")
                {
                    // exception case: clickable image requires the _top bookmark get registred with a relationship
                    var extLink = context.HostingPart.AddHyperlinkRelationship(new Uri("#_top", UriKind.Relative), false);
                    d.Inline.DocProperties.Append(
                        new a.HyperlinkOnClick() { Id = extLink.Id, Tooltip = alt });
                }
                else
                {
                    d.Inline.DocProperties.Append(
                        new a.HyperlinkOnClick() { Id = h.Id ?? h.Anchor, Tooltip = alt });
                }
            }
        }

        // can't use GetFirstChild<Run> or we may find the one containing the image
        List<Run> runs = [];
        foreach (var el in childElements)
        {
            if (el is Run r) runs.Add(r);
            // unroll paragraphs. CloneNode is need to unparent the run
            else runs.AddRange(el.Elements<Run>().Select(r => (Run) r.CloneNode(true)));
        }

        foreach (var run in runs.Where(run => !run.HasChild<Drawing>()))
        {
            run.RunProperties ??= new();
            run.RunProperties.RunStyle = context.DocumentStyle.GetRunStyle(
                    context.DocumentStyle.DefaultStyles.HyperlinkStyle);
        }

        // Append the processed elements and put them to the Run of the Hyperlink
        h.Append(runs);

        return [h];
    }

    private Hyperlink? CreateHyperlink(ParsingContext context)
    {
        string? att = linkNode.GetAttribute("href");
        Hyperlink? h = null;

        if (string.IsNullOrEmpty(att))
            return null;

        // Always accept _top anchor
        if (linkNode.IsTopAnchor())
        {
            h = new Hyperlink() { History = true, Anchor = "_top" };
        }
        // is it an anchor?
        else if (context.Converter.SupportsAnchorLinks && linkNode.Hash.Length > 1 && linkNode.Hash[0] == '#')
        {
            h = new Hyperlink(
                ) { History = true, Anchor = linkNode.Hash.Substring(1) };
        }
        // ensure the links does not start with javascript:
        else if (AngleSharpExtensions.TryParseUrl(att, UriKind.Absolute, out var uri))
        {
            var extLink = context.HostingPart.AddHyperlinkRelationship(uri!, true);

            h = new Hyperlink(
                ) { History = true, Id = extLink.Id };
        }

        if (h == null)
        {
            // link to a broken url, simply process the content of the tag
            return null;
        }

        if (!string.IsNullOrEmpty(linkNode.Title))
            h.Tooltip = linkNode.Title;
        return h;
    }
}