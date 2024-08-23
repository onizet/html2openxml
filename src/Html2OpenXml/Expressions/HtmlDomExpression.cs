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
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Represents the base processor of an HTML node (text, element, comment, processing instruction).
/// </summary>
abstract class HtmlDomExpression
{
    protected const string InternalNamespaceUri = "https://github.com/onizet/html2openxml";
    static readonly Dictionary<string, Func<IElement, HtmlDomExpression>> knownTags = InitKnownTags();
    static readonly HashSet<string> ignoreTags = new(StringComparer.OrdinalIgnoreCase) {
        TagNames.Xml, TagNames.AnnotationXml, TagNames.Button, TagNames.Progress,
        TagNames.Select, TagNames.Input, TagNames.Textarea, TagNames.Meter };

    private static Dictionary<string, Func<IElement, HtmlDomExpression>> InitKnownTags()
    {
        // A complete list of HTML tags can be found here: http://www.w3schools.com/tags/default.asp

        var knownTags = new Dictionary<string, Func<IElement, HtmlDomExpression>>(StringComparer.InvariantCultureIgnoreCase) {
            { TagNames.A, el => new HyperlinkExpression((IHtmlAnchorElement) el) },
            { TagNames.Abbr, el => new AbbreviationExpression((IHtmlElement) el) },
            { "acronym", el => new AbbreviationExpression((IHtmlElement) el) },
            { TagNames.B, el => new PhrasingElementExpression((IHtmlElement) el, new Bold()) },
            { TagNames.BlockQuote, el => new BlockQuoteExpression((IHtmlElement) el) },
            { TagNames.Br, _ => new LineBreakExpression() },
            { TagNames.Cite, el => new CiteElementExpression((IHtmlElement) el) },
            { TagNames.Dd, el => new BlockElementExpression((IHtmlElement) el, new Indentation() { FirstLine = "708" }, new SpacingBetweenLines() { After = "0" }) },
            { TagNames.Del, el => new PhrasingElementExpression((IHtmlElement) el, new Strike()) },
            { TagNames.Dfn, el => new AbbreviationExpression((IHtmlElement) el) },
            { TagNames.Em, el => new PhrasingElementExpression((IHtmlElement) el, new Italic()) },
            { TagNames.Figcaption, el => new FigureCaptionExpression((IHtmlElement) el) },
            { TagNames.Font, el => new FontElementExpression((IHtmlElement) el) },
            { TagNames.H1, el => new HeadingElementExpression((IHtmlElement) el) },
            { TagNames.H2, el => new HeadingElementExpression((IHtmlElement) el) },
            { TagNames.H3, el => new HeadingElementExpression((IHtmlElement) el) },
            { TagNames.H4, el => new HeadingElementExpression((IHtmlElement) el) },
            { TagNames.H5, el => new HeadingElementExpression((IHtmlElement) el) },
            { TagNames.H6, el => new HeadingElementExpression((IHtmlElement) el) },
            { TagNames.I, el => new PhrasingElementExpression((IHtmlElement) el, new Italic()) },
            { TagNames.Hr, el => new HorizontalLineExpression((IHtmlElement) el) },
            { TagNames.Img, el => new ImageExpression((IHtmlImageElement) el) },
            { TagNames.Ins, el => new PhrasingElementExpression((IHtmlElement) el, new Underline() { Val = UnderlineValues.Single }) },
            { TagNames.Ol, el => new ListExpression((IHtmlElement) el) },
            { TagNames.Pre, el => new PreElementExpression((IHtmlElement) el) },
            { TagNames.Q, el => new QuoteElementExpression((IHtmlElement) el) },
            { TagNames.Quote, el => new QuoteElementExpression((IHtmlElement) el) },
            { TagNames.Span, el => new PhrasingElementExpression((IHtmlElement) el) },
            { TagNames.S, el => new PhrasingElementExpression((IHtmlElement) el, new Strike()) },
            { TagNames.Strike, el => new PhrasingElementExpression((IHtmlElement) el, new Strike()) },
            { TagNames.Strong, el => new PhrasingElementExpression((IHtmlElement) el, new Bold()) },
            { TagNames.Sub, el => new PhrasingElementExpression((IHtmlElement) el, new VerticalTextAlignment() { Val = VerticalPositionValues.Subscript }) },
            { TagNames.Sup, el => new PhrasingElementExpression((IHtmlElement) el, new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript }) },
            { TagNames.Svg, el => new SvgExpression((AngleSharp.Svg.Dom.ISvgSvgElement) el) },
            { TagNames.Table, el => new TableExpression((IHtmlTableElement) el) },
            { TagNames.Time, el => new PhrasingElementExpression((IHtmlElement) el) },
            { TagNames.U, el => new PhrasingElementExpression((IHtmlElement) el, new Underline() { Val = UnderlineValues.Single }) },
            { TagNames.Ul, el => new ListExpression((IHtmlElement) el) },
        };

        return knownTags;
    }

    /// <summary>
    /// Process the interpretation of the Html node to its Word OpenXml equivalence.
    /// </summary>
    /// <param name="context">The parsing context.</param>
    public abstract IEnumerable<OpenXmlElement> Interpret (ParsingContext context);


    /// <summary>
    /// Create a new interpreter for the given html tag.
    /// </summary>
    public static HtmlDomExpression? CreateFromHtmlNode (INode node)
    {
        if (node.NodeType == NodeType.Text)
            return new TextExpression(node);
        else if (node.NodeType == NodeType.Element
            && !ignoreTags.Contains(node.NodeName))
        {
            if (knownTags.TryGetValue(node.NodeName, out Func<IElement, HtmlDomExpression>? handler))
                return handler((IElement) node);

            // fallback on the flow element which will cover all the semantic Html5 tags
            return new BlockElementExpression((IHtmlElement) node);
        }

        return null;
    }
}
