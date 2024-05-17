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
    static readonly Dictionary<string, Func<IHtmlElement, HtmlElementExpression>> knownTags = InitKnownTags();
    static readonly HashSet<string> ignoreTags = new(StringComparer.OrdinalIgnoreCase) {
        TagNames.Xml, TagNames.AnnotationXml, TagNames.Button, TagNames.Progress,
        TagNames.Select, TagNames.Input, TagNames.Textarea, TagNames.Meter };

    private static Dictionary<string, Func<IHtmlElement, HtmlElementExpression>> InitKnownTags()
    {
        // A complete list of HTML tags can be found here: http://www.w3schools.com/tags/default.asp

        var knownTags = new Dictionary<string, Func<IHtmlElement, HtmlElementExpression>>(StringComparer.InvariantCultureIgnoreCase) {
            { TagNames.A, el => new HyperlinkExpression(el) },
            { TagNames.Abbr, el => new BlockQuoteExpression(el) },
            { "acronym", el => new BlockQuoteExpression(el) },
            { TagNames.B, el => new PhrasingElementExpression(el, new Bold()) },
            { TagNames.BlockQuote, el => new BlockQuoteExpression(el) },
            { TagNames.Br, el => new LineBreakExpression(el) },
            { TagNames.Cite, el => new CiteElementExpression(el) },
            //{ TagNames.Dl, el => new DefinitionListExpression(el) },
            { TagNames.Del, el => new PhrasingElementExpression(el, new Strike()) },
            { TagNames.Dfn, el => new BlockQuoteExpression(el) },
            { TagNames.Em, el => new PhrasingElementExpression(el, new Italic()) },
            { TagNames.Figcaption, el => new FigureCaptionExpression(el) },
            { TagNames.Font, el => new FontElementExpression(el) },
            { TagNames.H1, el => new HeadingElementExpression(el) },
            { TagNames.H2, el => new HeadingElementExpression(el) },
            { TagNames.H3, el => new HeadingElementExpression(el) },
            { TagNames.H4, el => new HeadingElementExpression(el) },
            { TagNames.H5, el => new HeadingElementExpression(el) },
            { TagNames.H6, el => new HeadingElementExpression(el) },
            { TagNames.I, el => new PhrasingElementExpression(el, new Italic()) },
            { TagNames.Hr, el => new HorizontalLineExpression(el) },
            { TagNames.Img, el => new ImageExpression(el) },
            { TagNames.Ins, el => new PhrasingElementExpression(el, new Underline() { Val = UnderlineValues.Single }) },
            { TagNames.Ol, el => new ListExpression(el) },
            { TagNames.Pre, el => new PreElementExpression(el) },
            { TagNames.Q, el => new QuoteElementExpression(el) },
            { TagNames.Quote, el => new QuoteElementExpression(el) },
            { TagNames.Table, el => new TableExpression(el) },
            //{ TagNames.Caption, TableCaption },
            { TagNames.Span, el => new PhrasingElementExpression(el) },
            { TagNames.S, el => new PhrasingElementExpression(el, new Strike()) },
            { TagNames.Strike, el => new PhrasingElementExpression(el, new Strike()) },
            { TagNames.Strong, el => new PhrasingElementExpression(el, new Bold()) },
            { TagNames.Sub, el => new PhrasingElementExpression(el, new VerticalTextAlignment() { Val = VerticalPositionValues.Subscript }) },
            { TagNames.Sup, el => new PhrasingElementExpression(el, new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript }) },
            { TagNames.U, el => new PhrasingElementExpression(el, new Underline() { Val = UnderlineValues.Single }) },
            { TagNames.Ul, el => new ListExpression(el) },
        };

        return knownTags;
    }

    /// <summary>
    /// Process the interpretation of the Html node to its Word OpenXml equivalence.
    /// </summary>
    /// <param name="context">The parsing context.</param>
    public abstract IEnumerable<OpenXmlCompositeElement> Interpret (ParsingContext context);


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
            if (knownTags.TryGetValue(node.NodeName, out Func<IHtmlElement, HtmlElementExpression>? handler))
                return handler((IHtmlElement) node);

            // fallback on the flow element which will cover all the semantic Html5 tags
            return new BlockElementExpression((IHtmlElement) node);
        }

        return null;
    }
}
