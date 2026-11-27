/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
using System.Globalization;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Top parent expression, processing the <c>body</c> tag,
/// even if it is not directly specified in the provided Html.
/// </summary>
sealed class BodyExpression(IHtmlElement node, ParagraphStyleId? defaultStyle)
    : BlockElementExpression(node, defaultStyle)
{
    private bool shouldRegisterTopBookmark;


    public override IEnumerable<OpenXmlElement> Interpret(ParsingContext context)
    {
        MarBookmarkTargets();

        var elements = base.Interpret(context).ToList();

        if (shouldRegisterTopBookmark && elements.Count > 0)
        {
            var body = context.MainPart.Document!.Body!;

            // Check whether it already exists
            if (body.Descendants<BookmarkStart>().Any(b => b.Name?.Value == "_top"))
            {
                return elements;
            }

            var bookmarkId = IncrementBookmarkId(context).ToString(CultureInfo.InvariantCulture);
            // this is expected to stand in the 1st paragraph
            Paragraph? p = body.GetFirstChild<Paragraph>();
            p ??= body.PrependChild(new Paragraph());
            p.InsertAfter(new BookmarkEnd() { Id = bookmarkId }, p.ParagraphProperties);
            p.InsertAfter(new BookmarkStart() { Id = bookmarkId, Name = "_top" }, p.ParagraphProperties);
        }

        return elements;
    }

    protected override void ComposeStyles(ParsingContext context)
    {
        base.ComposeStyles(context);

        if (paraProperties.BiDi is not null)
        {
            var body = WordDocumentFactory.GetOrCreateDocumentBody(context.MainPart);
            var sectionProperties = body.GetLastChild<SectionProperties>();
            if (sectionProperties?.HasChild<BiDi>() == false)
            {
                sectionProperties.AddChild(paraProperties.BiDi.CloneNode(true));
            }
        }
    }

    /// <summary>
    /// Detect all bookmarks (in-document) and modify the DOM to mark the nodes for future processing.
    /// </summary>
    private void MarBookmarkTargets()
    {
        var links = node.QuerySelectorAll("a[href^='#']");
        if (links.Length == 0) return;

        foreach (var link in links.Cast<IHtmlAnchorElement>().Where(l => l.Hash.Length > 0))
        {
            if (link.IsTopAnchor())
            {
                shouldRegisterTopBookmark = true;
                continue;
            }

            var id = link.Hash.Substring(1);
            var target = node.Owner!.GetElementById(id);

            // `id` attribute is preferred but `name` is also valid
            target ??= node.Owner!.GetElementsByName(id).FirstOrDefault();

            if (target is null) continue;

            if (target.ParentElement is IHtmlHeadingElement heading)
                target = heading;

            // we will be able to retrieve the target during the processing
            target.Attributes.SetNamedItemWithNamespaceUri(
                new Attr("data-bookmark", id));
        }
    }
}