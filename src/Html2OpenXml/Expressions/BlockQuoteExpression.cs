/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of <c>blockquote</c>.
/// </summary>
sealed class BlockQuoteExpression(IHtmlElement node) : BlockElementExpression(node)
{

    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret(ParsingContext context)
    {
        var childElements = base.Interpret(context).ToList();
        if (childElements.Count == 0)
            return [];

        // Footnote or endnote are invalid inside header and footer
        if (context.HostingPart is not MainDocumentPart)
            return childElements;

        // Transform the inline acronym/abbreviation to a reference to a foot note.
        if (childElements[0] is Paragraph paragraph)
        {
            string? description = node.GetAttribute("cite");

            paragraph.ParagraphProperties ??= new();
            if (paragraph.ParagraphProperties.ParagraphStyleId is null)
                paragraph.ParagraphProperties.ParagraphStyleId = 
                    context.DocumentStyle.GetParagraphStyle(context.DocumentStyle.DefaultStyles.IntenseQuoteStyle);

            CascadeStyles(paragraph);

            if (!string.IsNullOrEmpty(description))
            {
                string runStyle;
                FootnoteEndnoteReferenceType reference;

                if (context.Converter.AcronymPosition == AcronymPosition.PageEnd)
                {
                    reference = new FootnoteReference() { Id = AbbreviationExpression.AddFootnoteReference(context, description!) };
                    runStyle = context.DocumentStyle.DefaultStyles.FootnoteReferenceStyle;
                }
                else
                {
                    reference = new EndnoteReference() { Id = AbbreviationExpression.AddEndnoteReference(context, description!) };
                    runStyle = context.DocumentStyle.DefaultStyles.EndnoteReferenceStyle;
                }

                paragraph.AppendChild(new Run(reference) {
                    RunProperties = new() {
                        RunStyle = context.DocumentStyle.GetRunStyle(runStyle) }
                    });
            }
        }

        return childElements;
    }
}