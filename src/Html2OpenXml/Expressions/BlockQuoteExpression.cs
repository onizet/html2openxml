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
        var childElements = base.Interpret(context);
        if (!childElements.Any())
            return [];
 
        // Footnote or endnote are invalid inside header and footer
        if (context.HostingPart is not MainDocumentPart)
            return childElements;

        // Transform the inline acronym/abbreviation to a reference to a foot note.
        if (childElements.First() is Paragraph paragraph)
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