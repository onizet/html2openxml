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
using System.Text.RegularExpressions;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of <c>abbr</c>, <c>acronym</c>.
/// </summary>
sealed class AbbreviationExpression(IHtmlElement node) : PhrasingElementExpression(node)
{

    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret(ParsingContext context)
    {
        var childElements = base.Interpret(context);

        // Transform the inline acronym/abbreviation to a reference to a foot note.
        // Footnote or endnote are invalid inside header and footer
        string? description = node.Title;
        if (string.IsNullOrEmpty(description) || context.HostingPart is not MainDocumentPart)
            return childElements;

        string runStyle;
        FootnoteEndnoteReferenceType reference;

        if (context.Converter.AcronymPosition == AcronymPosition.PageEnd)
        {
            reference = new FootnoteReference() { Id = AddFootnoteReference(context, description!) };
            runStyle = context.DocumentStyle.DefaultStyles.FootnoteReferenceStyle;
        }
        else
        {
            reference = new EndnoteReference() { Id = AddEndnoteReference(context, description!) };
            runStyle = context.DocumentStyle.DefaultStyles.EndnoteReferenceStyle;
        }

        return childElements.Append(new Run(reference) {
            RunProperties = new() {
                RunStyle = context.DocumentStyle.GetRunStyle(runStyle) }
            });
    }

    /// <summary>
    /// Add a note to the FootNotes part and ensure it exists.
    /// </summary>
    /// <param name="context">The parsing context.</param>
    /// <param name="description">The description of an acronym, abbreviation, some book references, ...</param>
    /// <returns>Returns the id of the footnote reference.</returns>
    public static long AddFootnoteReference(ParsingContext context, string description)
    {
        FootnotesPart? fpart = context.MainPart.FootnotesPart ?? context.MainPart.AddNewPart<FootnotesPart>();
        var footnotesRef = context.Properties<long?>("footnotesRef");


        if (footnotesRef.HasValue)
        {
            footnotesRef++;
        }
        else if (fpart.Footnotes == null)
        {
            // Insert a new Footnotes reference
            new Footnotes(
                new Footnote(
                    new Paragraph(
                        new ParagraphProperties {
                            SpacingBetweenLines = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }
                        },
                        new Run(
                            new SeparatorMark())
                    )
                ) { Type = FootnoteEndnoteValues.Separator, Id = -1 },
                new Footnote(
                    new Paragraph(
                        new ParagraphProperties {
                            SpacingBetweenLines = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }
                        },
                        new Run(
                            new ContinuationSeparatorMark())
                    )
                ) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 }).Save(fpart);
            footnotesRef = 1;
        }
        else
        {
            // The footnotesRef Id is a required field and should be unique. You can assign yourself some hard-coded
            // value but that's absolutely not safe. We will loop through the existing Footnote
            // to retrieve the highest Id.
            footnotesRef = 0;
            foreach (var fn in fpart.Footnotes.Elements<Footnote>())
            {
                if (fn.Id != null && fn.Id > footnotesRef) footnotesRef = fn.Id.Value;
            }
            footnotesRef++;
        }


        Paragraph p;
        fpart.Footnotes!.Append(
            new Footnote(
                p = new Paragraph(
                    new ParagraphProperties {
                        ParagraphStyleId = context.DocumentStyle.GetParagraphStyle(context.DocumentStyle.DefaultStyles.FootnoteTextStyle)
                    },
                    new Run(
                        new RunProperties {
                            RunStyle = context.DocumentStyle.GetRunStyle(context.DocumentStyle.DefaultStyles.FootnoteReferenceStyle)
                        },
                        new FootnoteReferenceMark()),
                    new Run(
                    // Word insert automatically a space before the definition to separate the
                    // reference number with its description
                        new Text(" ") { Space = SpaceProcessingModeValues.Preserve })
                )
            ) { Id = footnotesRef });


        // Description in footnote reference can be plain text or a web protocols/file share (like \\server01)
        Regex linkRegex = new(@"^((https?|ftps?|mailto|file)://|[\\]{2})(?:[\w][\w.-]?)");
        if (linkRegex.IsMatch(description) && Uri.TryCreate(description, UriKind.Absolute, out var uriReference))
        {
            // when URI references a network server (ex: \\server01), System.IO.Packaging is not resolving the correct URI and this leads
            // to a bad-formed XML not recognized by Word. To enforce the "original URI", a fresh new instance must be created
            uriReference = new Uri(uriReference.AbsoluteUri, UriKind.Absolute);
            HyperlinkRelationship extLink = fpart.AddHyperlinkRelationship(uriReference, true);
            var h = new Hyperlink(
                ) { History = true, Id = extLink.Id };

            h.Append(new Run(
                new RunProperties {
                    RunStyle = context.DocumentStyle.GetRunStyle(context.DocumentStyle.DefaultStyles.HyperlinkStyle)
                },
                new Text(description)));
            p.Append(h);
        }
        else
        {
            p.Append(new Run(
                new Text(description) { Space = SpaceProcessingModeValues.Preserve }));
        }

        fpart.Footnotes.Save();

        context.Properties("footnotesRef", footnotesRef);
        return footnotesRef!.Value;
    }

    /// <summary>
    /// Add a note to the Endnotes part and ensure it exists.
    /// </summary>
    /// <param name="context">The parsing context.</param>
    /// <param name="description">The description of an acronym, abbreviation, some book references, ...</param>
    /// <returns>Returns the id of the endnote reference.</returns>
    public static long AddEndnoteReference(ParsingContext context, string description)
    {
        EndnotesPart? fpart = context.MainPart.EndnotesPart ?? context.MainPart.AddNewPart<EndnotesPart>();
        var endnotesRef = context.Properties<long?>("endnotesRef");

        if (endnotesRef.HasValue)
        {
            endnotesRef++;
        }
        else if (fpart.Endnotes == null)
        {
            // Insert a new Footnotes reference
            new Endnotes(
                new Endnote(
                    new Paragraph(
                        new ParagraphProperties {
                            SpacingBetweenLines = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }
                        },
                        new Run(
                            new SeparatorMark())
                    )
                ) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = -1 },
                new Endnote(
                    new Paragraph(
                        new ParagraphProperties {
                            SpacingBetweenLines = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }
                        },
                        new Run(
                            new ContinuationSeparatorMark())
                    )
                ) { Id = 0 }).Save(fpart);
            endnotesRef = 1;
        }
        else
        {
            // The endnotesRef Id is a required field and should be unique. You can assign yourself some hard-coded
            // value but that's absolutely not safe. We will loop through the existing Footnote
            // to retrieve the highest Id.
            endnotesRef = 0;
            foreach (var p in fpart.Endnotes.Elements<Endnote>())
            {
                if (p.Id != null && p.Id > endnotesRef) endnotesRef = p.Id.Value;
            }
            endnotesRef++;
        }

        fpart.Endnotes!.Append(
            new Endnote(
                new Paragraph(
                    new ParagraphProperties {
                        ParagraphStyleId = context.DocumentStyle.GetParagraphStyle(context.DocumentStyle.DefaultStyles.EndnoteTextStyle)
                    },
                    new Run(
                        new RunProperties {
                            RunStyle = context.DocumentStyle.GetRunStyle(context.DocumentStyle.DefaultStyles.EndnoteReferenceStyle)
                        },
                        new FootnoteReferenceMark()),
                    new Run(
            // Word insert automatically a space before the definition to separate the reference number
            // with its description
                        new Text(" " + description) { Space = SpaceProcessingModeValues.Preserve })
                )
            ) { Id = endnotesRef });

        fpart.Endnotes.Save();

        context.Properties("endnotesRef", endnotesRef);
        return endnotesRef!.Value;
    }
}