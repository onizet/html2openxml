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
using System.Text.RegularExpressions;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of a heading element.
/// </summary>
sealed class HeadingElementExpression(IHtmlElement node) : NumberingExpressionBase(node)
{
    private static readonly Regex numberingRegex = new(@"^\s*(\d+\.?)*\s*");

    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret (ParsingContext context)
    {
        char level = node.NodeName[1];

        var childElements = base.Interpret(context);
        if (!childElements.Any()) // no text = skip this heading
            return childElements;

        var paragraph = childElements.FirstOrDefault() as Paragraph;

        paragraph ??= new Paragraph(childElements);
        paragraph.ParagraphProperties ??= new();
        paragraph.ParagraphProperties.ParagraphStyleId = 
            context.DocumentStyle.GetParagraphStyle(context.DocumentStyle.DefaultStyles.HeadingStyle + level);
 
        var runElement = childElements.FirstOrDefault();
        if (runElement != null && context.Converter.SupportsHeadingNumbering && IsNumbering(runElement))
        {
            var abstractNumId = GetOrCreateListTemplate(context, HeadingNumberingName);
            var instanceId = GetListInstance(abstractNumId);
            if (!instanceId.HasValue)
            {
                instanceId = IncrementInstanceId(context, abstractNumId);
            }

            var numbering = context.MainPart.NumberingDefinitionsPart!.Numbering!;
            numbering.Append(
                new NumberingInstance(
                    new AbstractNumId() { Val = abstractNumId }
                )
                { NumberID = instanceId });
            SetNumbering(paragraph, level - '0', instanceId.Value);
        }

        return [paragraph];
    }

    private static bool IsNumbering(OpenXmlElement runElement)
    {
        // Check if the line starts with a number format (1., 1.1., 1.1.1.)
        // If it does, make sure we make the heading a numbered item
        Match regexMatch = numberingRegex.Match(runElement.InnerText ?? string.Empty);

        // Make sure we only grab the heading if it starts with a number
        if (regexMatch.Groups.Count > 1 && regexMatch.Groups[1].Captures.Count > 0)
        {
             // Strip numbers from text
            runElement.InnerXml = runElement.InnerXml
                .Replace(runElement.InnerText!, runElement.InnerText!.Substring(regexMatch.Length));

            return true;
        }
        return false;
    }

    /// <summary>
    /// Apply numbering to the heading paragraph.
    /// </summary>
    private static void SetNumbering(Paragraph paragraph, int level, int instanceId)
    {
        // Apply numbering to paragraph
        paragraph.ParagraphProperties ??= new();
        paragraph.ParagraphProperties.NumberingProperties = new NumberingProperties {
            NumberingLevelReference = new() { Val = level - 1 },
            NumberingId = new() { Val = instanceId }
        };
    }
}