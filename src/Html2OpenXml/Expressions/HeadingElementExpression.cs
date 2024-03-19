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
sealed class HeadingElementExpression(IHtmlElement node) : FlowElementExpression(node)
{
    /// <inheritdoc/>
    public override IEnumerable<OpenXmlCompositeElement> Interpret (ParsingContext context)
    {
        char level = node.NodeName[1];

        var childElements = base.Interpret(context);
        if (!childElements.Any()) // no text = skip this heading
            return childElements;

        var paragraph = childElements.FirstOrDefault() as Paragraph;

        paragraph ??= new Paragraph(childElements);

        paragraph.InsertInProperties(prop => 
            prop.ParagraphStyleId = 
                context.DocumentStyle.GetParagraphStyle(context.DocumentStyle.DefaultStyles.HeadingStyle + level)
        );

        // Check if the line starts with a number format (1., 1.1., 1.1.1.)
        // If it does, make sure we make the heading a numbered item
        OpenXmlElement? firstElement = childElements.FirstOrDefault();
        Match regexMatch = Regex.Match(firstElement?.InnerText ?? string.Empty, @"(?m)^(\d+\.)*\s");

        // Make sure we only grab the heading if it starts with a number
        if (regexMatch.Groups.Count > 1 && regexMatch.Groups[1].Captures.Count > 0)
        {
            int indentLevel = regexMatch.Groups[1].Captures.Count;

            // Strip numbers from text
            if (firstElement != null)
                firstElement.InnerXml = firstElement.InnerXml
                    .Replace(firstElement.InnerText, firstElement.InnerText.Substring(indentLevel * 2 + 1)); // number, dot and whitespace

            //TODO: ici faut refaire
            context.DocumentStyle.NumberingList.ApplyNumberingToHeadingParagraph(paragraph, indentLevel);
        }

        return [paragraph];
    }
}