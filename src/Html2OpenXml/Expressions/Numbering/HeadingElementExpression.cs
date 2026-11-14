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
    private static readonly Regex numberingRegex = new(@"^\s*(?<number>\d+(?:\.\d+)*\.?\s*)",
        RegexOptions.Compiled, TimeSpan.FromMilliseconds(100));

    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret (ParsingContext context)
    {
        char level = node.NodeName[1];

        var childElements = base.Interpret(context).ToList();
        if (childElements.Count == 0) // no text = skip this heading
            return childElements;

        var paragraph = childElements[0] as Paragraph;
        string baseStyle;

        paragraph ??= new(childElements);
        paragraph.ParagraphProperties ??= new();

        var runElement = childElements.FirstOrDefault();
        if (runElement != null && context.Converter.SupportsHeadingNumbering &&
            TryStripHeadingNumbering(runElement))
        {
            if (string.Equals(
                context.DocumentStyle.DefaultStyles.HeadingStyle,
                context.DocumentStyle.DefaultStyles.NumberedHeadingStyle))
            {
                // Only apply the numbering if a custom numbered heading style has not been defined.
                // If the user defined a custom numbered heading style (with numbering), Word has
                // the numbering automatically done.
                // Defining a numbering here messes that up, so we only add the numbering if
                // a specific numbering heading style has not been provided
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

            // Apply numbered heading style
            baseStyle = context.DocumentStyle.DefaultStyles.NumberedHeadingStyle;
        }
        else
        {
            // Apply normal heading style
            baseStyle = context.DocumentStyle.DefaultStyles.HeadingStyle;
        }

        // Apply normal heading style
        paragraph.ParagraphProperties.ParagraphStyleId =
            context.DocumentStyle.GetParagraphStyle(baseStyle + level);

        return [paragraph];
    }

    internal static bool TryStripHeadingNumbering(OpenXmlElement runElement)
    {
        var headingText = runElement.InnerText;
        if (string.IsNullOrEmpty(headingText))
            return false;

        // Check if the line starts with a number format (1., 1.1., 1.1.1.)
        // If it does, make sure we make the heading a numbered item
        Match regexMatch;
        try
        {
            regexMatch = numberingRegex.Match(headingText);
        }
        catch (RegexMatchTimeoutException)
        {
            return false;
        }

        // Strip the numbering only when it is followed by actual heading text.
        // Values such as: "1", "1.1" are not considered numbered headings (issue #189)
        if (regexMatch.Success &&
            headingText.Length > regexMatch.Groups["number"].Length)
        {
            RemoveLeadingCharacters(runElement, regexMatch.Groups["number"].Length);

            return true;
        }
        return false;
    }

    /// <summary>
    /// Remove characters from the start of a Run. A Run can contain multiple Text nodes,
    /// and Run.InnerText is the concatenation of all of them.
    /// The numbering prefix may therefore span several Text nodes.
    /// </summary>
    /// <param name="runElement">The parent Run.</param>
    /// <param name="count">The number of characters to remove.</param>
    private static void RemoveLeadingCharacters(OpenXmlElement runElement, int count)
    {
        foreach (var text in runElement.Descendants<Text>())
        {
            if (count <= 0)
                break;
            
            if (text.Text.Length <= count)
            {
                count -= text.Text.Length;
                text.Text = string.Empty;
            }
            else
            {
                text.Text = text.Text.Substring(count);
                break;
            }
        }
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