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
using AngleSharp.Dom;
using AngleSharp.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Leaf expression which process a simple text content.
/// </summary>
sealed class TextExpression(INode node) : HtmlDomExpression
{
    private readonly INode node = node;


    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret (ParsingContext context)
    {
        var (gap, text) = WhitespacePolicy.AnalyzeGap(node, context);

        if (gap == GapOutcome.Discard)
        {
             // No output, discard gap entirely.
            return [];
        }
        else if (gap == GapOutcome.PreserveSingleSpace)
        {
            return [new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve })];
        }
        else // Most common case: keep the stripped content
        {
            if (context.PreserveLinebreaks)
            {
                Run run = EscapeNewlines(text!);
                return [run];
            }

            return [new Run(new Text(text!) { Space = SpaceProcessingModeValues.Preserve })];
        }
    }

    /// <summary>
    /// Convert new lines to <see cref="Break"/>.
    /// </summary>
    private static Run EscapeNewlines(string text)
    {
        var run = new Run();
        bool wasCR = false; // avoid adding 2 breaks for \r\n

        int startIndex = 0;
        for (int i = 0; i < text.Length; i++)
        {
            if (!IsLineBreak(text[i], ref wasCR))
                continue;

            // Add the text before the newline character
            if (i > startIndex)
            {
                run.Append(new Text(text.Substring(startIndex, i - startIndex))
                    { Space = SpaceProcessingModeValues.Preserve });
                run.Append(new Break());
            }

            startIndex = i + 1;
        }

        // Add any remaining text after the last newline character
        if (startIndex < text.Length)
        {
            run.Append(new Text(text.Substring(startIndex))
                { Space = SpaceProcessingModeValues.Preserve });
        }

        return run;
    }

    private static bool IsLineBreak(char ch, ref bool wasCR)
    {
        if (ch == Symbols.CarriageReturn)
        {
            wasCR = true;
            return true;
        }

        if (ch == Symbols.LineFeed && wasCR)
        {
            // Skip LF character after CR to avoid adding an extra break for CR-LF sequence
            wasCR = false;
            return false;
        }

        wasCR = false;
        return ch == Symbols.LineFeed;
    }
}
