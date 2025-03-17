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
#if NET5_0_OR_GREATER
using System.Collections.Frozen;
#endif
using System.Collections.Generic;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Leaf expression which process a simple text content.
/// </summary>
sealed class TextExpression(INode node) : HtmlDomExpression
{
    static readonly ISet<string> AllPhrasings = InitPhrasingSets();
    private readonly INode node = node;

    private static ISet<string> InitPhrasingSets()
    {
        var sets = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase) {
            TagNames.A, TagNames.Abbr, TagNames.B, TagNames.Big, TagNames.Cite, TagNames.Code,
            TagNames.Del, TagNames.Dfn, TagNames.Em, TagNames.Font, TagNames.Hr, TagNames.I,
            TagNames.Img, TagNames.Ins, TagNames.Kbd, TagNames.Mark, TagNames.NoBr, TagNames.Q,
            TagNames.Rp, TagNames.Rt, TagNames.S, TagNames.Samp, TagNames.Small, TagNames.Span,
            TagNames.Strike, TagNames.Strong, TagNames.Sub, TagNames.Sup, TagNames.Time,
            TagNames.Tt, TagNames.U, TagNames.Var
        };

#if NET5_0_OR_GREATER
        return sets.ToFrozenSet(StringComparer.InvariantCultureIgnoreCase);
#else
        return sets;
#endif
    }

    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret (ParsingContext context)
    {
        string text = node.TextContent.Normalize();

        if (text.Length == 0)
            return [];

        if (!context.PreserveLinebreaks)
        {
            text = text.CollapseLineBreaks();
            if (text.Length == 0)
                return [];
        }

        // https://developer.mozilla.org/en-US/docs/Web/API/Document_Object_Model/Whitespace
        // If there is a space between two phrasing elements, the user agent should collapse it to a single space character.
        if (context.CollapseWhitespaces)
        {
            bool startsWithSpace = text[0].IsWhiteSpaceCharacter(),
                endsWithSpace = text[text.Length - 1].IsWhiteSpaceCharacter(),
                preserveBorderSpaces = AllPhrasings.Contains(node.Parent!.NodeName),
                prevIsPhrasing = node.PreviousSibling is not null &&
                    (AllPhrasings.Contains(node.PreviousSibling.NodeName) || node.PreviousSibling!.NodeType == NodeType.Text),
                nextIsPhrasing = node.NextSibling is not null &&
                    (AllPhrasings.Contains(node.NextSibling.NodeName) || node.NextSibling!.NodeType == NodeType.Text);

            text = text.CollapseAndStrip();

            // keep a collapsed single space if it stands between 2 phrasings that respect.
            // doesn't ends/starts with a whitespace
            if (text.Length == 0 && prevIsPhrasing && nextIsPhrasing
                && (endsWithSpace || startsWithSpace)
                && !(node.PreviousSibling!.TextContent[node.PreviousSibling!.TextContent.Length - 1].IsWhiteSpaceCharacter()
                    || node.NextSibling!.TextContent[0].IsWhiteSpaceCharacter()
                ))
            {
                return [new Run(new Text(" "))];
            }
            // we strip out all whitespaces and we stand inside a div. Just skip this text content
            if (text.Length == 0 && !preserveBorderSpaces)
            {
                return [];
            }

            // if previous element is an image, append a space separator
            // if this is a non-empty phrasing element, append a space separator
            if (startsWithSpace && node.PreviousSibling is IHtmlImageElement)
            {
                text = " " + text;
            }
            else if (startsWithSpace && prevIsPhrasing
                && !node.PreviousSibling!.TextContent[node.PreviousSibling.TextContent.Length - 1].IsWhiteSpaceCharacter())
            {
                text = " " + text;
            }

            if (endsWithSpace && (
                // next run is not starting with a linebreak
                (nextIsPhrasing && node.NextSibling!.TextContent.Length > 0 &&
                    !node.NextSibling!.TextContent[0].IsLineBreak()) ||
                // if there is no more text element or is empty, eat the trailing space
                (preserveBorderSpaces && (node.NextSibling is not null
                    || node.Parent.NextSibling is not null))))
            {
                text += " ";
            }
        }


        if (text.Length == 0)
            return [];

        if (!context.PreserveLinebreaks)
            return [new Run(new Text(text))];

        Run run = EscapeNewlines(text);
        return [run];
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
