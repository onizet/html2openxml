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
#if NET5_0_OR_GREATER
using System.Collections.Frozen;
#endif
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
            TagNames.Ins, TagNames.Kbd, TagNames.Mark, TagNames.NoBr, TagNames.Q,
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
            var previousSibling = node.PreviousSibling;
            var nextSibling = node.NextSibling;
            bool startsWithSpace = text[0].IsWhiteSpaceCharacter(),
                endsWithSpace = text[text.Length - 1].IsWhiteSpaceCharacter(),
                preserveBorderSpaces = AllPhrasings.Contains(node.Parent!.NodeName),
                prevIsPhrasing = previousSibling is not null && 
                    (previousSibling.NodeType == NodeType.Text || AllPhrasings.Contains(previousSibling.NodeName)),
                nextIsPhrasing = nextSibling is not null && 
                    (nextSibling.NodeType == NodeType.Text || AllPhrasings.Contains(nextSibling.NodeName));

            text = text.CollapseAndStrip();

            // keep a collapsed single space if it stands between 2 phrasings that respect.
            // doesn't ends/starts with a whitespace
            if (text.Length == 0 && prevIsPhrasing && nextIsPhrasing
                && (endsWithSpace || startsWithSpace)
                && !(previousSibling!.TextContent.Length == 0
                    || nextSibling!.TextContent.Length == 0
                    || previousSibling!.TextContent[previousSibling!.TextContent.Length - 1].IsWhiteSpaceCharacter()
                    || nextSibling!.TextContent[0].IsWhiteSpaceCharacter()
                ))
            {
                return [new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve })];
            }

            // is this an inter-element whitespace btw 2 phrasings?
            var isWhitespace = text.Length == 0;

            // we strip out all whitespaces and we stand inside a div. Just skip this text content
            if (isWhitespace && !(prevIsPhrasing && nextIsPhrasing) && !preserveBorderSpaces)
            {
                return [];
            }

            // if previous element is an image, append a space separator
            if ((startsWithSpace && previousSibling is IHtmlImageElement)
                // if this is a non-empty phrasing element, append a space separator
                || (!isWhitespace && startsWithSpace && prevIsPhrasing
                && previousSibling!.TextContent.Length > 0
                && !previousSibling!.TextContent[previousSibling.TextContent.Length - 1].IsWhiteSpaceCharacter()))
            {
                text = " " + text;
            }
            // if no immediate previous sibling traverse nested phrasing content for the previous meaningful content
            // if previous meaningful content has trailing space, in that case skip it.
            else if (startsWithSpace && !isWhitespace && previousSibling is null && PreviousPhrasingContentNeedsLeadingSpace(node))
            {
                text = " " + text;
            }

            if (endsWithSpace && !isWhitespace && (
                // next run is not starting with a linebreak
                (nextIsPhrasing && nextSibling!.TextContent.Length > 0 &&
                    !nextSibling!.TextContent[0].IsLineBreak()) ||
                // if there is no more text element or is empty, eat the trailing space
                (preserveBorderSpaces && (nextSibling is not null
                    || HasFollowingPhrasingContent(node)))))
            {
                text += " ";
            }
        }


        if (text.Length == 0)
            return [];

        if (!context.PreserveLinebreaks)
            return [new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })];

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

    private static bool PreviousPhrasingContentNeedsLeadingSpace(INode node)
    {
        INode? current = node;
        while (current?.Parent is IHtmlElement parent)
        {
            var sibling = current.PreviousSibling;
            while (sibling is not null)
            {
                // meaningful only if it contains non-whitespace content
                // or content without trailing whitespacce.
                if (sibling.NodeType == NodeType.Text)
                {
                    var text = sibling.TextContent;
                    if (!string.IsNullOrWhiteSpace(sibling.TextContent))
                        return false;

                    return !text[text.Length - 1].IsWhiteSpaceCharacter();
                }
                // phrasing element may contain the actual preceding text.
                if (sibling is IHtmlElement siblingElement && AllPhrasings.Contains(siblingElement.NodeName))
                {
                    if (ContainsMeaningfulText(siblingElement))
                        return !EndsWithWhitespace(siblingElement);

                    sibling = sibling.PreviousSibling;
                    continue;
                }

                return false; // traversal must be containsed within block level element
            }

            if (!AllPhrasings.Contains(parent.NodeName))
                return false;

            current = parent;
        }
        return false;
    }

    private static bool HasFollowingPhrasingContent(INode node)
    {
        INode? current = node;
        while (current?.Parent is IHtmlElement parent)
        {
            var sibling = current.NextSibling;
            while (sibling is not null)
            {
                // meaningful only if it contains non-whitespace content.
                if (sibling.NodeType == NodeType.Text)
                {
                    if (!string.IsNullOrWhiteSpace(sibling.TextContent))
                        return true;

                    sibling = sibling.NextSibling;
                    continue;
                }

                if (AllPhrasings.Contains(sibling.NodeName))
                {
                    if (ContainsMeaningfulText(sibling))
                        return true;

                    sibling = sibling.NextSibling;
                    continue;
                }

                return false; // traversal must be containsed within block level element
            }

            if (!AllPhrasings.Contains(parent.NodeName))
                return false;

            current = parent;
        }
        return false;
    }

    private static bool ContainsMeaningfulText(INode element)
    {
        foreach (var child in element.ChildNodes)
        {
            if (child.NodeType == NodeType.Text)
            {
                if (!string.IsNullOrWhiteSpace(child.TextContent))
                    return true;
                continue;
            }

            if (child is IHtmlElement childElement && AllPhrasings.Contains(childElement.NodeName) && ContainsMeaningfulText(childElement))
                return true;
        }
        return false;
    }

    private static bool EndsWithWhitespace(IHtmlElement element)
    {
        var text = element.TextContent;
        return text.Length > 0 && text[text.Length - 1].IsWhiteSpaceCharacter();
    }
}
