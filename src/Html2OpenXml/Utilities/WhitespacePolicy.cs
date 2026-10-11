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
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Text;

namespace HtmlToOpenXml;

readonly record struct GapAnalysisResult(GapOutcome Outcome, string? ProcessedText = null);

enum GapOutcome
{
    /// <summary>
    /// The gap represents content that must be preserved as a single space.
    /// </summary>
    PreserveSingleSpace, 

    /// <summary>
    /// The gap is content that needs to be passed through with its own boundaries.
    /// </summary>
    ContentPassedThrough, 

    /// <summary>
    /// The gap should be completely discarded.
    /// </summary>
    Discard
}

static class WhitespacePolicy
{
    static readonly ISet<string> AllPhrasings = InitPhrasingSets();

    private static ISet<string> InitPhrasingSets()
    {
        var sets = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase) {
            TagNames.A, TagNames.Abbr, TagNames.B, TagNames.Big, TagNames.Cite, TagNames.Code,
            TagNames.Del, TagNames.Dfn, TagNames.Em, TagNames.Font, TagNames.Hr, TagNames.I,
            TagNames.Ins, TagNames.Kbd, TagNames.Mark, TagNames.NoBr, TagNames.Q,
            TagNames.Rp, TagNames.Rt, TagNames.S, TagNames.Samp, TagNames.Small, TagNames.Span,
            TagNames.Strike, TagNames.Strong, TagNames.Sub, TagNames.Sup, TagNames.Time,
            TagNames.Tt, TagNames.U, TagNames.Var,
            TagNames.Body // not a real phrasing but standalone text are parent to it
        };

#if NET5_0_OR_GREATER
        return sets.ToFrozenSet(StringComparer.InvariantCultureIgnoreCase);
#else
        return sets;
#endif
    }

    /// <summary>
    /// Analyzes the current node's text content and its neighbors to determine 
    /// how whitespace should be treated under CollapseWhitespaces mode.
    /// </summary>
    /// <seealso href="https://developer.mozilla.org/en-US/docs/Web/API/Document_Object_Model/Whitespace"/>
    public static GapAnalysisResult AnalyzeGap(INode node, ParsingContext context)
    {
        string text = node.TextContent.Normalize();

        if (text.Length == 0)
            return new(GapOutcome.Discard);

        if (!context.PreserveLinebreaks)
        {
            text = text.CollapseLineBreaks();
            if (text.Length == 0)
                return new(GapOutcome.Discard);
        }

        if (!context.CollapseWhitespaces)
            return new(GapOutcome.ContentPassedThrough, text);

        var previousNode = FindPreviousMeaningfulNode(node);
        var nextNode = FindNextMeaningfulNode(node);

        bool prevIsPhrasing = IsPhrasing(previousNode, allowsTextNode: true) && HasMeaningfulContent(previousNode);
        bool nextIsPhrasing = IsPhrasing(nextNode, allowsTextNode: true) && HasMeaningfulContent(nextNode);

        bool hasPrefixWhitespace = text.StartsWithWhitespace(),
             hasTrailingWhitespace = text.EndsWithWhitespace();
        bool hasBoundaryWhitespace = hasPrefixWhitespace || hasTrailingWhitespace;

        text = text.CollapseAndStrip();

        // Are we sandwiched between two run elements?
        if (prevIsPhrasing && nextIsPhrasing)
        {
            // Standalone whitespace in a tight inline context (e.g., between </b> and <b>).
            // This gap is semantically important and must be maintained. We reduce it to a single preserved space.
            if (text.Length == 0 && hasBoundaryWhitespace
                && !previousNode!.TextContent.EndsWithWhitespace())
            {
                return new(GapOutcome.PreserveSingleSpace);
            }
        }

        // Whitespace in a looser block context (e.g., inside a standard div)
        // and not serving as vital inter-element spacing. We collapse/discard it entirely.
        if (text.Length == 0)
        {
            return new(GapOutcome.Discard);
        }


        // Apply leading/trailing spacing adjustments before returning.


        // if previous element is an image, append a space separator
        if (previousNode is IHtmlImageElement)
        {
            text = Symbols.Space + text;
        }
        else if (hasPrefixWhitespace)
        {
            // if this is a non-empty phrasing element, append a space separator
            if (prevIsPhrasing && !previousNode!.TextContent.EndsWithWhitespace())
            {
                text = Symbols.Space + text;
            }
        }

        if (hasTrailingWhitespace)
        {
            // next run is not starting with a break line nor starting with a whitespace
            if (nextIsPhrasing && 
                !(nextNode!.TextContent[0].IsLineBreak() || nextNode.TextContent.StartsWithWhitespace()))
            {
                text += Symbols.Space;
            }
            // if there is no more text element or is empty, eat the trailing space
            else if (IsPhrasing(node.Parent) && nextNode is null)
            {
                // eat trailing space
            }
            // display exception: we enclose image with spaces
            else if (IsPhrasing(nextNode, allowsTextNode: true) || nextNode is IHtmlImageElement)
            {
                text += Symbols.Space;
            }
        }

        // If the context is loose (e.g., inside a div), we default to discard if it's just space.
        if (text.Length == 0)
        {
            return new(GapOutcome.Discard);
        }

        return new(GapOutcome.ContentPassedThrough, text);
    }

    private static INode? FindPreviousMeaningfulNode(INode node)
    {
        INode? prev = node.PreviousSibling;
        if (prev is null && IsPhrasing(node.Parent))
        {
            var parent = FindStructuralPhrasingAncestor(node);
            if (parent?.PreviousSibling != null)
            {
                prev = parent.PreviousSibling.NodeType == NodeType.Text? 
                    parent.PreviousSibling : parent.PreviousSibling.LastChild;
            }
        }

        while (prev != null)
        {
            // If we hit a dead end (non-phrasing/non-content), stop here.
            if (!IsPhrasing(prev) || HasMeaningfulContent(prev))
                return prev;

            prev = prev.PreviousSibling;
        }

        return null;
    }

    private static INode? FindNextMeaningfulNode(INode node)
    {
        INode? next = node.NextSibling;
        if (next is null && IsPhrasing(node.Parent))
        {
            var parent = FindStructuralPhrasingAncestor(node);
            if (parent?.NextSibling != null)
            {
                next = parent.NextSibling.NodeType == NodeType.Text? 
                    parent.NextSibling : parent.NextSibling.FirstChild;
            }
        }

        while (next != null)
        {
            if (!IsPhrasing(next) || HasMeaningfulContent(next))
                return next;

            next = next.NextSibling;
        }

        return null;
    }

    /// <summary>
    /// Climbs up the DOM hierarchy through transient containers until it reaches a node that defines 
    /// the beginning of a logical, structural content block.
    /// </summary>
    private static INode? FindStructuralPhrasingAncestor(INode node)
    {
        var parent = node.Parent!;
        while (parent != null)
        {
            // Stop when:
            // 1. The grandparent is NOT a phrasing element (i.e., we hit a block-level container like <div>, <p>, etc.).
            // OR 2. The grandparent HAS multiple children (i.e., it's a root container holding different structural components).
            if (IsPhrasing(parent.Parent) == false || parent.Parent.ChildNodes.Length > 1)
                break;
            parent = parent.Parent;
        }
        return parent;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool IsPhrasing([NotNullWhen(true)] INode? node, bool allowsTextNode = false)
        => node != null && (
            AllPhrasings.Contains(node.NodeName)
            || (allowsTextNode && node.NodeType == NodeType.Text ));

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool HasMeaningfulContent(INode node)
        => node.NodeType == NodeType.Text || node.TextContent.Length > 0;

    private static bool StartsWithWhitespace(this string text)
        => text[0].IsWhiteSpaceCharacter();

    private static bool EndsWithWhitespace(this string text)
        => text[^1].IsWhiteSpaceCharacter();
}
