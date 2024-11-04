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
using System.Runtime.CompilerServices;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Text;

namespace HtmlToOpenXml;

/// <summary>
/// Helper class that provide some extension methods to AngleSharp SDK.
/// </summary>
static class AngleSharpExtensions
{
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static HtmlAttributeCollection GetStyles(this IElement element)
    {
        return HtmlAttributeCollection.ParseStyle(element.GetAttribute("style"));
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static DirectionMode? GetTextDirection(this IHtmlElement element)
    {
        if ("rtl".Equals(element.Direction, StringComparison.OrdinalIgnoreCase))
            return DirectionMode.Rtl;
        if ("ltr".Equals(element.Direction, StringComparison.OrdinalIgnoreCase))
            return DirectionMode.Ltr;
        return null;
    }

    /// <summary>
    /// Gets whether the anchor is redirect to the `top` of the document.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool IsTopAnchor(this IHtmlAnchorElement element)
    {
        if (element.Hash.Length <= 1) return false; 
        return "#top".Equals(element.Hash, StringComparison.OrdinalIgnoreCase)
            || "#_top".Equals(element.Hash, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Gets whether the given child is preceded by any list element (<c>ol</c> or <c>ul</c>).
    /// </summary>
    public static bool IsPrecededByListElement(this INode child, out IElement? precedingElement)
    {
        precedingElement = null;

        if (child.Parent == null)
            return false;

        foreach (INode childNode in child.Parent!.ChildNodes)
        {
            if (childNode == child)
            {
                break;
            }

            if (childNode.NodeType == NodeType.Element && (
                ((IElement) childNode).LocalName == TagNames.Ol || ((IElement) childNode).LocalName == TagNames.Ul))
            {
                precedingElement = (IElement) childNode;
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Aggresively try to parse an URL.
    /// </summary>
    /// <remarks>Inline data in <see cref="IO.DataUri"/> would returns <see langword="false"/>.</remarks>
    public static bool TryParseUrl(string? uriString, UriKind uriKind,
#if NET5_0_OR_GREATER
    [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] 
#endif
    out Uri? result)
    {
        if (string.IsNullOrEmpty(uriString) || IO.DataUri.IsWellFormed(uriString!))
        {
            result = null;
            return false;
        }

        // handle link where the http:// is missing and that starts directly with www
        if(uriString!.StartsWith("www.", StringComparison.OrdinalIgnoreCase))
            uriString = "http://" + uriString;
        // or starts without the protocol
        if (uriString.StartsWith("://"))
            uriString = "http" + uriString;

        return Uri.TryCreate(uriString, uriKind, out result) 
            && (!result.IsAbsoluteUri || result.Scheme != "javascript");
    }

    /// <summary>
    /// Enumerates all the table sections (<c>tbody</c>, <c>thead</c> and <c>tfoot</c>).
    /// </summary>
    public static IEnumerable<IHtmlTableSectionElement> AsTablePartEnumerable(this IHtmlTableElement table)
    {
        if (table.Head != null) yield return table.Head;

        // AngleSharp gracefully remap the <tr> that does not expliclty sit below tbody
        foreach (var body in table.Bodies)
            yield return body;

        if (table.Foot != null) yield return table.Foot;
    }

    /// <summary>
    /// Collapse all line breaks from the given string.
    /// </summary>
    /// <param name="str">The string to examine.</param>
    /// <returns>A new string, which excludes the line breaks and
    /// ensure that two lines are merged with a space between them.</returns>
    public static string CollapseLineBreaks(this string str)
    {
        char[] chars = str.ToCharArray();
        int shift = 0, length = chars.Length, c = 0;
        while (c < length)
        {
            chars[c] = chars[c + shift];

            if (!chars[c].IsLineBreak())
            {
                c++;
                continue;
            }

            if (c > 1 && !chars[c - 1].IsWhiteSpaceCharacter() && c < length)
            {
                chars[c] = ' ';
                c++;
            }
            else
            {
                shift++;
                length--;
            }
        }

        return new string(chars, 0, length);
    }

    /// <summary>
    /// Determines whether the layout mode is inline vs block or flex.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool IsInlineLayout(string? displayMode)
    {
        return displayMode?.StartsWith("inline", StringComparison.OrdinalIgnoreCase) == true;
    }
}