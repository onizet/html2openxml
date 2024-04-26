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
using System.Runtime.CompilerServices;
using AngleSharp.Dom;

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
}