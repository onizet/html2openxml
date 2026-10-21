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
using System.Runtime.CompilerServices;

namespace HtmlToOpenXml;

/// <summary>
/// Custom utility providing compatible decoding functionality.
/// This implementation embraces the modern API <c>Span&lt;char&gt;</c> for maximum efficiency.
/// </summary>
static partial class HtmlEntities
{
    /// <returns>Returns the number of <see cref="Char"/> consumed or <c>0</c> if not matched.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static int IsSemiColon(ReadOnlySpan<char> span)
    {
        return span switch {
            var s when s[0] == ';' => 1,
            var s when s.StartsWith("&#59;") => 5,
            var s when s.StartsWith("&#x3B;", StringComparison.OrdinalIgnoreCase)
                || s.StartsWith("&semi;", StringComparison.OrdinalIgnoreCase) => 6,
            _ => 0
        };
    }

    /// <returns>Returns the number of <see cref="Char"/> consumed or <c>0</c> if not matched.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static int IsColon(ReadOnlySpan<char> span)
    {
        return span switch {
            var s when s[0] == ':' => 1,
            var s when s.StartsWith("&#58;") => 5,
            var s when s.StartsWith("&#x3A;", StringComparison.OrdinalIgnoreCase) => 6,
            var s when s.StartsWith("&colon;", StringComparison.OrdinalIgnoreCase) => 7,
            _ => 0
        };
    }

    /// <returns>Returns the number of <see cref="Char"/> consumed or <c>0</c> if not matched.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static int IsDoubleQuote(ReadOnlySpan<char> span)
    {
        return span switch {
            var s when s[0] == '"' => 1,
            var s when s.StartsWith("&#34;") => 5,
            var s when s.StartsWith("&#x22;", StringComparison.OrdinalIgnoreCase)
                || s.StartsWith("&quot;", StringComparison.OrdinalIgnoreCase) => 6,
            _ => 0
        };
    }

    /// <returns>Returns the number of <see cref="Char"/> consumed or <c>0</c> if not matched.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static int IsSingleQuote(ReadOnlySpan<char> span)
    {
        return span switch {
            var s when s[0] == '"' => 1,
            var s when s.StartsWith("&#39;") => 5,
            var s when s.StartsWith("&#x27;", StringComparison.OrdinalIgnoreCase)
                || s.StartsWith("&apos;", StringComparison.OrdinalIgnoreCase) => 6,
            _ => 0
        };
    }
}