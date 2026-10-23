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
using System.Globalization;
using System.Runtime.CompilerServices;

namespace HtmlToOpenXml;

/// <summary>
/// Custom utility providing compatible decoding functionality.
/// This implementation embraces the modern API <c>Span&lt;char&gt;</c> for maximum efficiency.
/// </summary>
static partial class HtmlEntities
{
    /// <param name="EncodedLength">The number  of <see cref="Char"/> consumed or <c>0</c> if not matched.</param>
    /// <param name="DecodedChar">The decoded character.</param>
    public readonly record struct EntityMatchResult(int EncodedLength, char DecodedChar)
    {
        public static readonly EntityMatchResult Nil = new(0, '\0');
    }

    /// <summary>
    /// Try to detect encoded html entities such as single or double quote.
    /// </summary>
    public static bool IsRecognizedEntity(ReadOnlySpan<char> span, out EntityMatchResult result)
    {
        var semiIndex = span.IndexOf(';');
        if (semiIndex == -1)
        {
            result = EntityMatchResult.Nil;
            return false;
        }

        span = span.Slice(0, semiIndex + 1);

        // fast path for common named entities
        if (span.Equals("&quot;", StringComparison.OrdinalIgnoreCase))
        {
            result = new EntityMatchResult(6, '"');
            return true;
        }
        else if (span.Equals("&apos;", StringComparison.OrdinalIgnoreCase))
        {
            result = new EntityMatchResult(6, '\'');
            return true;
        } 
        else if (span.Equals("&colon;", StringComparison.OrdinalIgnoreCase))
        {
            result = new EntityMatchResult(7, ':');
            return true;
        }
        else if (span.Equals("&semi;", StringComparison.OrdinalIgnoreCase))
        {
            result = new EntityMatchResult(6, ';');
            return true;
        }
        else if (span.Equals("&nbsp;", StringComparison.OrdinalIgnoreCase))
        {
            result = new EntityMatchResult(6, ' ');
            return true;
        }

        if (TryDecodeNumericEntity(span, out var codePoint)
            // not using System.Text.Rune.IsValid() because this parser is only interested
            // in a minimal subset: ; : ( ) " '
            && codePoint <= char.MaxValue)
        {
            result = new EntityMatchResult(span.Length, (char) codePoint);
            return true;
        }

        result = EntityMatchResult.Nil;
        return false;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static int IsColon(ReadOnlySpan<char> span)
    {
        return span switch {
            var s when s[0] == ':' => 1,
            var s when s.StartsWith("&#58;") => 5,
            var s when s.StartsWith("&#x3A;", StringComparison.OrdinalIgnoreCase) => 6,
            var s when s.StartsWith("&colon;", StringComparison.OrdinalIgnoreCase) => 7,
            _ => 0
        };
    }

    /// <summary>
    /// Try to parse an HTML nnumeric entity and returns its Unicode code point.
    /// </summary>
    /// <param name="entity">The entity text without the surrounding <c>&amp;</c> and <c>;</c>.</param>
    /// <param name="codePoint">The Unicode point representing that entity.</param>
    private static bool TryDecodeNumericEntity(ReadOnlySpan<char> entity, out uint codePoint)
    {
        if (entity.Length < 4 || entity[1] != '#')
        {
            codePoint = 0;
            return false;
        }

        NumberStyles numberStyles = NumberStyles.Integer;
        int start = 2;
        if (entity[2] is 'x' or 'X')
        {
            numberStyles = NumberStyles.AllowHexSpecifier;
            start = 3;
        }

        entity = entity.Slice(start, entity.Length - 1 - start);

        return uint.TryParse(
#if NET5_0_OR_GREATER
            entity,
#else
            entity.ToString(),
#endif
            numberStyles, CultureInfo.InvariantCulture, out codePoint);
    }
}