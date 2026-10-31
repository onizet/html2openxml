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
/// Polyfill helper class to provide extension methods for <see cref="ReadOnlySpan{T}"/>.
/// </summary>
static partial class SpanExtensions
{
    /// <summary>
    /// Shim method to convert <see cref="string"/> to <see cref="byte"/>.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    [System.Diagnostics.DebuggerHidden]
    public static byte AsByte(this ReadOnlySpan<char> span, NumberStyles style)
    {
#if NET5_0_OR_GREATER
        return byte.Parse(span, style);
#else
        return byte.Parse(span.ToString(), style);
#endif
    }

    /// <summary>
    /// Shim method to convert <see cref="string"/> to <see cref="double"/>.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    [System.Diagnostics.DebuggerHidden]
    public static double AsDouble(this ReadOnlySpan<char> span)
    {
#if NET5_0_OR_GREATER
        return double.Parse(span, CultureInfo.InvariantCulture);
#else
        return double.Parse(span.ToString(), CultureInfo.InvariantCulture);
#endif
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool Is(this ReadOnlySpan<char> span, string value)
        => span.Equals(value, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Convert a potential percentage value to its numeric representation.
    /// Saturation and Lightness can contains both a percentage value or a value comprised between 0.0 and 1.0. 
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    [System.Diagnostics.DebuggerHidden]
    public static double AsPercent (this ReadOnlySpan<char> span)
    {
        int index = span.IndexOf('%');
        if (index > -1)
        {
            double parsedValue = span.Slice(0, index).AsDouble() / 100d;
            return Math.Min(1, Math.Max(0, parsedValue));
        }

        return span.AsDouble();
    }

    /// <summary>
    /// Parses the source <see cref="ReadOnlySpan{T}"/> for the specified style attribute separators, 
    /// populating the <paramref name="destination"/> span with <see cref="Range"/> instances
    /// representing the regions between the separators.
    /// </summary>
    /// <param name="span">The source span to parse.</param>
    /// <param name="destination">The destination span into which the resulting ranges are written.</param>
    /// <param name="separator">A character that delimits the regions in this instance.</param>
    /// <param name="skipSeparatorIfPrecededBy">If <paramref name="separator"/> is preceded by this character, the separator will be treated as a normal character.</param>
    /// <returns>The number of ranges written into <paramref name="destination"/>.</returns>
    public static int SplitCompositeAttribute(this ReadOnlySpan<char> span, Span<Range> destination,
        char separator, char? skipSeparatorIfPrecededBy = null)
    {
        // If the destination is empty, there's nothing to do.
        if (span.IsEmpty || destination.IsEmpty)
            return 0;

        int matchCount = 0, tokenStart = 0;
        bool isEscaping = false;
        char endEscapingChar = '\0';

        for (int i = 0; i < span.Length; )
        {
            char ch = span[i];
            int consumed = 1;

            // is this an encoded html entity?
            if (ch == '&' && HtmlEntities.IsRecognizedEntity(span.Slice(i), out var entity))
            {
                ch = entity.DecodedChar;
                consumed = entity.EncodedLength;
            }

            // are we going into escape mode, e.g., matches the separator but enclosed inside boundary
            // such as quote or parenthesis
            if (!isEscaping && IsEscapable(ch))
            {
                endEscapingChar = ch is '('? ')' : ch;
                isEscaping = true;
            }
            // this is a separator but maybe we will need to skip it
            // eg: "Arial, Verdana bold 1em" -> the space after the comma must be skipped
            // eg: rgb(233, 233, 233) -> rgb(233,233,233)
            else if (!isEscaping && ch == separator &&
                (!skipSeparatorIfPrecededBy.HasValue || (i > 0 && span[i - 1] != skipSeparatorIfPrecededBy.Value)))
            {
                if (i  > tokenStart)
                {
                    destination[matchCount++] = new Range(tokenStart, i);
                    if (matchCount >= destination.Length)
                        return matchCount;
                }
                tokenStart = i + consumed;
            }
            // end of escaping sequence
            else if (ch == endEscapingChar)
            {
                isEscaping = false;
            }

            i += consumed;
        }

        // process the reminder match
        if (tokenStart < span.Length && matchCount < destination.Length)
        {
            destination[matchCount++] = new Range(tokenStart, span.Length);
        }

        return matchCount;
    }

    /// <summary>
    /// Trim the (single or double) quote around text.
    /// Supports also the html-encoded quote entity.
    /// </summary>
    public static ReadOnlySpan<char> Unquote(this ReadOnlySpan<char> span)
    {
        int start = 0;
        while (start < span.Length)
        {
            var ch = span[start];
            int consumed = 1;
            if (ch is '&' && HtmlEntities.IsRecognizedEntity(span.Slice(start), out var entity))
            {
                ch = entity.DecodedChar;
                consumed = entity.EncodedLength;
            }

            if (!IsQuote(ch))
                break;

            start += consumed;
        }

        int end = span.Length - 1;
        while (end >= start)
        {
            var ch = span[end];
            int consumed = 1, ampIndex;
            if (ch is ';' &&
                (ampIndex = span.Slice(0, end + 1).LastIndexOf('&')) > -1 &&
                HtmlEntities.IsRecognizedEntity(span.Slice(ampIndex), out var entity))
            {
                ch = entity.DecodedChar;
                consumed = entity.EncodedLength;
            }

            if (!IsQuote(ch))
                break;

            end -= consumed;
        }

        return span.Slice(start, end - start + 1);
    }

    private static bool IsQuote(char c)
        => c is '"' or '\'';

    private static bool IsEscapable(char c)
        => c is '(' or '"' or '\'';
}
