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
static class SpanExtensions
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
    /// Shim method to remain compliant with pre-NET 8 framework.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    [System.Diagnostics.DebuggerHidden]
    public static ReadOnlySpan<T> Slice<T>(this ReadOnlySpan<T> span, Range range)
    {
#if NET5_0_OR_GREATER
        return span[range];
#else
        var (start, length) = range.GetOffsetAndLength(span.Length);
        return span.Slice(start, length);
#endif
    }

#if !NET5_0_OR_GREATER
    /// <summary>
    /// Parses the source <see cref="ReadOnlySpan{T}"/> for the specified <paramref name="separator"/>, 
    /// populating the <paramref name="destination"/> span with <see cref="Range"/> instances
    /// representing the regions between the separators.
    /// </summary>
    /// <param name="span">The source span to parse.</param>
    /// <param name="destination">The destination span into which the resulting ranges are written.</param>
    /// <param name="separator">A character that delimits the regions in this instance.</param>
    /// <param name="options">A bitwise combination of the enumeration values that specifies whether to trim whitespace and include empty ranges.</param>
    /// <returns>The number of ranges written into <paramref name="destination"/>.</returns>
    public static int Split(this ReadOnlySpan<char> span, Span<Range> destination,
        char separator, StringSplitOptions options = StringSplitOptions.None)
    {
        // If the destination is empty, there's nothing to do.
        if (destination.IsEmpty)
            return 0;

        int matches = 0;
        int startIndex = 0;
        while (span.Length > 0)
        {
            int index = span.IndexOf(separator);
            if (index == -1) index = span.Length;
            if (options == StringSplitOptions.RemoveEmptyEntries && index == 0)
            {
                span = span.Slice(1);
                startIndex++;
                continue;
            }

            destination[matches] = new Range(startIndex, startIndex + index);
            matches++;

            if (matches >= destination.Length || span.Length <= index)
               break;

            // move to next token
            span = span.Slice(index + 1);
            startIndex += index + 1;
        }

        return matches;
    }
#endif

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
        char separator = ' ', char? skipSeparatorIfPrecededBy = null)
    {
        // If the destination is empty, there's nothing to do.
        if (destination.IsEmpty)
            return 0;

        int matches = 0, startIndex = 0, offsetIndex = 0;
        bool isEscaping = false;
        char endEscapingChar = '\0';
        ReadOnlySpan<char> searchValues = [separator, '(', '\'', '"'];

        while (span.Length > 0)
        {
            bool isPositiveMatch = true;

            // Remove the spaces that could appear inside a token.
            // Eg: rgb(233, 233, 233) -> rgb(233,233,233)
            int index = isEscaping?
                span.IndexOf(endEscapingChar) :
                span.IndexOfAny(searchValues);

            if (index == -1)
            {
                    // process the last match
                destination[matches] = new Range(startIndex, startIndex + offsetIndex + span.Length);
                matches++;
                break;
            }

            // we find the beginning of an escaping sequence
            var ch = span[index];
            if (ch != separator && !isEscaping)
            {
                if (ch == '(')
                {
                    endEscapingChar = ')';
                    offsetIndex += index + 1;
                }
                else
                {
                    endEscapingChar = ch; // ' or "
                    if (index == 0) startIndex++; // exclude the quote from the captured range
                }
                isEscaping = true;
                isPositiveMatch = false;
            }
            // end of escaping sequence
            else if (ch == endEscapingChar)
            {
                if (ch == ')') index++; // include that closing parenthesis in the range
                isEscaping = false;
            }
            // this is a separator but maybe we will need to skip it
            // eg: "Arial, Verdana bold 1em" -> the space after the comma must be skipped
            else if (ch == separator && index > 0 &&
                skipSeparatorIfPrecededBy.HasValue && span[index -1] == skipSeparatorIfPrecededBy)
            {
                index++;
                offsetIndex += index + 1;
                isPositiveMatch = false;
            }
            else if (index == 0) // empty token
            {
                startIndex++;
                isPositiveMatch = false;
            }

            // index > 0 to exclude empty entries
            if (!isEscaping && index > 0 && isPositiveMatch)
            {
                destination[matches] = new Range(startIndex, startIndex + offsetIndex + index);
                matches++;
                startIndex += index + offsetIndex + 1;
                offsetIndex = 0;
            }

            if (matches >= destination.Length || span.Length <= index)
               break;

            // move to next token
            span = span.Slice(index + 1);
        }

        return matches;
    }
}
