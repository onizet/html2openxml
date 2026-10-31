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

namespace HtmlToOpenXml;

#if !NET5_0_OR_GREATER

/// <summary>
/// Polyfill helper class to provide extension methods for <see cref="ReadOnlySpan{T}"/>.
/// </summary>
static partial class SpanExtensions
{
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
}
#endif
