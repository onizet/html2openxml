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

delegate T SpanParser<T>(ReadOnlySpan<char> span);

/// <summary>
/// Parse a CSS shorthand property containing top, right, bottom and left.
/// </summary>
static class CssShorthand
{
    public static bool TryParse<T>(
        ReadOnlySpan<char> span,
        SpanParser<T> parser,
        Func<T, bool> isValid,
        out (T Top, T Right, T Bottom, T Left) result)
    {
        result = default;

        Span<Range> tokens = stackalloc Range[4];
        var count = span.SplitCompositeAttribute(tokens, ' ');

        if (count < 1 || count > 4)
            return false;

        Span<T> values = new T[4];
        for (int i = 0; i < count; i++)
        {
            values[i] = parser(span[tokens[i]]);

            if (!isValid(values[i]))
                return false;
        }

        result = count switch
        {
            1 => (values[0], values[0], values[0], values[0]),
            2 => (values[0], values[1], values[0], values[1]),
            3 => (values[0], values[1], values[2], values[1]),
            _ => (values[0], values[1], values[2], values[3])
        };

        return true;
    }
}