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
using System.Buffers;
using System.Globalization;

namespace HtmlToOpenXml;

/// <summary>
/// Custom utility providing compatible encoding and decoding functionality.
/// This implementation embraces the modern API <c>Span&lt;char&gt;</c> for maximum efficiency.
/// 
/// <para>
/// For .Net 462 and .Net Standard 2.0, this implementation avoid reliance on the legacy <c>System.Web</c> library.
/// </para>
/// </summary>
static partial class HttpUtility
{
    /// <summary>
    /// Converts a string that has been HTML-encoded for HTTP transmission into a decoded string.
    /// </summary>
    public static string HtmlDecode(ReadOnlySpan<char> span)
    {
        if (span.IsEmpty)
            return string.Empty;

        if (span.IndexOf('&') < 0)
            return span.ToString();

        var buffer = ArrayPool<char>.Shared.Rent(span.Length);
        var length = 0;

        while (!span.IsEmpty)
        {
            int ampIndex = span.IndexOf('&'),
                endIndex;

            if (ampIndex == -1 || ((endIndex = span.Slice(ampIndex+1).IndexOfAny(';', '&')) == -1))
            {
                // Dangling '&' at EOF or sequence end. Treat it literally.
                span.CopyTo(buffer.AsSpan(length));
                length += span.Length;
                break;
            }

            endIndex += ampIndex + 1;

            // have we found another '&' while expecting ';'? If so, this is not an entity
            // but this new ampersan might be
            if (span[endIndex] == '&')
            {
                span.Slice(0, endIndex).CopyTo(buffer.AsSpan(length));
                length += endIndex;
                span = span.Slice(endIndex);
                continue;
            }

            span.Slice(0, ampIndex).CopyTo(buffer.AsSpan(length));
            length += ampIndex;

            var entity = span.Slice(ampIndex + 1, endIndex - ampIndex - 1);
            if (entity.Length > 0 && entity[0] == '#')
            {
                if (TryDecodeNumericEntity(entity, out uint codePoint))
                {
                    if (System.Text.Rune.IsValid(codePoint))
                    {
                        length += new System.Text.Rune(codePoint)
                            .EncodeToUtf16(buffer.AsSpan(length));
                    }

                    span = span.Slice(entity.Length + 1);
                    continue;
                }

                buffer[length++] = '&';
                span = span.Slice(1);
                continue;
            }

            if (IsKnownEntities(entity, out var ch))
                buffer[length++] = ch;
            else
                span.Slice(endIndex).CopyTo(buffer);

            // move to next token
            span = span.Slice(endIndex + 1);
        }

        var result = new String(buffer, 0, length);

        ArrayPool<char>.Shared.Return(buffer);
        return result;
    }

    /// <summary>
    /// Try to parse an HTML nnumeric entity and returns its Unicode code point.
    /// </summary>
    /// <param name="entity">The entity text without the surrounding <c>&amp;</c> and <c>;</c>.</param>
    /// <param name="codePoint">The Unicode point representing that entity.</param>
    private static bool TryDecodeNumericEntity(ReadOnlySpan<char> entity, out uint codePoint)
    {
        if (entity.Length < 2 || entity[0] != '#')
        {
            codePoint = 0;
            return false;
        }

        NumberStyles numberStyles = NumberStyles.Integer;
        if (entity[1] is 'x' or 'X')
        {
            numberStyles = NumberStyles.AllowHexSpecifier;
            entity = entity.Slice(2);
        }
        else
        {
            entity = entity.Slice(1);
        }

        return uint.TryParse(
#if NET5_0_OR_GREATER
            entity,
#else
            entity.ToString(),
#endif
            numberStyles, CultureInfo.InvariantCulture, out codePoint);
    }

    private static bool IsKnownEntities(ReadOnlySpan<char> entity, out char ch)
    {
        // avoid to allocate a new string for most common entities
        // and this is faster to use this pattern matching
        ch = entity switch {
            "amp" => '&',
            "lt"  => '<',
            "gt"  => '>',
            "quot" => '"',
            "apos" => '\'',
            "nbsp" => '\x00a0',
            "copy" => '©',
            "reg" => '®',
            "trade" => '™',
            _ => '\0'
        };

        if (ch != '\0')
            return true;

        // fallback on the hard-coded list
        return HtmlEntities.Lookup(entity, out ch);
    }

    /// <summary>
    /// Converts a string that represents an Html-encoded URL to a decoded string.
    /// </summary>
    public static string UrlDecode(ReadOnlySpan<char> text)
    {
        if (text.IsEmpty) return string.Empty;

        // pre-process for + sign space formatting since System.Uri doesn't handle it
        // plus literals are encoded as %2b normally so this should be safe
        // http://www.west-wind.com/weblog/posts/2009/Feb/05/Html-and-Uri-String-Encoding-without-SystemWeb
        if (text.IndexOf('+') > 0)
            return Uri.UnescapeDataString(text.ToString().Replace("+", " "));

#if NET9_0_OR_GREATER
        return Uri.UnescapeDataString(text);
#else
        return Uri.UnescapeDataString(text.ToString());
#endif
    }
}