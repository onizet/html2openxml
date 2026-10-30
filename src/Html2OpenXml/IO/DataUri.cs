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
using System.Diagnostics.CodeAnalysis;
using System.Globalization;

namespace HtmlToOpenXml.IO;

/// <summary>
/// Represents a Data URI containing inline content encoded directly in the URI string,
/// typically using Base64 encoding.
/// 
/// Data URIs are commonly used to embed images in HTML without referencing an external
/// resource, for example: <c>data:image/png;base64,...</c>
/// </summary>
/// <param name="MimeType">The MIME type declared by the Data URI, such as image/png or image/jpeg.</param>
/// <param name="Charset">The encoding of the content when the data is text (e.g., UTF-8).</param>
/// <param name="Data">The decoded binary content.</param>
/// <param name="Content">The decoded human-readable content.</param>
readonly record struct DataUri(string MimeType, string? Charset = null, byte[]? Data = null, string? Content = null)
{
    public static readonly DataUri Invalid = new(null!);

    /// <summary>
    /// Try to parse an instance of the Data URI scheme but stop if it detects that the URI does not contain an embedded image.
    /// This method avoids excessive memory allocations, as the purpose of this library is reading base64 images.
    /// </summary>
    /// <param name="uri">The inline Data URI to parse.</param>
    /// <param name="result">When this method returns, contains a DataUri constructed
    /// from <paramref name="uri"/>. This parameter is passed uninitialized.</param>
    /// <returns>Returns <see langword="true"/> if the DataUri was successfully created;
    /// otherwise, <see langword="false"/>.</returns>
    public static bool TryCreateImage(string uri, [NotNullWhen(true)] out DataUri result)
    {
        return TryCreate(uri, breakIfNotImage: true, out result);
    }

    /// <summary>
    /// Parse an instance of the Data URI scheme.
    /// </summary>
    /// <param name="uri">The inline Data URI to parse.</param>
    /// <param name="result">When this method returns, contains a DataUri constructed
    /// from <paramref name="uri"/>. This parameter is passed uninitialized.</param>
    /// <returns>Returns <see langword="true"/> if the DataUri was successfully created;
    /// otherwise, <see langword="false"/>.</returns>
    public static bool TryCreate(string uri, [NotNullWhen(true)] out DataUri result)
    {
        return TryCreate(uri, breakIfNotImage: false, out result);
    }

    private static bool TryCreate(string uri, bool breakIfNotImage, [NotNullWhen(true)] out DataUri result)
    {
        // expected format: data:[<MIME-type>][;charset=<encoding>][;base64],<data>
        // The encoding is indicated by ;base64. If it's present the data is encoded as base64. Without it the data (as a sequence of octets)
        // is represented using ASCII encoding for octets inside the range of safe URL characters and using the standard %xx hex encoding
        // of URLs for octets outside that range. If <MIME-type> is omitted, it defaults to text/plain;charset=US-ASCII.
        // (As a shorthand, the type can be omitted but the charset parameter supplied.)
        // Some browsers (Chrome, Opera, Safari, Firefox) accept a non-standard ordering if both ;base64 and ;charset are supplied,
        // while Internet Explorer requires that the charset's specification must precede the base64 token.
        // http://en.wikipedia.org/wiki/Data_URI_scheme

        result = Invalid;

        if (uri.Length < 5)
            return false;

        var uriSpan = uri.AsSpan();
        int commaIndex = uriSpan.IndexOf(',');

        // 6 is minimal valid data URI: `data:,`
        // 128 to avoid allocating too much memory
        if (commaIndex < 5 || commaIndex > 128)
            return false;

        // must starts with `data:`
        if (!IsDataUri(uriSpan))
            return false;

        // if mime is omitted, set default value as it stands in the norm
        string mime = "text/plain";
        string charset = "ascii";
        System.Text.Encoding encoding = System.Text.Encoding.ASCII;
        bool isBase64 = false;

        Span<Range> tokens = stackalloc Range[8];
        var tokenCount = uriSpan.Slice(0, commaIndex).Split(tokens, ';', StringSplitOptions.RemoveEmptyEntries);
        for (var i = 0; i < tokenCount; i++)
        {
            var token = uriSpan[tokens[i]];
            if (token.Equals("base64", StringComparison.OrdinalIgnoreCase))
            { 
                isBase64 = true;
            }
            else if (token.StartsWith("data:", StringComparison.OrdinalIgnoreCase) && token.Length > 6)
            {
                mime = token.Slice(5).Trim().ToString();
            }
            else if (token.StartsWith("charset=", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    charset = token.Slice(8 /* 'charset=' */).ToString();
                    encoding = System.Text.Encoding.GetEncoding(charset);
                }
                catch (ArgumentException)
                {
                    // charSet was not recognized
                    return false;
                }
            }

            // else ignore, arbitrary value not needed in this context
        }

        // ignoreCase is not needed as this span was lowered
        if (breakIfNotImage && !mime.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
            return false;

        // use the uriSpan to preserve the casing for the data segment
        var data = uriSpan.Slice(commaIndex + 1).Trim();
        if (data.IsEmpty)
        {
            result = new(mime, charset);
            return true; // this is still a valid data URI
        }

        if (isBase64)
        {
            byte[] dataBinary = DecodeBase64(data);
            result = new(mime, Data: dataBinary);
            return dataBinary.Length > 0;
        }

        // the <data> represents some text (like html snippet)
        // and must be decoded respecting the provided charset.
        try
        {
            var bytes = UnescapePercentEncodedBytes(data);
            result = new(mime, charset, Content: encoding.GetString(bytes));
        }
        catch (FormatException)
        {
            return false;
        }
        return true;
    }

    private static byte[] DecodeBase64(ReadOnlySpan<char> value)
    {
        // 4 is the minimal length for a base64 string
        if (value.Length < 4)
            return [];

#if NET5_0_OR_GREATER
        int decodedLength = (value.Length * 3 + 3) / 4;
        if (value[^1] == '=')
        {
            decodedLength -= value[^2] == '='? 2 : 1;
        }

        var buffer = GC.AllocateUninitializedArray<byte>(decodedLength);
        if (!Convert.TryFromBase64Chars(value, buffer, out int bytesWritten))
            return [];

        // mainly for multiline data URI where we waste bytes
        if (buffer.Length != bytesWritten)
            Array.Resize(ref buffer, bytesWritten);
        return buffer;
#else
        try
        {
            return Convert.FromBase64String(value.ToString());
        }
        catch (FormatException)
        {
            // Base64 data is invalid
            return [];
        }
#endif
    }

    /// <summary>
    /// Convert the data portion of a non-base64 Data URI from percent-encoded form (%xx) back into its original bytes.
    /// </summary>
    private static byte[] UnescapePercentEncodedBytes(ReadOnlySpan<char> value)
    {
        if (value.IsEmpty)
            return [];

#if NET5_0_OR_GREATER
        byte[] buffer = GC.AllocateUninitializedArray<byte>(value.Length);
#else
        byte[] buffer = new byte[value.Length];
#endif
        int bytesWritten = 0;

        for (int i = 0; i < value.Length; i++)
        {
            if (value[i] == '%' && i + 2 < value.Length &&
                byte.TryParse(
#if NET5_0_OR_GREATER
                    value.Slice(i + 1, 2),
#else
                    value.Slice(i + 1, 2).ToString(),
#endif
                    NumberStyles.HexNumber,
                    CultureInfo.InvariantCulture, out var b))
            {
                buffer[bytesWritten++] = b;
                i += 2;
            }
            else
            {
                if (value[i] > 0xFF)
                    throw new FormatException("Invalid percent-encoded data.");

                buffer[bytesWritten++] = (byte) value[i];
            }
        }

        return bytesWritten == value.Length? buffer :
            buffer.AsSpan(0, bytesWritten).ToArray();
    }

    /// <summary>
    /// Determines whether the uri looks like a Data Uri.
    /// </summary>
    public static bool IsDataUri(ReadOnlySpan<char> uri)
    {
        return uri.StartsWith("data:", StringComparison.OrdinalIgnoreCase);
    }
}
