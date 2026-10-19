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

namespace HtmlToOpenXml.IO;

/// <summary>
/// Represents a Data URI containing inline content encoded directly in the URI string,
/// typically using Base64 encoding.
/// 
/// Data URIs are commonly used to embed images in HTML without referencing an external
/// resource, for example: <c>data:image/png;base64,...</c>
/// </summary>
[System.Diagnostics.DebuggerDisplay("{Mime,nq}")]
public sealed class DataUri
{
    private DataUri(string mimeType)
    {
        Mime = mimeType;
    }

    /// <summary>
    /// Parse an instance of the Data URI scheme.
    /// </summary>
    /// <param name="uri">The inline Data URI to parse.</param>
    /// <param name="result">When this method returns, contains a DataUri constructed
    /// from <paramref name="uri"/>. This parameter is passed uninitialized.</param>
    /// <returns>A <see cref="Boolean"/> value that is true if the DataUri was 
    /// successfully created; otherwise, false.</returns>
    public static bool TryCreate(string uri, [NotNullWhen(true)] out DataUri? result)
    {
        // expected format: data:[<MIME-type>][;charset=<encoding>][;base64],<data>
        // The encoding is indicated by ;base64. If it's present the data is encoded as base64. Without it the data (as a sequence of octets)
        // is represented using ASCII encoding for octets inside the range of safe URL characters and using the standard %xx hex encoding
        // of URLs for octets outside that range. If <MIME-type> is omitted, it defaults to text/plain;charset=US-ASCII.
        // (As a shorthand, the type can be omitted but the charset parameter supplied.)
        // Some browsers (Chrome, Opera, Safari, Firefox) accept a non-standard ordering if both ;base64 and ;charset are supplied,
        // while Internet Explorer requires that the charset's specification must precede the base64 token.
        // http://en.wikipedia.org/wiki/Data_URI_scheme

        result = null;

        if (uri.Length < 5)
            return false;

        var uriSpan = uri.AsSpan();
        int commaIndex = uriSpan.IndexOf(',');

        // 6 is minimal valid data URI: `data:,`
        // 128 to avoid allocating too much memory
        if (commaIndex < 5 || commaIndex > 128)
            return false;

        Span<char> loweredValue = stackalloc char[commaIndex];
        uriSpan.Slice(0, commaIndex).ToLowerInvariant(loweredValue);

        // must starts with `data:`
        if (!loweredValue.StartsWith("data:"))
            return false;

        // if mime is omitted, set default value as it stands in the norm
        string mime = "text/plain";
        string charSet = "ascii";
        bool isBase64 = false;

        Span<Range> tokens = stackalloc Range[5];
        var tokenCount = loweredValue.Split(tokens, ';', StringSplitOptions.RemoveEmptyEntries);
        for (var i = 0; i < tokenCount; i++)
        {
            var token = loweredValue[tokens[i]];
            if (token is "base64") isBase64 = true;
            else if (token.StartsWith("data:") && token.Length > 6)
            {
                mime = token.Slice(5).Trim().ToString();
            }
            else if (token.StartsWith("charset="))
            {
                try
                {
                    charSet = System.Text.Encoding.GetEncoding(token.Slice(8 /* 'charset=' */).ToString()).WebName;
                }
                catch (ArgumentException)
                {
                    // charSet was not recognized
                    return false;
                }
            }

            // else ignore, arbitrary value not needed in this context
        }

        var data = uriSpan.Slice(commaIndex + 1).Trim();
        if (data.IsEmpty)
        {
            result = new(mime);
            return true; // this is still a valid data URI
        }

        if (isBase64)
        {
            // be careful that the raw data is encoded for url (standard %xx hex encoding)
            string base64 = HttpUtility.HtmlDecode(data);

            try
            {
                result = new(mime) { Data = Convert.FromBase64String(base64) };
                return true;
            }
            catch (FormatException)
            {
                // Base64 data is invalid
                return false;
            }
        }

        // the <data> represents some text (like html snippet) and must be decoded.
        result = new(mime) { Encoding = charSet, Content = HttpUtility.UrlDecode(data) };
        return true;
    }

    /// <summary>
    /// Determines whether the uri looks like a Data Uri.
    /// </summary>
    public static bool IsDataUri(string uri)
    {
        return uri.StartsWith("data:", StringComparison.OrdinalIgnoreCase);
    }

    //____________________________________________________________________
    //

    /// <summary>
    /// Gets the MIME type declared by the Data URI, such as image/png or image/jpeg.
    /// </summary>
    public string Mime { get; }

    /// <summary>
    /// Gets the decoded binary content.
    /// </summary>
    public byte[]? Data { get; init; }

    /// <summary>
    /// Gets the encoding of the content when the data is text (e.g., UTF-8).
    /// </summary>
    public string? Encoding { get; init; }

    /// <summary>
    /// Gets the decoded human-readable content.
    /// </summary>
    public string? Content { get; init; }
}
