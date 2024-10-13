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
using System.Text;
using System.Text.RegularExpressions;

namespace HtmlToOpenXml.IO;

/// <summary>
/// Represents an URI that includes inline data as if they were external resources.
/// </summary>
[System.Diagnostics.DebuggerDisplay("{Mime,nq}")]
public sealed class DataUri
{
    private readonly static Regex dataUriRegex = new Regex(
            @"data\:(?<mime>\w+/\w+)?(?:;charset=(?<charset>[a-zA-Z_0-9-]+))?(?<base64>;base64)?,(?<data>.*)",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);

    private DataUri(string mime, byte[] data)
    {
        this.Mime = mime;
        this.Data = data;
    }

    /// <summary>
    /// Parse an instance of the Data URI scheme.
    /// </summary>
    /// <param name="uri">The inline Data URI to parse.</param>
    /// <param name="result">When this method returns, contains a DataUri constructed
    /// from <paramref name="uri"/>. This parameter is passed uninitialized.</param>
    /// <returns>A <see cref="Boolean"/> value that is true if the DataUri was 
    /// successfully created; otherwise, false.</returns>
    public static bool TryCreate(string uri, out DataUri? result)
    {
        // expected format: data:[<MIME-type>][;charset=<encoding>][;base64],<data>
        // The encoding is indicated by ;base64. If it's present the data is encoded as base64. Without it the data (as a sequence of octets)
        // is represented using ASCII encoding for octets inside the range of safe URL characters and using the standard %xx hex encoding
        // of URLs for octets outside that range. If <MIME-type> is omitted, it defaults to text/plain;charset=US-ASCII.
        // (As a shorthand, the type can be omitted but the charset parameter supplied.)
        // Some browsers (Chrome, Opera, Safari, Firefox) accept a non-standard ordering if both ;base64 and ;charset are supplied,
        // while Internet Explorer requires that the charset's specification must precede the base64 token.
        // http://en.wikipedia.org/wiki/Data_URI_scheme

        // We will stick for IE compliance for the moment...

        Match match = dataUriRegex.Match(uri);
        result = null;

        if (!match.Success) return false;

        byte[] rawData;
        string mime;
        Encoding charSet = Encoding.ASCII;

        // if mime is omitted, set default value as it stands in the norm
        if (match.Groups["mime"].Length == 0)
            mime = "text/plain";
        else
            mime = match.Groups["mime"].Value;

        if (match.Groups["charset"].Length > 0)
        {
            try
            {
                charSet = Encoding.GetEncoding(match.Groups["charset"].Value);
            }
            catch (ArgumentException)
            {
                // charSet was not recognized
                return false;
            }
        }

        // is it encoded in base64?
        if (match.Groups["base64"].Length > 0)
        {
            // be careful that the raw data is encoded for url (standard %xx hex encoding)
#if NET5_0_OR_GREATER
            string base64 = System.Web.HttpUtility.HtmlDecode(match.Groups["data"].Value);
#else 
            string base64 = HttpUtility.HtmlDecode(match.Groups["data"].Value);
#endif

            try
            {
                rawData = Convert.FromBase64String(base64);
            }
            catch (FormatException)
            {
                // Base64 data is invalid
                return false;
            }
        }
        else
        {
            // the <data> represents some text (like html snippet) and must be decoded.
            string? raw = HttpUtility.UrlDecode(match.Groups["data"].Value)!;
            if (string.IsNullOrEmpty(raw))
                return false;
            // we convert back to UTF-8 for easier processing later and to have a "referential" encoding
            rawData = Encoding.Convert(charSet, Encoding.UTF8, charSet.GetBytes(raw));
        }

        result = new DataUri(mime, rawData);
        return true;
    }

    /// <summary>
    /// Indicates whether the string is well-formed by attempting to construct a DataUri with the string.
    /// </summary>
    public static bool IsWellFormed(string uri)
    {
        return dataUriRegex.IsMatch(uri);
    }

    //____________________________________________________________________
    //

    /// <summary>
    /// Gets the MIME type of the encoded data.
    /// </summary>
    public string Mime { get; private set; }

    /// <summary>
    /// Gets the decoded data.
    /// </summary>
    public byte[] Data { get; private set; }
}
