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

/// <summary>
/// Controls where the definition of an acronym or abbreviation is rendered during conversion.
/// 
/// AcronymFor example, the definition assocaited with an HTML <c>abbr</c> element can be
/// rendered at the end of the page or collected at the end of the document.
/// </summary>
public enum AcronymPosition
{
    /// <summary>
    /// Collect acronym definitions at the end of the current page.
    /// </summary>
    PageEnd = 0,
    /// <summary>
    /// Collect acronym definitions at the end of the document.
    /// </summary>
    DocumentEnd = 1,
}

/// <summary>
/// Predefined quote style as defined by the browser (used for the <c>q</c> HTML tag).
/// </summary>
/// <remarks>
/// Initializes a new instance of <see cref="QuoteChars"/> class.
/// </remarks>
/// <param name="begin">The characters at the beginning of a quote.</param>
/// <param name="end">The characters at the end of a quote.</param>
public readonly struct QuoteChars(string begin, string end)
{
    /// <summary>Internet Explorer style: « abc » </summary>
    public static readonly QuoteChars IE = new("« ", " »");
    /// <summary>Firefox style: “abc”</summary>
    public static readonly QuoteChars Gecko = new("“", "”");
    /// <summary>Chrome/Safari/Opera style: "abc"</summary>
    public static readonly QuoteChars WebKit = new("\"", "\"");

    internal string Prefix { get; } = begin;
    internal string Suffix { get; } = end;
}

/// <summary>
/// Controls whether images are embedded in the generated document
/// or kept as external references.
/// </summary>
public enum ImageProcessingMode
{
    /// <summary>
    /// Downloads and embeds all images into the document (default behaviour).
    /// This creates self-contained documents but may result in large file sizes.
    /// </summary>
    Embed = 0,
    /// <summary>
    /// Links to external images via external relationships instead of downloading them.
    /// This keeps document size small but images won't display offline or if URLs become unavailable.
    /// Data URI images (base64 encoded) are still embedded.
    /// </summary>
    LinkExternal = 1,
    /// <summary>
    /// Only embeds data URI images (base64 encoded inline images).
    /// External images (http/https/file) are skipped entirely.
    /// </summary>
    EmbedDataUriOnly = 2,
}
