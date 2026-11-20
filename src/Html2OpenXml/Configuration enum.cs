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

/// <summary>
/// Specifies the page orientation used when creating a new Word document.
/// The orientation determines whether pages are laid out vertically (portrait)
/// or horizontally (landscape).
/// </summary>
public enum PageOrientation
{
    /// <summary>
    /// Pages are taller than they are wide.
    ///
    /// <para>
    /// This is the default orientation for most documents such as letters, reports,
    /// contracts, and articles.
    /// </para>
    /// </summary>
    Portrait,
    /// <summary>
    /// Pages are wider than they are tall.
    ///
    /// <para>
    /// This orientation is commonly used for wide tables, spreadsheets, diagrams,
    /// and content requiring additional horizontal space.
    /// </para>
    /// </summary>
    Landscape
}

/// <summary>
/// Defines a logical insertion point within the document.
/// </summary>
public enum DocumentPosition
{
    /// <summary>
    /// Inserts the content at the beginning of the document.
    /// </summary>
    DocumentStart,
    /// <summary>
    /// Insert the content immediately after the cover page.
    /// If no cover page exists, the content is inserted at the beginning of the document.
    /// </summary>
    AfterCover,
    /// <summary>
    /// Appends the content at the end of the document.
    /// </summary>
    DocumentEnd
}

/// <summary>
/// Defines the paper size used when creating new document page.
/// Sizes correspond to the standard paper formats supported by Microsoft Word and OpenXML.
/// </summary>
public enum PaperSize
{
    /// <summary>
    /// ISO 216 A3 paper size (297 × 420 mm).
    /// Commonly used for large diagrams, drawings, and wide reports.
    /// </summary>
    A3,

    /// <summary>
    /// ISO 216 A4 paper size (210 × 297 mm).
    /// The default paper size for most business documents.
    /// </summary>
    A4,

    /// <summary>
    /// ISO 216 A5 paper size (148 × 210 mm).
    /// Often used for booklets, manuals, and brochures.
    /// </summary>
    A5,

    /// <summary>
    /// ISO 216 A6 paper size (105 × 148 mm).
    /// Commonly used for postcards, invitations, and small leaflets.
    /// </summary>
    A6,

    /// <summary>
    /// ISO 216 B3 paper size (353 × 500 mm).
    /// Larger than A3 and typically used for posters and technical drawings.
    /// </summary>
    B3,

    /// <summary>
    /// ISO 216 B4 paper size (250 × 353 mm).
    /// Commonly used for magazines and larger printed documents.
    /// </summary>
    B4,

    /// <summary>
    /// ISO 216 B5 paper size (176 × 250 mm).
    /// Frequently used for books and notebooks.
    /// </summary>
    B5,

    /// <summary>
    /// ISO 216 B6 paper size (125 × 176 mm).
    /// Commonly used for pocket-sized books and booklets.
    /// </summary>
    B6,

    /// <summary>
    /// US Letter paper size (8.5 × 11 in / 216 × 279 mm).
    /// The standard paper size used in North America.
    /// </summary>
    Letter,

    /// <summary>
    /// US Legal paper size (8.5 × 14 in / 216 × 356 mm).
    /// Often used for legal and contractual documents.
    /// </summary>
    Legal,
}
