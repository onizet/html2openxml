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
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml;

/// <summary>
/// Provides helper methods for creating Word documents configured for use with <see cref="HtmlConverter"/>. 
///
/// <para>
/// This factory simplifies the creation of new DOCX documents by configuring the required 
/// OpenXML package structure, document parts, and page settings.
/// The generated documents can be used directly with <see cref="HtmlConverter"/> to convert HTML content into WordProcessingML.
/// </para>
///
/// </summary>
/// <example>
/// Basic quick start:
///
/// <code>
/// await using var generatedDocument = new MemoryStream();
/// using var package = WordDocumentFactory.Create(generatedDocument, PageOrientation.Portrait);
/// HtmlConverter converter = new(package);
/// await converter.ParseBody(html);
/// </code>
/// </example>
public static class WordDocumentFactory
{
    const string BuiltInGalleryCoverPageName = "Cover Pages";
    private const uint A4PortraitWidth = 11906U;
    private const uint A4PortraitHeight = 16838U;

    /// <summary>
    /// Create a new Word document on the specified stream and configures its page settings.
    /// 
    /// <para>
    ///The created document contains the minimum required OpenXML structure, including a main document part
    /// and an empty body. The returned document can be passed directly to <see cref="HtmlConverter"/> or
    /// further customised using the Open XML SDK.
    /// </para>
    /// </summary>
    /// <param name="stream">The destination stream that will receive the generated DOCX package
    /// The stream must be writable.</param>
    /// <param name="orientation">The page orientation to apply to the document.</param>
    /// <exception cref="ArgumentNullException">Thrown when "stream" is null reference.</exception>
    /// <exception cref="IOException">Thrown when "stream" is not opened with Write access.</exception>
    public static WordprocessingDocument Create(Stream stream,
        PageOrientation orientation = PageOrientation.Portrait)
    {
        var package = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);

        var mainPart = GetOrCreateMainDocumentPart(package);
        var sectionProperties = CreateSectionProperties(orientation);
        mainPart.Document!.Body!.AddChild(sectionProperties);
        return package;
    }

    /// <summary>
    /// Returns the document cover page, if one exists.
    /// <para>
    /// The cover page is represented using a Word content control associated with
    /// the built-in Cover Page gallery.
    /// </para>
    /// </summary>
    /// <returns>
    /// The cover page content control, or <see langword="null"/> when the document
    /// does not contain a cover page.
    /// </returns>
    public static SdtBlock? GetCoverPage(this HtmlConverter converter)
    {
        var body = GetOrCreateDocumentBody(converter.mainPart);
        foreach (var stdBlock in body.Descendants<SdtBlock>())
        {
            if (stdBlock.SdtProperties?
                .GetFirstChild<SdtContentDocPartObject>()?
                .DocPartGallery?.Val?.Value == BuiltInGalleryCoverPageName)
            {
                return stdBlock;
            }
        }

        return null;
    }

    /// <summary>
    /// Creates or replaces the document cover page.
    /// <para>
    /// The supplied HTML is converted using the current <see cref="HtmlConverter"/> configuration
    /// and stored inside a Word cover page content control.
    /// </para>
    /// <para>
    /// If a cover page already exists, its content is replaced while the existing content control is preserved.
    /// </para>
    /// </summary>
    /// <param name="converter">The converter owning the target document.</param>
    /// <param name="html">The HTML content to render into the cover page.</param>
    /// <param name="cancellationToken">The cancellation token.</param>
    public static async Task SetCoverPage(this HtmlConverter converter, string html,
        CancellationToken cancellationToken = default)
    {
        SdtBlock? coverStdBlock = converter.GetCoverPage();
        if (coverStdBlock is null)
        {
            coverStdBlock = new();
            converter.mainPart.Document!.Body!.InsertAt(coverStdBlock, 0);
        }

        coverStdBlock.SdtProperties ??= new();
        var partObject = coverStdBlock.SdtProperties.GetOrAddFirstChild<SdtContentDocPartObject>();
        partObject.DocPartGallery ??= new() { Val = BuiltInGalleryCoverPageName };
        partObject.DocPartUnique ??= new();
        coverStdBlock.SdtContentBlock?.Remove();

        var elements = await converter.ParseAsync(html, cancellationToken);
        coverStdBlock.SdtContentBlock = new(elements);
    }

    /// <summary>
    /// Gets the main document part associated with the document, creating it when it does not already exist.
    /// </summary>
    internal static MainDocumentPart GetOrCreateMainDocumentPart(WordprocessingDocument package)
    {
        var mainPart = package.MainDocumentPart;
        if (mainPart is null)
        {
            mainPart = package.AddMainDocumentPart();
            new Document(new Body()).Save(mainPart);
        }

        return mainPart;
    }

    /// <summary>
    /// Gets the document body part associated with the document, creating it when it does not already exist.
    /// </summary>
    internal static Body GetOrCreateDocumentBody(MainDocumentPart mainPart)
    {
        if (mainPart.Document == null)
            new Document(new Body()).Save(mainPart);
        else if (mainPart.Document.Body == null)
            mainPart.Document.Body = new Body();

        return mainPart.Document!.Body!;
    }

    /// <summary>
    /// Generate the required OpenXml element for handling page orientation.
    /// </summary>
    internal static SectionProperties CreateSectionProperties(PageOrientation orientation)
    {
        PageSize pageSize;
        if (orientation == PageOrientation.Portrait)
        {
            pageSize = new() {
                Width = A4PortraitWidth, Height = A4PortraitHeight,
                Orient = PageOrientationValues.Portrait
            };
        }
        else
        {
            pageSize = new() {
                Width = A4PortraitHeight, Height = A4PortraitWidth,
                Orient = PageOrientationValues.Landscape
            };
        }

        return new SectionProperties (
            pageSize,
            new PageMargin() {
                Top = 1440, Right = (UInt32Value) 1440U, Bottom = 1440, Left = (UInt32Value) 1440U,
                Header = (UInt32Value) 720U, Footer = (UInt32Value) 720U, Gutter = (UInt32Value) 0U
            },
            new Columns() { Space = "720" },
            new DocGrid() { LinePitch = 360 }
        );
    }

    /// <summary>
    /// Determines whether the specified page size represents a landscape orientation.
    /// </summary>
    internal static bool IsLandscape(PageSize? pageSize)
    {
        if (pageSize is null)
            return false;

        return pageSize.Width?.Value >= A4PortraitHeight;
    }
}
