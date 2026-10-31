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
