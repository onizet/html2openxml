/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
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
    internal static readonly PageSize A4PageSize = new() { Width = 11906U, Height = 16838U };


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
    /// <param name="paperSize">The paper size of the first page.</param>
    /// <exception cref="ArgumentNullException">Thrown when "stream" is null reference.</exception>
    /// <exception cref="IOException">Thrown when "stream" is not opened with Write access.</exception>
    public static WordprocessingDocument Create(Stream stream,
        PageOrientation orientation = PageOrientation.Portrait,
        PaperSize paperSize = PaperSize.A4)
    {
        var package = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);

        var mainPart = GetOrCreateMainDocumentPart(package);
        var sectionProperties = CreateSectionProperties(orientation, paperSize);
        mainPart.Document!.Body!.AppendChild(sectionProperties);
        return package;
    }

    /// <summary>
    /// Starts a new page in the document using the specified paper size and orientation.
    ///
    /// <para>
    /// The new page is created as a separate Word section, allowing the
    /// page size and orientation to differ from the surrounding content.
    /// </para>
    /// </summary>
    /// <param name="converter">The converter owning the target document.</param>
    /// <param name="orientation">The orientation of the new page.</param>
    /// <param name="paperSize">The paper size of the new page.</param>
    public static void AddPage(this HtmlConverter converter, PageOrientation orientation, PaperSize paperSize = PaperSize.A4)
    {
        var body = GetOrCreateDocumentBody(converter.mainPart);
        var sectionProperties = body.GetLastChild<SectionProperties>();
        var pageSize = sectionProperties?.GetFirstChild<PageSize>();

        if (pageSize is null)
        {
            converter.AddPageBreak();
            body.AppendChild(CreateSectionProperties(orientation, paperSize));
            return;
        }

        var (w, h) = GetPageDimensions(paperSize);
        if (orientation == PageOrientation.Landscape)
            (w, h) = (h, w);

        if (pageSize.Width?.Value != w || pageSize.Height?.Value != h)
        {
            // detach from parent
            sectionProperties!.Remove();

            body.AppendChild(new Paragraph(
                new ParagraphProperties(
                    sectionProperties
                ),
                new Run(
                   new LastRenderedPageBreak())));

            body.AppendChild(CreateSectionProperties(orientation, paperSize));
        }
        else
        {
            if (sectionProperties!.PreviousSibling() is Paragraph lastPara)
            {
                lastPara.AppendChild(new Run(
                    new Break { Type = BreakValues.Page }));
            }
            else
            {
                body.AppendChild(new Paragraph(
                    new Run(
                        new Break { Type = BreakValues.Page })));
            }
        }
    }

    /// <summary>
    /// Inserts a page break immediately after the specified element.
    /// </summary>
    /// <param name="converter">The converter owning the target document.</param>
    public static void AddPageBreak(this HtmlConverter converter)
    {
        var pageBreak = new Paragraph(
            new Run(
                new Break { Type = BreakValues.Page })
        );
        var body = GetOrCreateDocumentBody(converter.mainPart);
        var sectionProperties = body.GetLastChild<SectionProperties>();
        body.InsertBefore(pageBreak, sectionProperties);
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

    internal static SectionProperties GetOrCreateSectionProperties(MainDocumentPart mainPart,
        PageOrientation orientation = PageOrientation.Portrait, PaperSize paperSize = PaperSize.A4)
    {
        var body = GetOrCreateDocumentBody(mainPart);
        var sectionProperties = body.GetLastChild<SectionProperties>();
        if (sectionProperties is null)
            body.AppendChild(sectionProperties = CreateSectionProperties(orientation, paperSize));

        return sectionProperties;
    }

    /// <summary>
    /// Generate the required OpenXml element for handling page orientation.
    /// </summary>
    internal static SectionProperties CreateSectionProperties(PageOrientation orientation, PaperSize paperSize)
    {
        var (width, height) = GetPageDimensions(paperSize);
        PageSize pageSize = new() { Width = width, Height = height };

        if (orientation == PageOrientation.Landscape)
            pageSize.FlipPage();

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

    private static (uint Width, uint Height) GetPageDimensions(PaperSize paperSize)
    {
        return paperSize switch {
            PaperSize.A6 => (5953, 8391),
            PaperSize.A5 => (8391, 11906),
            PaperSize.A4 => (11906, 16838),
            PaperSize.A3 => (16838, 23811),

            PaperSize.B6 => (7087, 9978),
            PaperSize.B5 => (9978, 14173),
            PaperSize.B4 => (14173, 20013),
            PaperSize.B3 => (20013, 28346),

            PaperSize.Letter => (12240, 15840),
            PaperSize.Legal => (12240, 20160),

            _ => throw new ArgumentOutOfRangeException(nameof(paperSize))
        };
    }
}
