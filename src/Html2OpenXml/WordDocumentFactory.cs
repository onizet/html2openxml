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
using System.Globalization;
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
    const string BuiltInGalleryTOCName = "Table of Contents";

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
        mainPart.Document!.Body!.AddChild(sectionProperties);
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
            converter.AddPageBreakAfter(null);
            body.AppendChild(CreateSectionProperties(orientation, paperSize));
            return;
        }

        var (w, h) = GetPageDimensions(paperSize);
        if (orientation == PageOrientation.Landscape)
            (w, h) = (h, w);

        if (pageSize.Width?.Value != w || pageSize.Height?.Value != h)
        {
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
    /// Inserts a page break immediately before the specified element.
    ///
    /// <para>
    /// When <paramref name="element"/> is <see langword="null"/>, the page break
    /// is inserted at the beginning of the document.
    /// </para>
    /// </summary>
    /// <param name="converter">The converter owning the target document.</param>
    /// <param name="element">The element that should be preceded by a page break.</param>
    public static void AddPageBreakBefore(this HtmlConverter converter, OpenXmlElement? element)
    {
        var pageBreak = CreatePageBreak();
        GetOrCreateDocumentBody(converter.mainPart).InsertBefore(pageBreak, element);
    }

    /// <summary>
    /// Inserts a page break immediately after the specified element.
    ///
    /// <para>
    /// When <paramref name="element"/> is <see langword="null"/>, the page break
    /// is inserted at the end of the document.
    /// </para>
    /// </summary>
    /// <param name="converter">The converter owning the target document.</param>
    /// <param name="element">
    /// The element that should be followed by a page break.
    /// </param>
    public static void AddPageBreakAfter(this HtmlConverter converter, OpenXmlElement? element)
    {
        var pageBreak = CreatePageBreak();
        var body = GetOrCreateDocumentBody(converter.mainPart);
        if (element is null)
            body.InsertOrMove(pageBreak, DocumentPosition.DocumentEnd);
        else
            body.InsertAfter(pageBreak, element);
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
        => GetOrCreateDocumentBody(converter.mainPart).FindNamedSdtBlock(BuiltInGalleryCoverPageName);

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
    internal static SdtBlock? GetCoverPage(this Body body)
        => body.FindNamedSdtBlock(BuiltInGalleryCoverPageName);

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
        var fragment = await converter.LoadDocumentAsync(html, cancellationToken);
        await converter.SetCoverPage(fragment, cancellationToken);
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
    /// <param name="fragment">A parsed HTML fragment to render into the cover page.</param>
    /// <param name="cancellationToken">The cancellation token.</param>
    public static async Task SetCoverPage(this HtmlConverter converter, HtmlFragment fragment,
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
        partObject.DocPartUnique ??= new() { Val = true };
        coverStdBlock.SdtContentBlock?.Remove();

        var elements = await converter.ParseAsync(fragment, cancellationToken);
        var content = new SdtContentBlock();
        content.Append(elements);
        content.Append(new Paragraph(
            new Run(
                new Break() { Type = BreakValues.Page })));

        coverStdBlock.SdtContentBlock = content;
    }

    /// <summary>
    /// Returns the table of contents, if one exists.
    /// <para>
    /// The TOC is represented using a Word content control associated with
    /// the built-in TOC gallery.
    /// </para>
    /// </summary>
    /// <returns>
    /// The TOC content control, or <see langword="null"/> when the document
    /// does not contain a table of contents.
    /// </returns>
    public static SdtBlock? GetTableOfContents(this HtmlConverter converter)
        => GetOrCreateDocumentBody(converter.mainPart).FindNamedSdtBlock(BuiltInGalleryTOCName);

    /// <summary>
    /// Create or updates the document table of contents.
    /// <para>
    /// The table of contents is generated from paragraphs using Word heading styles 
    /// (<c>Heading 1</c> through <c>Heading 9</c>, see also <see cref="HtmlConverter.HtmlStyles"/>)
    /// and represented as a Word content control associated with the built-in TOC gallery.
    /// </para>
    ///
    /// If a table of contents already exists, its content is replaced.
    /// </summary>
    /// <param name="converter">The converter owning the target document.</param>
    /// <param name="maxLevel">The maximum heading level to include in the table of contents. Valid values
    /// are typically between 1 and 9.</param>
    /// <param name="caption">The optional caption displayed above the table of contents.
    /// Specify <see langword="null"/> or an empty string to omit the caption.</param>
    /// <param name="position">Determines where the table of contents is inserted when it does not already exist.</param>
    /// <returns>The content control representing the table of contents.</returns>
    public static SdtBlock SetTableOfContents(
        this HtmlConverter converter, int maxLevel = 3,
        string caption = "Table Of Contents",
        DocumentPosition position = DocumentPosition.AfterCover)
    {
        var mainPart = converter.mainPart;
        var body = GetOrCreateDocumentBody(mainPart);
        SdtBlock? tocStdBlock = converter.GetTableOfContents() ?? new();
        body.InsertOrMove(tocStdBlock, position);


        tocStdBlock.SdtProperties ??= new();
        var partObject = tocStdBlock.SdtProperties.GetOrAddFirstChild<SdtContentDocPartObject>();
        partObject.DocPartGallery ??= new() { Val = BuiltInGalleryTOCName };
        partObject.DocPartUnique ??= new() { Val = true };
        tocStdBlock.SdtContentBlock?.Remove();

        // append main caption if provided
        var sdtContent = new SdtContentBlock();
        if (!string.IsNullOrEmpty(caption))
        {
            sdtContent.AppendChild(new Paragraph(
                new ParagraphProperties {
                    ParagraphStyleId = converter.HtmlStyles.GetParagraphStyle(converter.HtmlStyles.DefaultStyles.TOCHeading)
                },
                new Run(
                    new Text(caption))));
        }

        // compose list of headings. Look for styles.xml for any named heading.
        var headingStyleIds = converter.HtmlStyles.CollectHeadingStyles();

        maxLevel = Math.Max(1, Math.Min(9, maxLevel));
        var headings = body.Descendants<Paragraph>()
            .Where(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value != null &&
                headingStyleIds.TryGetValue(p.ParagraphProperties.ParagraphStyleId.Val.Value, out var level)
                && level <= maxLevel).ToList();

        var parsingContext = new ParsingContext(converter, mainPart, imageLoader: null!);
        foreach (var heading in headings)
        {
            int level = headingStyleIds[heading.ParagraphProperties!.ParagraphStyleId!.Val!.Value!];

            // create bookmark around the heading as the landing target of the TOC heading links
            string bookmarkId = EncloseHeadingInBookmark(parsingContext, heading);

            var hyperlink = new Hyperlink(
                new Run(
                    new RunProperties(new NoProof()) {
                        RunStyle = converter.HtmlStyles.GetRunStyle(converter.HtmlStyles.DefaultStyles.HyperlinkStyle)
                    },
                    new Text(heading.InnerText)
                ),
                new Run(
                    new RunProperties(
                        new NoProof(),
                        new WebHidden()
                    ),
                    new TabChar(),
                    // we can't reliabily guess the page number where the heading stands.
                    // we will leave the text empty but ready for a manual refresh from Word
                    new FieldChar { FieldCharType = FieldCharValues.Begin },
                    new FieldCode($@" PAGEREF {bookmarkId} \h ") {
                        Space = SpaceProcessingModeValues.Preserve },
                    new FieldChar { FieldCharType = FieldCharValues.Separate },
                    new Text("") { Space = SpaceProcessingModeValues.Preserve },
                    new FieldChar { FieldCharType = FieldCharValues.End }
                )
            ) {
                Anchor = bookmarkId, History = true };

            sdtContent.AppendChild(new Paragraph(
                new ParagraphProperties(
                    new Tabs(
                        new TabStop { Val = TabStopValues.Right, Leader = TabStopLeaderCharValues.Dot, Position = 9016 }
                    )
                ) {
                    ParagraphStyleId = converter.HtmlStyles.GetParagraphStyle($"TOC{level}")
                },
                hyperlink
            ));
        }

        var paragraphs = sdtContent.ChildElements.OfType<Paragraph>().Take(2);
        if (paragraphs.Count()  < 2)
            sdtContent.AppendChild(new Paragraph());

        Paragraph toc1 = paragraphs.ElementAt(1);

        // this fieldCode allows Word to auto-refresh the TOC
        toc1.InsertAfter(
            new Run(
                new FieldChar { FieldCharType = FieldCharValues.Begin },
                new FieldCode($@" TOC \o ""1-{maxLevel}"" \h \z \u ") {
                    Space = SpaceProcessingModeValues.Preserve },
                new FieldChar { FieldCharType = FieldCharValues.Separate }
            ),
            toc1.ParagraphProperties
        );
        sdtContent.AppendChild(new Paragraph(
            new Run(
                new FieldChar { FieldCharType = FieldCharValues.End })));

        tocStdBlock.SdtContentBlock = sdtContent;
        return tocStdBlock;
    }

    /// <summary>
    /// Ensures that the specified heading paragraph is enclosed within a bookmark.
    /// If the heading already contains a bookmark, the existing bookmark name is reused.
    /// Otherwise, a new bookmark is generated and inserted around the heading content.
    /// </summary>
    /// <returns>The bookmark name that can referenced by hyperlinks and PAGEREF fields.</returns>
    private static string EncloseHeadingInBookmark(ParsingContext parsingContext, Paragraph heading)
    {
        var bookmark = heading.GetFirstChild<BookmarkStart>();

        if (bookmark is not null)
        {
            if (bookmark.Name?.HasValue == false)
                bookmark.Name = parsingContext.IncrementBookmarkId().ToString(CultureInfo.InvariantCulture);
            return bookmark.Name!.Value!;
        }

        var bookmarkId = parsingContext.IncrementBookmarkId().ToString(CultureInfo.InvariantCulture);
        var bookmarkName = "_Toc" + bookmarkId;
        // need to be inserted after pPr to avoid schema warning
        heading.InsertAfter(new BookmarkStart() { Id = bookmarkId, Name = bookmarkName }, heading.GetFirstChild<ParagraphProperties>());
        heading.AppendChild(new BookmarkEnd() { Id = bookmarkId });

        return bookmarkName;
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

    static Paragraph CreatePageBreak() => new Paragraph(
        new Run(
            new Break { Type = BreakValues.Page }),
        new Run(
            new LastRenderedPageBreak())
    );
}
