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
using AngleSharp;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml.IO;

namespace HtmlToOpenXml;

/// <summary>
/// Primary entry point for converting HTML into OpenXml Word document.
/// 
/// The convert can be used with newly created documents as well as existing document templates.
/// When a template is used, styles, themes, numbering definitions, bookmarks, and other Word settings
/// are automatically reused.
/// <para>
/// Typical usage:
/// <list type="number">
/// <item>Create an HtmlConverter from a MainDocumentPart obtained from a WordProcessingDocument.</item>
/// <item>Call ParseBody() to append the converted content to the document body.</item>
/// </list>
/// </para>
/// 
/// Advanced scenarios can use ParseAsync() to control where the generated OpenXml elements are inserted.
/// </summary>
/// <example>
/// Basic quick start:
/// 
/// <code>
/// await using var generatedDocument = new MemoryStream();
/// using var package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document);
/// var mainPart = package.AddMainDocumentPart();
/// new Document(new Body()).Save(mainPart);
/// HtmlConverter converter = new(mainPart);
/// await converter.ParseBody(html);
/// </code>
/// </example>
public partial class HtmlConverter
{
    internal readonly MainDocumentPart mainPart;
    // Cache all the ImagePart processed to avoid downloading the same image
    private IImageLoader? headerImageLoader, bodyImageLoader, footerImageLoader;
    private readonly WordDocumentStyle htmlStyles;
    private readonly IWebRequest webRequester;


    /// <summary>
    /// Create a converter bound to a Word document.
    /// 
    /// <para>
    /// Reuse the same HtmlConverter instance for the lifetime of a document
    /// to avoid reloading cached configuration such as styles and bookmarks.
    /// </para>
    /// Do not use the same converter instance with multiple documents.
    /// </summary>
    /// <param name="mainPart">The mainDocumentPart must be the document MainDocumentPart where converted content will be inserted.</param>
    /// <param name="webRequester">Control retrieval of external resources such as images.</param>
    public HtmlConverter(MainDocumentPart mainPart, IWebRequest? webRequester = null)
    {
        this.mainPart = mainPart ?? throw new ArgumentNullException(nameof(mainPart));
        this.htmlStyles = new WordDocumentStyle(mainPart);
        this.webRequester = webRequester ?? new DefaultWebRequest();
    }

    /// <summary>
    /// Convert HTML into OpenXml elements and return them to the caller.
    /// 
    /// <para>
    /// Use this method when your HTML is simple and don't need to download any external
    /// resources.
    /// </para>
    /// This method exist for backward compatibility reason.
    /// Prefer ParseAsync() or for most scenarios, use ParseBody().
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <returns>Returns a collection of generated OpenXml elements.</returns>
    public IList<OpenXmlCompositeElement> Parse(string html)
    {
        bodyImageLoader ??= new ImagePrefetcher<MainDocumentPart>(mainPart, webRequester, ImageProcessing);
        return ParseCoreAsync(html, mainPart, bodyImageLoader,
            new ParallelOptions() { CancellationToken = CancellationToken.None })
            .ConfigureAwait(false).GetAwaiter().GetResult().ToList();
    }

    /// <summary>
    /// Start the asynchronous parse processing where the output is intended to be inserted in <see cref="MainDocumentPart"/>.
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <param name="cancellationToken">The cancellation token.</param>
    /// <returns>Returns a collection of generated OpenXml elements.</returns>
    [Obsolete("Use ParseAsync instead to respect naming convention")]
    [NuSpec.AI.AiIgnore]
    [System.Diagnostics.CodeAnalysis.ExcludeFromCodeCoverage]
    public Task<IEnumerable<OpenXmlCompositeElement>> Parse(string html, CancellationToken cancellationToken = default)
    {
        return ParseAsync(html, cancellationToken);
    }

    /// <summary>
    /// Convert HTML into OpenXml elements and return them to the caller.
    /// 
    /// <para>
    /// Use this method when you need full control over the insertion point or need
    /// to inspect or modify the generated elements before adding them to the document.
    /// </para>
    /// For most scenarios, prefer ParseBody().
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <param name="cancellationToken">The cancellation token.</param>
    /// <returns>Returns a collection of generated OpenXml elements.</returns>
    public Task<IEnumerable<OpenXmlCompositeElement>> ParseAsync(string html, CancellationToken cancellationToken = default)
    {
        return ParseAsync(html, new ParallelOptions { CancellationToken = cancellationToken });
    }

    /// <summary>
    /// Convert HTML into OpenXml elements and return them to the caller.
    /// 
    /// <para>
    /// Use this method when you need full control over the insertion point or need
    /// to inspect or modify the generated elements before adding them to the document.
    /// </para>
    /// For most scenarios, prefer ParseBody().
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <param name="parallelOptions">Control the parallelism when downloading the remote resources such as images.</param>
    /// <returns>Returns a collection of generated OpenXml elements.</returns>
    public Task<IEnumerable<OpenXmlCompositeElement>> ParseAsync(string html, ParallelOptions parallelOptions)
    {
        bodyImageLoader ??= new ImagePrefetcher<MainDocumentPart>(mainPart, webRequester, ImageProcessing);

        return ParseCoreAsync(html, mainPart, bodyImageLoader, parallelOptions);
    }

    /// <summary>
    /// Parse some HTML and append the genereated content to a Word header.
    /// 
    /// <para>
    /// Typical uses include company logos, document titles, confidentiality notices,
    /// report identifiers, and other recurring content displayed at the top of each page.
    /// </para>
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <param name="headerType">Determines the page(s) on which the current header shall be displayed.
    /// If omitted, the value <see cref="HeaderFooterValues.Default"/> is used.</param>
    /// <param name="cancellationToken">The cancellation token.</param>
    /// <seealso cref="HeaderPart"/>
    public async Task ParseHeader(string html, HeaderFooterValues? headerType = null,
        CancellationToken cancellationToken = default)
    {
        headerType ??= HeaderFooterValues.Default;
        var headerPart = ResolveHeaderFooterPart<HeaderReference, HeaderPart>(headerType);

        headerPart.Header ??= new();
        headerImageLoader ??= new ImagePrefetcher<HeaderPart>(headerPart, webRequester, ImageProcessing);

        var paragraphs = await ParseCoreAsync(html, headerPart, headerImageLoader,
            new ParallelOptions() { CancellationToken = cancellationToken },
            htmlStyles.GetParagraphStyle(htmlStyles.DefaultStyles.HeaderStyle))
            .ConfigureAwait(false);

        headerPart.Header.Append(paragraphs);
    }

    /// <summary>
    /// Parse some HTML and append the generated content into a Word footer.
    ///
    /// <para>
    /// Typical uses include legal disclaimers, contact information, page numbering, copyright notices,
    /// and other recurring content displayed at the bottom of each page.
    /// </para>
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <param name="footerType">Determines the page(s) on which the current footer shall be displayed.
    /// If omitted, the value <see cref="HeaderFooterValues.Default"/> is used.</param>
    /// <param name="cancellationToken">The cancellation token.</param>
    /// <seealso cref="FooterPart"/>
    public async Task ParseFooter(string html, HeaderFooterValues? footerType = null,
        CancellationToken cancellationToken = default)
    {
        footerType ??= HeaderFooterValues.Default;
        var footerPart = ResolveHeaderFooterPart<FooterReference, FooterPart>(footerType);

        footerPart.Footer ??= new();
        footerImageLoader ??= new ImagePrefetcher<FooterPart>(footerPart, webRequester, ImageProcessing);

        var paragraphs = await ParseCoreAsync(html, footerPart, footerImageLoader,
            new ParallelOptions() { CancellationToken = cancellationToken },
            htmlStyles.GetParagraphStyle(htmlStyles.DefaultStyles.FooterStyle))
            .ConfigureAwait(false);

        footerPart.Footer.Append(paragraphs);
    }

    /// <summary>
    /// Parse asynchronously the Html and append the output into the Body of the document.
    /// This is the recommended method for most scenarios.
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <param name="cancellationToken">The cancellation token.</param>
    /// <seealso cref="MainDocumentPart"/>
    public async Task ParseBody(string html, CancellationToken cancellationToken = default)
    {
        bodyImageLoader ??= new ImagePrefetcher<MainDocumentPart>(mainPart, webRequester, ImageProcessing);
        var paragraphs = await ParseCoreAsync(html, mainPart, bodyImageLoader,
            new ParallelOptions() { CancellationToken = cancellationToken },
            htmlStyles.GetParagraphStyle(htmlStyles.DefaultStyles.Paragraph))
            .ConfigureAwait(false);

        if (!paragraphs.Any())
            return;

        Body body = mainPart.Document!.Body!;
        SectionProperties? sectionProperties = body.GetLastChild<SectionProperties>();
        foreach (var para in paragraphs)
            body.Append(para);

        // we automatically create the _top bookmark if missing. To avoid having an empty paragrah,
        // let's try to merge with its next paragraph.
        var p = body.GetFirstChild<Paragraph>();
        if (p != null && p.GetFirstChild<BookmarkStart>()?.Name == "_top"
            && !p.HasChild<Run>()
            && p.NextSibling() is Paragraph nextPara)
        {
            nextPara.PrependChild(p.GetFirstChild<BookmarkEnd>()?.CloneNode(false));
            nextPara.PrependChild(p.GetFirstChild<BookmarkStart>()!.CloneNode(false));
            p.Remove();
        }

        // Push the sectionProperties as the last element of the Body
        // (required by OpenXml schema to avoid the bad formatting of the document)
        if (sectionProperties != null)
        {
            sectionProperties.Remove();
            body.AddChild(sectionProperties);
        }
    }

    /// <summary>
    /// Start the asynchronous parse processing. Use this overload if you want to control the downloading of images.
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <param name="parallelOptions">The configuration of parallelism while downloading the remote resources.</param>
    /// <returns>Returns a list of parsed paragraph.</returns>
    [Obsolete("Use ParseAsync instead to respect naming convention")]
    [NuSpec.AI.AiIgnore]
    [System.Diagnostics.CodeAnalysis.ExcludeFromCodeCoverage]
    public Task<IEnumerable<OpenXmlCompositeElement>> Parse(string html, ParallelOptions parallelOptions)
    {
        bodyImageLoader ??= new ImagePrefetcher<MainDocumentPart>(mainPart, webRequester, ImageProcessing);

        return ParseCoreAsync(html, mainPart, bodyImageLoader, parallelOptions);
    }

    /// <summary>
    /// Start the asynchronous parse processing and append the output into the Body of the document.
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <param name="cancellationToken">The cancellation token.</param>
    [Obsolete("Use ParseBody instead for output clarification")]
    [NuSpec.AI.AiIgnore]
    [System.Diagnostics.CodeAnalysis.ExcludeFromCodeCoverage]
    public Task ParseHtml(string html, CancellationToken cancellationToken = default)
    {
        return ParseBody(html, cancellationToken);
    }

    /// <summary>
    /// Reloads the style cache from the current Word document (<see cref="WordDocumentStyle"/>).
    /// Call this method if styles are added after the HtmlConverter instance has been created.
    /// </summary>
    /// <remarks>You don't need to call this method if you register the missing style 
    /// from <see cref="WordDocumentStyle.StyleMissing"/> event.</remarks>
    public void RefreshStyles()
    {
        htmlStyles.PrepareStyles(mainPart);
    }

    /// <summary>
    /// Core method to start the asynchronous parse processing.
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <param name="hostingPart">The OpenXml container where the content will be inserted into.</param>
    /// <param name="imageLoader">The image resolver service linked to the <paramref name="hostingPart"/>.</param>
    /// <param name="parallelOptions">The configuration of parallelism while downloading the remote resources.</param>
    /// <param name="defaultParagraphStyleId">The default OpenXml style to apply on paragraphs.</param> 
    /// <returns>Returns a list of parsed paragraph.</returns>
    private async Task<IEnumerable<OpenXmlCompositeElement>> ParseCoreAsync(string html,
        OpenXmlPartContainer hostingPart, IImageLoader imageLoader,
        ParallelOptions parallelOptions,
        ParagraphStyleId? defaultParagraphStyleId = null)
    {
        if (string.IsNullOrWhiteSpace(html))
            return [];

        var browsingContext = BrowsingContext.New();
        var htmlDocument = await browsingContext.OpenAsync(req => req.Content(html), parallelOptions.CancellationToken).ConfigureAwait(false);
        if (htmlDocument == null)
            return [];

        if (mainPart.Document == null)
            new Document(new Body()).Save(mainPart);
        else if (mainPart.Document.Body == null)
            mainPart.Document.Body = new Body();

        await PreloadImages(htmlDocument, imageLoader, parallelOptions).ConfigureAwait(false);

        Expressions.HtmlDomExpression expression;
        if (hostingPart is MainDocumentPart)
            expression = new Expressions.BodyExpression(htmlDocument.Body!, defaultParagraphStyleId);
        else
            expression = new Expressions.BlockElementExpression(htmlDocument.Body!, defaultParagraphStyleId);

        var parsingContext = new ParsingContext(this, hostingPart, imageLoader);
        var paragraphs = expression.Interpret(parsingContext);
        return paragraphs.Cast<OpenXmlCompositeElement>();
    }

    /// <summary>
    /// Walk through all the <c>img</c> tags and preload all the remote images.
    /// We save the image chunks into the underlying WordProcessingDocument to keep memory low.
    /// </summary>
    private static async Task PreloadImages(AngleSharp.Dom.IDocument htmlDocument,
        IImageLoader imageLoader, ParallelOptions parallelOptions)
    {
        var imageUris = htmlDocument.QuerySelectorAll("img[src]")
            .Cast<IHtmlImageElement>()
            .Where(e => AngleSharpExtensions.TryParseUrl(e.GetAttribute("src"), UriKind.RelativeOrAbsolute, out var _))
            .Select(e => e.GetAttribute("src")!)
            .Distinct();
        if (!imageUris.Any())
            return;

        await imageUris.ForEachAsync(
            async (img, cts) => await imageLoader.Download(img, cts),
            parallelOptions).ConfigureAwait(false);
    }

    /// <summary>
    /// Create or resolve the header/footer related to the type.
    /// </summary>
    private TPart ResolveHeaderFooterPart<TRefType, TPart>(HeaderFooterValues? type)
        where TPart: OpenXmlPart, IFixedContentTypePart
        where TRefType: HeaderFooterReferenceType, new()
    {
        bool wasRefSet = false;
        TPart? part = null;

        var sectionProps = mainPart.Document!.Body!.Elements<SectionProperties>();
        if (!sectionProps.Any())
        {
            sectionProps = [new SectionProperties()];
            mainPart.Document.Body!.AddChild(sectionProps.First());
        }
        else
        {
            var reference = sectionProps.SelectMany(sectPr => sectPr.Elements<TRefType>())
                .Where(r => r.Id?.HasValue == true)
                .FirstOrDefault(r => r.Type?.Value == type);

            if (reference != null)
                part = (TPart) mainPart.GetPartById(reference.Id!);
            wasRefSet = part is not null;
        }

        part ??= mainPart.AddNewPart<TPart>();

        if (!wasRefSet)
        {
            sectionProps.First().PrependChild(new TRefType() {
                Id = mainPart.GetIdOfPart(part),
                Type = type
            });
        }

        return part;
    }

    //____________________________________________________________________
    //
    // Configuration

    /// <summary>
    /// Defines the location where to add the acronym or abbreviation explanation tag.
    /// Defaults to the end of the page (<see cref="AcronymPosition.PageEnd"/>).
    /// </summary>
    public AcronymPosition AcronymPosition { get; set; }

    /// <summary>
    /// Defines whether internal anchor hyperlinks are converted.
    /// 
    /// <para>
    /// Anchor links are hyperlinks targeting a location within the document, such as
    /// <c>#_top</c> or a bookmark reference (<c>#bookmarkReference</c>).
    /// </para>
    /// Anchor links are enabled by default. If the target cannot be resolved in the current Word
    /// document, the content will be rendered as a simple text.
    /// </summary>
    /// <remarks>
    /// It exists some predefined anchors used by Word such as <c>#_top</c> to refer to the top of the document.
    /// This built-in anchor is always accepted regardless this property value.
    /// For others anchors like referring to your own bookmark or a title, add a 
    /// <see cref="DocumentFormat.OpenXml.Wordprocessing.BookmarkStart"/> and 
    /// <see cref="DocumentFormat.OpenXml.Wordprocessing.BookmarkEnd"/> elements
    /// and set the value of href to <c>#your_bookmark</c>.
    /// </remarks>
    public bool SupportsAnchorLinks { get; set; } = true;

    /// <summary>
    /// Defines whether anchor links are included or not in the conversion.
    /// </summary>
    /// <remarks>An anchor is a term used to define a hyperlink destination inside a document.
    /// <see href="http://www.w3schools.com/HTML/html_links.asp"/>.
    /// <br/>
    /// It exists some predefined anchors used by Word such as _top to refer to the top of the document.
    /// The anchor <i>#_top</i> is always accepted regardless this property value.
    /// For others anchors like refering to your own bookmark or a title, add a 
    /// <see cref="DocumentFormat.OpenXml.Wordprocessing.BookmarkStart"/> and 
    /// <see cref="DocumentFormat.OpenXml.Wordprocessing.BookmarkEnd"/> elements
    /// and set the value of href to <i><c>#name of your bookmark</c></i>.
    /// </remarks>
    [Obsolete("Use SupportsAnchorLink instead, if ExcludeLinkAnchor = true -> SupportsAnchorLink = false")]
    [NuSpec.AI.AiIgnore]
    public bool ExcludeLinkAnchor { get => !SupportsAnchorLinks; set => SupportsAnchorLinks = !value; }

    /// <summary>
    /// Gets the style manager for the current conversion.
    /// 
    /// <para>
    /// HtmlStyles controls how HTML elements are translated into Word styles.
    /// Through this object you can customize the default style mappings, add custom styles,
    /// and react when a referenced style is missing from the document.
    /// </para>
    /// </summary>
    public WordDocumentStyle HtmlStyles
    {
        get { return htmlStyles; }
    }

    /// <summary>
    /// Defines where the Legend tag (<c>caption</c>) should be rendered above or below the table.
    /// Defaults above the table.
    /// </summary>
    public CaptionPositionValues TableCaptionPosition { get; set; }

    /// <summary>
    /// Defines whether the preformatted blocks (<c>pre</c>) are rendered as tables.
    /// 
    /// <para>
    /// When enabled, <c>pre</c> elements are converted to a single-cell table to preserve whitespace,
    /// indentation and line breaks.
    /// This is particularly useful for source code, console output, and technical blog posts.
    /// </para>
    /// When disabled, preformatted content is rendered using regular Word paragraphs.
    /// Defaults to <see langword="false"/>.
    ///  </summary>
    public bool RenderPreAsTable { get; set; }

    /// <summary>
    /// Controls how images are handled during conversion.
    /// 
    /// <para>
    /// Images are embedded in the generated document by default.
    /// External resources are resolved through <see cref="IWebRequest"/>, which can be
    /// customised to provide authentication, support additional formats such as WebP (see the wiki),
    /// or resolve relative URLs.
    /// </para>
    /// Default: <see cref="ImageProcessingMode.Embed"/>.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Use <see cref="ImageProcessingMode.Embed"/> (default) to download and embed all images,
    /// creating self-contained documents but potentially large file sizes.
    /// </para>
    /// <para>
    /// Use <see cref="ImageProcessingMode.LinkExternal"/> to link to external images via relationships,
    /// keeping document size small but requiring internet access to view images.
    /// Data URI images (base64 encoded) are still embedded.
    /// </para>
    /// <para>
    /// Use <see cref="ImageProcessingMode.EmbedDataUriOnly"/> to only embed data URI images
    /// and skip external images entirely.
    /// </para>
    /// </remarks>
    public ImageProcessingMode ImageProcessing { get; set; } = ImageProcessingMode.Embed;

    /// <summary>
    /// Defines whether numbering is preserved across multiple ordered lists (<c>ol</c>).
    /// 
    /// <para>
    /// By default, a subsequent <c>ol</c> continues numbering from the previous one.
    /// Use the HTML <c>start</c> attribute to reset or override the next number. When this property
    /// is <see langword="false"/>, each <c>ol</c> starts at 1 unless specified otherwise.
    /// </para>
    /// Defaults to true.
    /// </summary>
    public bool ContinueNumbering { get; set; } = true;

    /// <summary>
    /// Defines whether (<c>h1-h6</c>) elements are rendered using the corresponding Word heading style.
    /// 
    /// <para>
    /// Missing heading styles are added automatically when required.
    /// When heading text beings with a numbering pattern such as <c>"1.", "1.1.", or "1 "</c> or the list
    /// use the CSS class <c>`decimal-tiered`</c>, the converter
    /// interprets it as a numbered heading. Any associated heading numbering is then managed by Word
    /// through the Heading styles.
    /// </para>
    /// This feature is enabled by default.
    /// </summary>
    public bool SupportsHeadingNumbering { get; set; } = true;
}
