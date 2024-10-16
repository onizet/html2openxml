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
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using AngleSharp;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml.IO;

namespace HtmlToOpenXml;

/// <summary>
/// Helper class to convert some Html text to OpenXml elements.
/// </summary>
public partial class HtmlConverter
{
    private readonly MainDocumentPart mainPart;
    // Cache all the ImagePart processed to avoid downloading the same image
    private IImageLoader? headerImageLoader, bodyImageLoader, footerImageLoader;
    private readonly WordDocumentStyle htmlStyles;
    private readonly IWebRequest webRequester;


    /// <summary>
    /// Constructor.
    /// </summary>
    /// <param name="mainPart">The mainDocumentPart of a document where to write the conversion to.</param>
    /// <remarks>We preload some configuration from inside the document such as style, bookmarks,...</remarks>
    public HtmlConverter(MainDocumentPart mainPart) : this(mainPart, null)
    {
    }

    /// <summary>
    /// Constructor.
    /// </summary>
    /// <param name="mainPart">The mainDocumentPart of a document where to write the conversion to.</param>
    /// <param name="webRequester">Factory to download the images.</param>
    /// <remarks>We preload some configuration from inside the document such as style, bookmarks,...</remarks>
    public HtmlConverter(MainDocumentPart mainPart, IWebRequest? webRequester = null)
    {
        this.mainPart = mainPart ?? throw new ArgumentNullException(nameof(mainPart));
        this.htmlStyles = new WordDocumentStyle(mainPart);
        this.webRequester = webRequester ?? new DefaultWebRequest();
    }

    /// <summary>
    /// Parse some HTML content where the output is intented to be inserted in <see cref="MainDocumentPart"/>.
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <returns>Returns a list of parsed paragraph.</returns>
    public IList<OpenXmlCompositeElement> Parse(string html)
    {
        bodyImageLoader ??= new ImagePrefetcher<MainDocumentPart>(mainPart, webRequester);
        return ParseCoreAsync(html, mainPart, bodyImageLoader,
            new ParallelOptions() { CancellationToken = CancellationToken.None })
            .ConfigureAwait(false).GetAwaiter().GetResult().ToList();
    }

    /// <summary>
    /// Start the asynchroneous parse processing where the output is intented to be inserted in <see cref="MainDocumentPart"/>.
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <param name="cancellationToken">The cancellation token.</param>
    /// <returns>Returns a list of parsed paragraph.</returns>
    [Obsolete("Use ParseAsync instead to respect naming convention")]
    [System.Diagnostics.CodeAnalysis.ExcludeFromCodeCoverage]
    public Task<IEnumerable<OpenXmlCompositeElement>> Parse(string html, CancellationToken cancellationToken = default)
    {
        return ParseAsync(html, cancellationToken);
    }

    /// <summary>
    /// Start the asynchroneous parse processing where the output is intented to be inserted in <see cref="MainDocumentPart"/>.
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <param name="cancellationToken">The cancellation token.</param>
    /// <returns>Returns a list of parsed paragraph.</returns>
    public Task<IEnumerable<OpenXmlCompositeElement>> ParseAsync(string html, CancellationToken cancellationToken = default)
    {
        return ParseAsync(html, new ParallelOptions { CancellationToken = cancellationToken });
    }

    /// <summary>
    /// Start the asynchroneous parse processing where the output is intented to be inserted in <see cref="MainDocumentPart"/>.
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <param name="parallelOptions">The configuration of parallelism while downloading the remote resources.</param>
    /// <returns>Returns a list of parsed paragraph.</returns>
    public Task<IEnumerable<OpenXmlCompositeElement>> ParseAsync(string html, ParallelOptions parallelOptions)
    {
        bodyImageLoader ??= new ImagePrefetcher<MainDocumentPart>(mainPart, webRequester);

        return ParseCoreAsync(html, mainPart, bodyImageLoader, parallelOptions);
    }

    /// <summary>
    /// Parse asynchroneously the Html and append the output into the Header of the document.
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
        headerImageLoader ??= new ImagePrefetcher<HeaderPart>(headerPart, webRequester);

        var paragraphs = await ParseCoreAsync(html, headerPart, headerImageLoader,
            new ParallelOptions() { CancellationToken = cancellationToken },
            htmlStyles.GetParagraphStyle(htmlStyles.DefaultStyles.HeaderStyle));

        headerPart.Header.Append(paragraphs);
    }

    /// <summary>
    /// Parse asynchroneously the Html and append the output into the Footer of the document.
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
        footerImageLoader ??= new ImagePrefetcher<FooterPart>(footerPart, webRequester);

        var paragraphs = await ParseCoreAsync(html, footerPart, footerImageLoader,
            new ParallelOptions() { CancellationToken = cancellationToken },
            htmlStyles.GetParagraphStyle(htmlStyles.DefaultStyles.FooterStyle));

        footerPart.Footer.Append(paragraphs);
    }

    /// <summary>
    /// Parse asynchroneously the Html and append the output into the Body of the document.
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <param name="cancellationToken">The cancellation token.</param>
    /// <seealso cref="MainDocumentPart"/>
    public async Task ParseBody(string html, CancellationToken cancellationToken = default)
    {
        bodyImageLoader ??= new ImagePrefetcher<MainDocumentPart>(mainPart, webRequester);
        var paragraphs = await ParseCoreAsync(html, mainPart, bodyImageLoader,
            new ParallelOptions() { CancellationToken = cancellationToken },
            htmlStyles.GetParagraphStyle(htmlStyles.DefaultStyles.Paragraph));

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
    /// Start the asynchroneous parse processing. Use this overload if you want to control the downloading of images.
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <param name="parallelOptions">The configuration of parallelism while downloading the remote resources.</param>
    /// <returns>Returns a list of parsed paragraph.</returns>
    [Obsolete("Use ParseAsync instead to respect naming convention")]
    [System.Diagnostics.CodeAnalysis.ExcludeFromCodeCoverage]
    public Task<IEnumerable<OpenXmlCompositeElement>> Parse(string html, ParallelOptions parallelOptions)
    {
        bodyImageLoader ??= new ImagePrefetcher<MainDocumentPart>(mainPart, webRequester);

        return ParseCoreAsync(html, mainPart, bodyImageLoader, parallelOptions);
    }

    /// <summary>
    /// Start the asynchroneous parse processing and append the output into the Body of the document.
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <param name="cancellationToken">The cancellation token.</param>
    [Obsolete("Use ParseBody instead for output clarification")]
    [System.Diagnostics.CodeAnalysis.ExcludeFromCodeCoverage]
    public Task ParseHtml(string html, CancellationToken cancellationToken = default)
    {
        return ParseBody(html, cancellationToken);
    }

    /// <summary>
    /// Refresh the cache of styles presents in the document.
    /// </summary>
    public void RefreshStyles()
    {
        htmlStyles.PrepareStyles(mainPart);
    }

    /// <summary>
    /// Start the asynchroneous parse processing. Use this overload if you want to control the downloading of images.
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
    /// </summary>
    private async Task PreloadImages(AngleSharp.Dom.IDocument htmlDocument,
        IImageLoader imageLoader, ParallelOptions parallelOptions)
    {
        var imageUris = htmlDocument.QuerySelectorAll("img[src]")
            .Cast<AngleSharp.Html.Dom.IHtmlImageElement>()
            .Where(e => AngleSharpExtensions.TryParseUrl(e.GetAttribute("src"), UriKind.RelativeOrAbsolute, out var _))
            .Select(e => e.GetAttribute("src")!);
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

        var sectionProps = mainPart.Document.Body!.Elements<SectionProperties>();
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
    /// Gets or sets where to render the acronym or abbreviation tag.
    /// </summary>
    public AcronymPosition AcronymPosition { get; set; }

    /// <summary>
    /// Gets or sets whether anchor links are included or not in the convertion.
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
    public bool ExcludeLinkAnchor { get; set; }

    /// <summary>
    /// Gets the Html styles manager mapping to OpenXml style properties.
    /// </summary>
    public WordDocumentStyle HtmlStyles
    {
        get { return htmlStyles; }
    }

    /// <summary>
    /// Gets or sets where the Legend tag (<c>caption</c>) should be rendered (above or below the table).
    /// </summary>
    public CaptionPositionValues TableCaptionPosition { get; set; }

    /// <summary>
    /// Gets or sets whether the <c>pre</c> tag should be rendered as a table (default <see langword="false"/>).
    /// </summary>
    /// <remarks>The table will contains only one cell.</remarks>
    public bool RenderPreAsTable { get; set; }

    /// <summary>
    /// Defines whether ordered lists (<c>ol</c>) continue incrementing existing numbering
    /// or restarts to 1 (defaults continues numbering).
    /// </summary>
    public bool ContinueNumbering { get; set; } = true;

    /// <summary>
    /// Gets the mainDocumentPart of the destination OpenXml document.
    /// </summary>
    internal MainDocumentPart MainPart
    {
        get => mainPart;
    }
}
