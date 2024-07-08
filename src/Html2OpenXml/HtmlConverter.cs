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
    /// <summary>Cache all the ImagePart processed to avoid downloading the same image.</summary>
    private ImagePrefetcher? imagePrefetcher;
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
    /// Start the parse processing.
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <returns>Returns a list of parsed paragraph.</returns>
    public IList<OpenXmlCompositeElement> Parse(string html)
    {
        return Parse(html, CancellationToken.None).ConfigureAwait(false).GetAwaiter().GetResult().ToList();
    }

    /// <summary>
    /// Start the parse processing.
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <param name="cancellationToken">The cancellation token.</param>
    /// <returns>Returns a list of parsed paragraph.</returns>
    public Task<IEnumerable<OpenXmlCompositeElement>> Parse(string html, CancellationToken cancellationToken = default)
    {
        return Parse(html, new ParallelOptions() { CancellationToken = cancellationToken });
    }

    /// <summary>
    /// Start the parse processing. Use this overload if you want to control the downloading of images.
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <param name="parallelOptions">The configuration of parallelism while downloading the remote resources.</param>
    /// <returns>Returns a list of parsed paragraph.</returns>
    public async Task<IEnumerable<OpenXmlCompositeElement>> Parse(string html, ParallelOptions parallelOptions)
    {
        if (string.IsNullOrWhiteSpace(html))
            return [];

        // ensure a body exists to avoid any errors when trying to access it
        if (mainPart.Document == null)
            new Document(new Body()).Save(mainPart);
        else if (mainPart.Document.Body == null)
            mainPart.Document.Body = new Body();

        var browsingContext = BrowsingContext.New();
        var htmlDocument = await browsingContext.OpenAsync(req => req.Content(html), parallelOptions.CancellationToken);
        if (htmlDocument == null)
            return [];

        await PreloadImages(htmlDocument, parallelOptions).ConfigureAwait(false);

        var parsingContext = new ParsingContext(this, mainPart);
        var body = new Expressions.BodyExpression (htmlDocument.Body!);
        var paragraphs = body.Interpret (parsingContext);
        return paragraphs.Cast<OpenXmlCompositeElement>();
    }

    /// <summary>
    /// Start the parse processing and append the converted paragraphs into the Body of the document.
    /// </summary>
    /// <param name="html">The HTML content to parse</param>
    /// <param name="cancellationToken">The cancellation token.</param>
    public async Task ParseHtml(string html, CancellationToken cancellationToken = default)
    {
        // This method exists because we may ensure the SectionProperties remains the last element of the body.
        // It's mandatory when dealing with page orientation

        var paragraphs = await Parse(html, cancellationToken);

        Body body = mainPart.Document.Body!;
        SectionProperties? sectionProperties = body.GetLastChild<SectionProperties>();
        foreach (var para in paragraphs)
            body.Append(para);

        // move the paragraph with BookmarkStart `_GoBack` as the last child
        var p = body.GetFirstChild<Paragraph>();
        if (p != null && p.HasChild<BookmarkStart>())
        {
            p.Remove();
            body.Append(p);
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
    /// Refresh the cache of styles presents in the document.
    /// </summary>
    public void RefreshStyles()
    {
        htmlStyles.PrepareStyles(mainPart);
    }

    /// <summary>
    /// Walk through all the <c>img</c> tags and preload all the remote images.
    /// </summary>
    private async Task PreloadImages(AngleSharp.Dom.IDocument htmlDocument, ParallelOptions parallelOptions)
    {
        var imageUris = htmlDocument.QuerySelectorAll("img[src]")
            .Cast<AngleSharp.Html.Dom.IHtmlImageElement>()
            .Where(e => AngleSharpExtensions.TryParseUrl(e.GetAttribute("src"), UriKind.RelativeOrAbsolute, out var _))
            .Select(e => e.GetAttribute("src")!);
        if (!imageUris.Any())
            return;

        await imageUris.ForEachAsync(
            async (img, cts) => await ImagePrefetcher.Download(img, cts),
            parallelOptions).ConfigureAwait(false);
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
    /// Resolve a remote or inline image resource.
    /// </summary>
    internal ImagePrefetcher ImagePrefetcher
    {
        get => imagePrefetcher ??= new ImagePrefetcher(mainPart, webRequester);
    }
}
