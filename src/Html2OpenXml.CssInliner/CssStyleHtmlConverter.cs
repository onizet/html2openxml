/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
using AngleSharp;
using AngleSharp.Dom;
using DocumentFormat.OpenXml.Packaging;

namespace HtmlToOpenXml.CssInliner;

/// <summary>
/// An extended HTML-to-OpenXML converter that automatically compiles and inlines 
/// global <c>&lt;style&gt;</c> sheets into native HTML element attributes before document processing.
/// </summary>
[Obsolete("Prefer converter.UseCssInliner()", true)]
public class CssStyleHtmlConverter : HtmlConverter
{
    /// <summary>
    /// Create a converter bound to a Word document, with global CSS style sheet support enabled.
    ///
    /// <para>
    /// Reuse the same HtmlConverter instance for the lifetime of a document
    /// to avoid reloading cached configuration such as styles and bookmarks.
    /// </para>
    /// Do not use the same converter instance with multiple documents.
    /// </summary>
    /// <param name="mainPart">The mainDocumentPart must be the document MainDocumentPart where converted content will be inserted.</param>
    /// <param name="webRequester">Control retrieval of external resources such as images.</param>
    public CssStyleHtmlConverter(MainDocumentPart mainPart, IO.IWebRequest? webRequester = null)
        : base (mainPart, webRequester)
    {
    }

    /// <summary>
    /// Create a converter bound to a Word document, with global CSS style sheet support enabled.
    ///
    /// <para>
    /// Reuse the same HtmlConverter instance for the lifetime of a document
    /// to avoid reloading cached configuration such as styles and bookmarks.
    /// </para>
    /// Do not use the same converter instance with multiple documents.
    /// </summary>
    /// <param name="package">The top-level class of a Word document where converted content will be inserted.</param>
    /// <param name="webRequester">Control retrieval of external resources such as images.</param>
    public CssStyleHtmlConverter(WordprocessingDocument package, IO.IWebRequest? webRequester = null)
        : base (package, webRequester)
    {
    }

    /// <summary>
    /// Controls how external style sheets are handled during conversion.
    ///
    /// <para>
    /// Internal style tags are evaluated and inlined into elements by default.
    /// External resources can be resolved through web loaders, which can be configured
    /// to support relative URLs or skipped entirely to optimize processing speeds.
    /// </para>
    /// Default: <see cref="CssProcessingMode.All"/>.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Use <see cref="CssProcessingMode.InternalOnly"/> to only process internal style tags,
    /// avoiding external network calls and keeping processing fast and safe.
    /// </para>
    /// <para>
    /// Use <see cref="CssProcessingMode.All"/> to automatically download, parse, and
    /// merge external stylesheets referenced by link elements into the structural flow,
    /// as well as internal style tags.
    /// </para>
    /// </remarks>
    public CssProcessingMode CssProcessing { get; set; } = CssProcessingMode.All;


    /// <inheritdoc />
    protected override IConfiguration GetParserConfiguration()
    {
        // Enables downloading the external stylesheets
        if (CssProcessing == CssProcessingMode.All)
        {
            return Configuration.Default
                .With(new WebRequestRequester(WebRequester))
                .WithDefaultLoader(new AngleSharp.Io.LoaderOptions {
                IsResourceLoadingEnabled = true
            })
            .WithCss();
        }

        return Configuration.Default.WithCss();
    }

    /// <inheritdoc />
    protected override async Task OnDocumentLoadedAsync(IDocument document, IBrowsingContext browsingContext,
        CancellationToken cancellationToken)
    {
        CssStyleInliner.InlineStyles(document, CssProcessing);
    }
}
