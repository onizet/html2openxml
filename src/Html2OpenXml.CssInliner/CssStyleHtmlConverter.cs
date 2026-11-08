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
using AngleSharp.Css.Dom;
using AngleSharp.Css.Parser;
using AngleSharp.Dom;
using DocumentFormat.OpenXml.Packaging;

namespace HtmlToOpenXml.CssInliner;

/// <summary>
/// An extended HTML-to-OpenXML converter that automatically compiles and inlines 
/// global <c>&lt;style&gt;</c> sheets into native HTML element attributes before document processing.
/// </summary>
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
        // Enables full CSS parsing and computation mechanisms
        var configuration = Configuration.Default.WithCss();
        if (CssProcessing == CssProcessingMode.All)
        {
            configuration.WithDefaultLoader();
        }

        return configuration;
    }

    /// <inheritdoc />
    protected override async Task OnDocumentLoadedAsync(IDocument document, IBrowsingContext browsingContext,
        CancellationToken cancellationToken)
    {
        InlineStyles(document);
    }

    /// <summary>
    /// Compiles internal CSS rules found in &lt;style&gt; tags into native inline attributes 
    /// on the matched HTML elements within the document. This process effectively merges 
    /// embedded stylesheets into the structural content during conversion.
    /// </summary>
    protected void InlineStyles(IDocument document)
    {
        // gather all <style> tags using AngleSharp.Css
        var styleElements = document.QuerySelectorAll("style");
        var cssParser = new CssParser();

        foreach (var styleEl in styleElements)
        {
            var stylesheet = cssParser.ParseStyleSheet(styleEl.TextContent);

            // map selectors down into raw inline style strings
            foreach (var rule in stylesheet.Rules.OfType<ICssStyleRule>())
            {
                var matchingElements = document.QuerySelectorAll(rule.SelectorText);
                foreach (var element in matchingElements)
                {
                    // Convert rule.Style into native inline attributes
                    AppendStylesToElement(element, rule.Style, cssParser);
                }
            }

            // Remove the style tag so it isn't parsed again
            styleEl.Remove();
        }
    }

    private static void AppendStylesToElement(IElement element, ICssStyleDeclaration newStyles, CssParser parser)
    {
        string? currentStyleText = element.GetAttribute("style");
        if (currentStyleText is null)
            return;

        var currentStyles = parser.ParseDeclaration(currentStyleText);
        if (currentStyles is null)
            return;

        foreach (var property in newStyles)
        {
            var priority = newStyles.GetPropertyPriority(property.Name);
            bool isPriority = priority.Equals("important", StringComparison.OrdinalIgnoreCase);
            if (string.IsNullOrEmpty(currentStyles.GetPropertyValue(property.Name)) || isPriority)
            {
                currentStyles.SetProperty(property.Name, property.Value, isPriority? priority : null);
            }
        }

        element.SetAttribute("style", currentStyles.ToCss());
    }
}
