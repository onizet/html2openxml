/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
using AngleSharp;
using AngleSharp.Css.Dom;
using AngleSharp.Css.Parser;
using AngleSharp.Dom;

namespace HtmlToOpenXml.CssInliner;

/// <summary>
/// Core engine to compiles and inlines global stylesheets and <c>&lt;style&gt;</c> sheets
/// into native HTML element attributes before document processing.
/// </summary>
public static class CssStyleInliner
{
    /// <summary>
    /// Compiles internal CSS rules found in &lt;style&gt; tags into native inline attributes
    /// on the matched HTML elements within the document. This process effectively merges
    /// embedded stylesheets into the structural content during conversion.
    /// </summary>
    public static void InlineStyles(IDocument document, CssProcessingMode processingMode)
    {
        var cssParser = new CssParser();

        if (processingMode == CssProcessingMode.All)
        {
            foreach (var stylesheet in document.StyleSheets.OfType<ICssStyleSheet>())
            {
                // map selectors down into raw inline style strings
                ApplyCssRules(document, cssParser, stylesheet.Rules);
            }
        }

        // gather all <style> tags using AngleSharp.Css
        foreach (var styleEl in document.QuerySelectorAll("style"))
        {
            var stylesheet = cssParser.ParseStyleSheet(styleEl.TextContent);

            ApplyCssRules(document, cssParser, stylesheet.Rules);

            // Remove the style tag so it isn't parsed again
            styleEl.Remove();
        }
    }

    private static void ApplyCssRules(IDocument document, CssParser cssParser, ICssRuleList rules)
    {
        // map selectors down into raw inline style strings
        foreach (var rule in rules.OfType<ICssStyleRule>())
        {
            var matchingElements = document.QuerySelectorAll(rule.SelectorText);
            foreach (var element in matchingElements)
            {
                // Convert rule.Style into native inline attributes
                AppendStylesToElement(element, rule.Style, cssParser);
            }
        }
    }

    private static void AppendStylesToElement(IElement element, ICssStyleDeclaration newStyles, CssParser parser)
    {
        string? currentStyleText = element.GetAttribute("style") ?? "";
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
