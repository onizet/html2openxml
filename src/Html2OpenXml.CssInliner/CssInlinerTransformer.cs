/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
using AngleSharp.Dom;

namespace HtmlToOpenXml.CssInliner;

/// <summary>
/// Transformer pipeline which automatically compiles and inlines global <c>&lt;style&gt;</c> sheets
/// into native HTML element attributes before document processing.
/// </summary>
/// <param name="mode">Controls how external style sheets are handled during conversion.</param>
public class CssInlinerTransformer(CssProcessingMode mode) : IHtmlTransformer
{
    /// <inheritdoc />
    public Task TransformAsync(IDocument document, CancellationToken cancellationToken)
    {
        CssStyleInliner.InlineStyles(document, mode);
        return Task.CompletedTask;
    }
}
