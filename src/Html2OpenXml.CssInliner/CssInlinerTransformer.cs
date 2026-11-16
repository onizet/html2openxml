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
