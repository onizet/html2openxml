/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
using HtmlToOpenXml.CssInliner;

namespace HtmlToOpenXml;

/// <summary>
/// Provides extension methods that enable CSS stylesheet processing
/// for <see cref="HtmlConverter"/> instances.
/// </summary>
public static class CssHtmlConverterExtensions
{
    /// <summary>
    /// Enables CSS stylesheet support for the converter.
    /// </summary>
    /// <param name="converter">The converter to configure.</param>
    /// <param name="mode">Defines how internal and external stylesheets are processed.</param>
    /// <returns>The configured converter.</returns>
    public static HtmlConverter UseCssInliner(this HtmlConverter converter,
        CssProcessingMode mode = CssProcessingMode.All)
    {
        converter.ParserConfigurators.Add(
            new CssParserExtension(mode, converter.WebRequester));

        converter.Transformers.Add(
            new CssInlinerTransformer(mode));

        return converter;
    }
}
