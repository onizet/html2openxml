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
