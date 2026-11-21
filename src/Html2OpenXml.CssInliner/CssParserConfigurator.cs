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

namespace HtmlToOpenXml.CssInliner;

/// <summary>
/// Extend the parser configuration <c>&lt;style&gt;</c> sheets
/// into native HTML element attributes before document processing.
/// </summary>
/// <param name="mode">Controls how external style sheets are handled during conversion.</param>
/// <param name="webRequester">Controls how external style sheets are handled during conversion.</param>
public class CssParserExtension(CssProcessingMode mode, IO.IWebRequest webRequester) : IHtmlParserConfigurator
{
    /// <inheritdoc />
    public IConfiguration Configure(IConfiguration configuration)
    {
        // Enables downloading the external stylesheets
        if (mode == CssProcessingMode.All)
        {
            return configuration
                .With(new WebRequestRequester(webRequester))
                .WithDefaultLoader(new AngleSharp.Io.LoaderOptions {
                IsResourceLoadingEnabled = true
            })
            .WithCss();
        }

        return configuration.WithCss();
    }
}
