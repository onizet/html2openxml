/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
namespace HtmlToOpenXml.CssInliner;

/// <summary>
/// Controls whether external stylesheets are ignored or downloaded and processed.
/// </summary>
public enum CssProcessingMode
{
    /// <summary>
    /// Process internal <c>style</c> tags embedded within the HTML document.
    /// No external web requests are initiated for linked resources.
    /// </summary>
    InternalOnly,

    /// <summary>
    /// Process internal style tags and automatically download external stylesheets
    /// referenced via <c>&lt;link rel="stylesheet" href="..."&gt;</c> tags.
    /// </summary>
    All
}