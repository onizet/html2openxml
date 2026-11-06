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