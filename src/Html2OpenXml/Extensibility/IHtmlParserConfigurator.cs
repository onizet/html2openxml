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

namespace HtmlToOpenXml;

/// <summary>
/// Configures the HTML parsing environment before a document is loaded.
///
/// <para>
/// Implementations can register additional parser services such as CSS engines, resource loaders,
/// scripting support, or custom AngleSharp integrations.
/// </para>
/// </summary>
public interface IHtmlParserConfigurator
{
    /// <summary>
    /// Configure the HTML parsing environment.
    /// </summary>
    /// <param name="configuration">The current HTML parser configuration.</param>
    /// <returns>Returns the augmented configuration.</returns>
    IConfiguration Configure(IConfiguration configuration);
}