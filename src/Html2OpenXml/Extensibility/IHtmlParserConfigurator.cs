/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
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