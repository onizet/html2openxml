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

namespace HtmlToOpenXml.IO;

/// <summary>
/// Handle retrieval of external resources referenced during HTML conversion.
///
/// Resources may include images, stylesheets, scripts and local files.
/// <para>
/// Implement this interface to customize resource resolution, provide authentication,
/// support additional formats, rewrite URLs, or transform downlaoded content before it is
/// inserted into the document.
/// </para>
/// <para>
/// Implementations are expected to support the http, https and file protocols.
/// </para>
/// </summary>
public interface IWebRequest
{
    /// <summary>
    /// Retrieves a resource referenced by the HTML document.
    ///
    /// Returns <see langword="null"/> when the resource cannot be retrieved.
    /// The returned <see cref="Resource"/> must provide a readable content stream.
    /// Returned <see cref="Resource"/> instances are disposed by HtmlToOpenXml after processing.
    /// </summary>
    /// <param name="requestUri">The Uri the request is sent to.</param>
    /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive
    /// notice of cancellation.</param>
    /// <returns>The task that will eventually give the resource's response data.</returns>
    Task<Resource?> FetchAsync(Uri requestUri, CancellationToken cancellationToken);

    /// <summary>
    /// Checks if the given protocol is supported.
    /// </summary>
    /// <param name="protocol">The protocol to check for, e.g. http.</param>
    bool SupportsProtocol(string protocol);
}