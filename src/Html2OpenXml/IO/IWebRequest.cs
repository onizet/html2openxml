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
using System;
using System.Threading;
using System.Threading.Tasks;

namespace HtmlToOpenXml.IO
{
    /// <summary>
    /// Interface used to handle resource requests for a document. These
    /// requests include, but are not limited to, media, script and styling
    /// resources.
    /// The expected protocols to support are: http, https and file.
    /// </summary>
    public interface IWebRequest : IDisposable
    {
        /// <summary>
        /// Performs an asynchronous request that can be cancelled.
        /// </summary>
        /// <param name="requestUri">The Uri the request is sent to.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive
        /// notice of cancellation.</param>
        /// <returns>The task that will eventually give the resource's response data.</returns>
        Task<Resource> FetchAsync(Uri requestUri, CancellationToken cancellationToken);

        /// <summary>
        /// Checks if the given protocol is supported.
        /// </summary>
        /// <param name="protocol">The protocol to check for, e.g. http.</param>
        bool SupportsProtocol(string protocol);
    }
}