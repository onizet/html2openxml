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
using System.Collections.Generic;
using System.IO;
using System.Net;

namespace HtmlToOpenXml.IO
{
    /// <summary>
    /// Specifies what is stored when receiving data.
    /// </summary>
    public class Resource : IDisposable
    {
        /// <summary>
        /// Gets the status code that has been send with the response.
        /// </summary>
        public HttpStatusCode StatusCode { get; set; }

        /// <summary>
        /// Gets the headers that have been send with the response.
        /// </summary>
        public IDictionary<string, string> Headers { get; private set; } = new Dictionary<string, string>();

        /// <summary>
        /// Gets the content that has been send with the response.
        /// </summary>
        public Stream Content { get; set; } = Stream.Null;

        void IDisposable.Dispose()
        {
            Content?.Dispose();
            Headers.Clear();
        }
    }
}