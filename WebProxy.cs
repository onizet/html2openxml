/* Copyright (C) Olivier Nizet http://html2openxml.codeplex.com - All Rights Reserved
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
using System.Collections.Specialized;
using System.Net;

namespace NotesFor.HtmlToOpenXml
{
    /// <summary>
    /// Represents the configuration used to download some data such as the images.
    /// </summary>
    public sealed class WebProxy
    {
        private CookieContainer cookies;
        private WebHeaderCollection requestHeaders;


        /// <summary>
        /// Gets or sets the credentials to submit to the proxy server for authentication.
        /// </summary>
        public ICredentials Credentials { get; set; }

        /// <summary>
        /// Gets or sets the proxy access.
        /// </summary>
        public IWebProxy Proxy { get; set; }

        /// <summary>
        /// Gets or sets the collection of authorization ticket cookie (mainly for Claims authentication).
        /// </summary>
        public CookieContainer Cookies
        {
            get { return cookies ?? (cookies = new CookieContainer()); }
        }

        /// <summary>
        /// Gets or sets the Http headers that will be sent when requesting an image.
        /// </summary>
        public WebHeaderCollection HttpRequestHeaders
        {
            get { return requestHeaders ?? (requestHeaders = new WebHeaderCollection()); }
        }
    }
}