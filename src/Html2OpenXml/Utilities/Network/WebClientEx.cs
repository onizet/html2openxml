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
using System.Net;

namespace HtmlToOpenXml
{
#if FEATURE_NETHTTP
    /// <summary>
    /// Provides some utilies methods for translating Http attributes to OpenXml elements.
    /// </summary>
    sealed class WebClientEx : System.Net.WebClient
    {
        private readonly WebProxy proxy;


        public WebClientEx(WebProxy proxy)
        {
            this.proxy = proxy;

            if (proxy != null)
            {
                if (proxy.Credentials != null)
                    this.Credentials = proxy.Credentials;
                if (proxy.Proxy != null)
                    this.Proxy = proxy.Proxy;
            }
        }

        protected override WebRequest GetWebRequest(Uri address)
        {
            WebRequest request = base.GetWebRequest(address);

            HttpWebRequest httpRequest = request as HttpWebRequest;
            if (httpRequest != null && proxy != null)
            {
                httpRequest.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                httpRequest.CookieContainer = proxy.Cookies;
                httpRequest.Headers = proxy.HttpRequestHeaders;
            }

            return request;
        }

        protected override WebResponse GetWebResponse(WebRequest request, IAsyncResult result)
        {
            WebResponse response = base.GetWebResponse(request, result);
            if (proxy != null) ReadCookies(response);
            return response;
        }

        protected override WebResponse GetWebResponse(WebRequest request)
        {
            WebResponse response = base.GetWebResponse(request);
            if (proxy != null) ReadCookies(response);
            return response;
        }

        private void ReadCookies(WebResponse r)
        {
            var response = r as HttpWebResponse;
            if (response != null)
            {
                CookieCollection cookies = response.Cookies;
                proxy.Cookies.Add(cookies);
            }
        }
    }
#endif
}