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

namespace HtmlToOpenXml
{
    /// <summary>
    /// Shared the BackChannel to connect to the download external resources.
    /// </summary>
    static class BackChannels
    {
#if !FEATURE_NETHTTP
        // HttpClient contains pooling and is optimized to be used thread-safe and reentrant.
        // The best practices is to make it static and to not dispose it

        /// <summary>
        /// Gets the shared Http client.
        /// </summary>
        internal static System.Net.Http.HttpClient HttpClient { get; } = new System.Net.Http.HttpClient();
#endif

        /// <summary>
        /// Process the download of a Http resource.
        /// </summary>
        /// <param name="requestUri">The remote endpoint to retrieve.</param>
        /// <param name="proxy">The configuration <see cref="HtmlConverter.WebProxy"/> for this http request.</param>
        public static HttpResponse CreateWebRequest(Uri requestUri, WebProxy proxy)
        {
            var httpResponse = new HttpResponse();

            try
            {
#if !FEATURE_NETHTTP
                var requestMessage = new System.Net.Http.HttpRequestMessage();
                requestMessage.RequestUri = requestUri;
                var response = HttpClient.SendAsync(requestMessage).Result;
                httpResponse.Body = response.Content.ReadAsByteArrayAsync().Result;

                if (requestUri.Scheme.StartsWith("http"))
                    httpResponse.ContentType = response.Content.Headers.ContentType?.ToString();
#else
                System.Net.WebClient webClient = new WebClientEx(proxy);
                httpResponse.Body = webClient.DownloadData(requestUri);

                // For requested url with no filename, we need to read the media mime type if provided
                if (requestUri.Scheme.StartsWith("http"))
                    httpResponse.ContentType = webClient.ResponseHeaders[System.Net.HttpResponseHeader.ContentType];
#endif
            }
            catch (Exception exc)
            {
                if (Logging.On) Logging.PrintError("ImageDownloader.DownloadData(\"" + requestUri.AbsoluteUri + "\")", exc);
                return null;
            }

            return httpResponse;
        }
    }
}