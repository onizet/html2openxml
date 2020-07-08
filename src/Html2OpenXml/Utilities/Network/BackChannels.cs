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
        // HttpClient contains pooling and is optimized to be used thread-safe and reentrant.
        // The best practices is to make it static and to not dispose it

        /// <summary>
        /// Gets the shared Http client.
        /// </summary>
        internal static System.Net.Http.HttpClient HttpClient { get; } = new System.Net.Http.HttpClient();

        /// <summary>
        /// Process the download of a Http resource.
        /// </summary>
        /// <param name="requestUri">The remote endpoint to retrieve.</param>
        public static HttpResponse CreateWebRequest(Uri requestUri)
        {
            var httpResponse = new HttpResponse();

            try
            {
                var requestMessage = new System.Net.Http.HttpRequestMessage();
                requestMessage.RequestUri = requestUri;
                var response = HttpClient.SendAsync(requestMessage).Result;
                httpResponse.Body = response.Content.ReadAsByteArrayAsync().Result;

                if (requestUri.Scheme.StartsWith("http"))
                    httpResponse.ContentType = response.Content.Headers.ContentType?.ToString();
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