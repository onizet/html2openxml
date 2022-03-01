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
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace HtmlToOpenXml.IO
{
    /// <summary>
    /// Default implementation of the <see cref="IWebRequest"/>.
    /// Supports http, https, local file and inline data.
    /// </summary>
    public class DefaultWebRequest : IWebRequest
    {
        private static HashSet<string> SupportedProtocols = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            "http", "https", "file"
        };
        private Uri baseImageUri;
        private static readonly HttpClient DefaultHttp = new HttpClient(new HttpClientHandler() {
            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
        });
        private readonly HttpClient httpClient;


        /// <summary>
        /// Initialize a new instance of the <see cref="DefaultWebRequest"/> class.
        /// </summary>
        public DefaultWebRequest() : this(DefaultHttp) { }

        /// <summary>
        /// Initialize a new instance of the <see cref="DefaultWebRequest"/> class with
        /// the specified <see cref="HttpClient"/>.
        /// </summary>
        /// <param name="httpClient">The HTTP client to use to download remote resources.</param>
        public DefaultWebRequest(HttpClient httpClient)
        {
            this.httpClient = httpClient ?? DefaultHttp;
            this.httpClient.DefaultRequestHeaders.AcceptEncoding.ParseAdd("gzip, deflate");
        }

        /// <inheritdoc/>
        public virtual void Dispose()
        {
        }

        /// <inheritdoc/>
        public virtual Task<Resource> FetchAsync(Uri requestUri, CancellationToken cancellationToken)
        {
            if (!requestUri.IsAbsoluteUri && BaseImageUrl != null)
            {
                requestUri = new Uri(BaseImageUrl, requestUri);
            }

            if (requestUri.IsFile)
            {
                return DownloadLocalFile(requestUri, cancellationToken);
            }

            return DownloadHttpFile(requestUri, cancellationToken);
        }

        /// <summary>
        /// Process to the read of a file from the File System.
        /// </summary>
        protected virtual Task<Resource> DownloadLocalFile(Uri requestUri, CancellationToken cancellationToken)
        {
            // replace string %20 in LocalPath by daviderapicavoli (patch #15938)
            string localPath = Uri.UnescapeDataString(requestUri.LocalPath);

            try
            {
                return Task.FromResult(new Resource() {
                    Content = System.IO.File.OpenRead(localPath),
                    StatusCode = HttpStatusCode.OK
                });
            }
            catch (Exception exc)
            {
                if (exc is System.IO.IOException || exc is UnauthorizedAccessException || exc is System.Security.SecurityException || exc is NotSupportedException)
                    return null;
                throw;
            }
        }

        /// <summary>
        /// Process to the download of a resource with Http/Https protocol.
        /// </summary>
        protected virtual async Task<Resource> DownloadHttpFile(Uri requestUri, CancellationToken cancellationToken)
        {
            var resource = new Resource();

            try
            {
                var response = await httpClient.GetAsync(requestUri, cancellationToken).ConfigureAwait(false);
                if (response == null) return null;
                resource.StatusCode = response.StatusCode;

                if (response.IsSuccessStatusCode)
                    resource.Content = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);

                foreach (var header in response.Headers)
                    resource.Headers.Add(header.Key, string.Join(", ", header.Value));
            }
            catch (TaskCanceledException)
            {
                if (cancellationToken.IsCancellationRequested)
                    return null;
                throw;
            }

            return resource;
        }

        /// <inheritdoc/>
        public virtual bool SupportsProtocol(string protocol)
            => SupportedProtocols.Contains(protocol);

        /// <summary>
        /// Gets or sets the base Uri used to automaticaly resolve relative images 
        /// if used with ImageProcessing = AutomaticDownload.
        /// </summary>
        public Uri BaseImageUrl
        {
            get { return this.baseImageUri; }
            set
            {
                if (value != null)
                {
                    if (!value.IsAbsoluteUri)
                        throw new ArgumentException("BaseImageUrl should be an absolute Uri");
                    // in case of local uri (file:///) we need to be sure the uri ends with '/' or the
                    // combination of uri = new Uri(@"C:\users\demo\images", "pic.jpg");
                    // will eat the images part
                    if (value.IsFile && value.LocalPath[value.LocalPath.Length - 1] != '/')
                        value = new Uri(value.OriginalString + '/');
                }
                this.baseImageUri = value;
            }
        }
    }
}