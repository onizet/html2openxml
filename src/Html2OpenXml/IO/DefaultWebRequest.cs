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
using System.Net;
#if !NET5_0_OR_GREATER
using System.Net.Http;
#endif
using Microsoft.Extensions.Logging;

namespace HtmlToOpenXml.IO;

/// <summary>
/// Default implementation of <see cref="IWebRequest"/>.
/// Supports http, https, local file and inline data (base64).
///
/// Derive from this class to customise resource retrieval, authenatication, image processing,
/// URL resolution, or content transformation while reusing the built-in behaviour.
/// </summary>
public class DefaultWebRequest : IWebRequest
{
    private static readonly HashSet<string> SupportedProtocols = new(StringComparer.OrdinalIgnoreCase) {
        "http", "https", "file"
    };
    private Uri? baseImageUri;
    private static readonly HttpClient DefaultHttp = new(new HttpClientHandler() {
        AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
    });
    private readonly HttpClient httpClient;
    private readonly ILogger? logger;



    /// <summary>
    /// Initialize a new instance of the <see cref="DefaultWebRequest"/> class.
    /// </summary>
    public DefaultWebRequest(ILogger? logger = null) : this(DefaultHttp, logger) { }

    /// <summary>
    /// Initialize a new instance of the <see cref="DefaultWebRequest"/> class with
    /// the specified <see cref="HttpClient"/>.
    /// <para>
    /// Supply your own <see cref="HttpClient"/> to customise how external resources are retrieved,
    /// for example to configure cookies, authentication headers, API keys, proxies, or message handlers.
    /// </para>
    /// </summary>
    /// <param name="httpClient">The HTTP client to use to download remote resources.</param>
    /// <param name="logger">Provide a logging mechanism for diagnose.</param>
    public DefaultWebRequest(HttpClient httpClient, ILogger? logger = null)
    {
        this.httpClient = httpClient ?? DefaultHttp;
        this.httpClient.DefaultRequestHeaders.AcceptEncoding.ParseAdd("gzip, deflate");
        this.logger = logger;
    }

    /// <inheritdoc/>
    public virtual Task<Resource?> FetchAsync(Uri requestUri, CancellationToken cancellationToken)
    {
        if (!requestUri.IsAbsoluteUri && BaseImageUrl != null)
        {
            requestUri = UrlCombine(BaseImageUrl, requestUri.OriginalString);
        }

        bool isLocalFile;
        try
        {
            isLocalFile = requestUri.IsFile;
        }
        catch (InvalidOperationException)
        {
            isLocalFile = false;
        }

        if (isLocalFile)
        {
            return DownloadLocalFile(requestUri, cancellationToken);
        }

        return DownloadHttpFile(requestUri, cancellationToken);
    }

    /// <summary>
    /// Process to the read of a file from the File System.
    /// </summary>
    protected virtual Task<Resource?> DownloadLocalFile(Uri requestUri, CancellationToken cancellationToken)
    {
        // replace string %20 in LocalPath by daviderapicavoli (patch #15938)
        string localPath = Uri.UnescapeDataString(requestUri.LocalPath);

        try
        {
            if (logger?.IsEnabled(LogLevel.Debug) == true)
            {
                logger.LogDebug("Downloading local file: {RequestUri}", requestUri);
            }

            return Task.FromResult<Resource?>(new Resource() {
                Content = File.OpenRead(localPath),
                StatusCode = HttpStatusCode.OK
            });
        }
        catch (Exception exc)
        {
            logger?.LogError(exc, "Failed to download file: {RequestUri}", requestUri);

            if (exc is IOException || exc is UnauthorizedAccessException || exc is System.Security.SecurityException || exc is NotSupportedException)
                return Task.FromResult<Resource?>(null);
            throw;
        }
    }

    /// <summary>
    /// Process to the download of a resource with Http/Https protocol.
    /// </summary>
    protected virtual async Task<Resource?> DownloadHttpFile(Uri requestUri, CancellationToken cancellationToken)
    {
        var resource = new Resource();

        try
        {
            if (logger?.IsEnabled(LogLevel.Debug) == true)
            {
                logger.LogDebug("Downloading remote file: {RequestUri}", requestUri);
            }

            if (httpClient.BaseAddress is null && !requestUri.IsAbsoluteUri)
                return null;

            var response = await httpClient.GetAsync(requestUri, cancellationToken).ConfigureAwait(false);
            if (response == null) return null;
            resource.StatusCode = response.StatusCode;

            if (response.IsSuccessStatusCode)
            {
                resource.Content = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
                if (response.Content.Headers.TryGetValues("Content-Type", out var mime))
                {
                    resource.Headers.Add("Content-Type", string.Join(", ", mime));
                }
            }

            foreach (var header in response.Headers)
                resource.Headers.Add(header.Key, string.Join(", ", header.Value));
        }
        catch (TaskCanceledException)
        {
            if (cancellationToken.IsCancellationRequested)
                return null;
            throw;
        }
        catch(Exception exc)
        {
            logger?.LogError(exc, "Failed to download file: {RequestUri}", requestUri);
            throw;
        }

        return resource;
    }

    /// <inheritdoc/>
    public virtual bool SupportsProtocol(string protocol)
        => SupportedProtocols.Contains(protocol);

    /// <summary>
    /// Combine two URIs.
    /// </summary>
    /// <param name="baseUrl">The absolute base uri</param>
    /// <param name="path">The relative uri</param>
    private static Uri UrlCombine(Uri baseUrl, string path)
    {
        /* `new Uri(Uri baseUri, string relativeUri)` is counter intuitive.
         Uri baseUri = new Uri("https://www.example.com/api");
         Uri result = new Uri(baseUri, "v1/helloworld");
         ---> https://www.example.com/v1/helloworld (missing the /api)
        */
        string url1 = baseUrl.AbsoluteUri.TrimEnd('/', '\\');
        path = path.TrimStart('/', '\\');

        return new Uri(string.Format("{0}/{1}", url1, path), UriKind.Absolute);
    }

    /// <summary>
    /// Defines the base URI used to resolve image URLS.
    /// <para>
    /// Use this property when HTML content contains relative images references such as
    /// <c>/images/logo.png</c> or <c>/../assets/banner.jpg</c>.
    /// During conversion, relative URLs are combined with this base URI to locate
    /// and download the image resource
    /// </para>
    /// </summary>
    public Uri? BaseImageUrl
    {
        get { return baseImageUri; }
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
            baseImageUri = value;
        }
    }
}
