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
using System.Collections.Concurrent;
using DocumentFormat.OpenXml.Packaging;

namespace HtmlToOpenXml.IO;

interface IImageLoader
{
    /// <summary>
    /// Download the remote or local image located at the specified url.
    /// </summary>
    Task<HtmlImageInfo?> Download(string imageUri, CancellationToken cancellationToken);
}

/// <summary>
/// Download and provison the metadata of a requested image.
/// </summary>
sealed class ImagePrefetcher<T> : ImagePrefetcherBase, IImageLoader
    where T: OpenXmlPartContainer, ISupportedRelationship<ImagePart>
{
    private readonly T hostingPart;
    private readonly IWebRequest resourceLoader;
    private readonly ConcurrentDictionary<string, HtmlImageInfo> prefetchedImages;
    private readonly object lockObject = new();
    private readonly ImageProcessingMode processingMode;


    /// <summary>
    /// Constructor.
    /// </summary>
    /// <param name="hostingPart">The image will be linked to that hosting part.
    /// Images are not shared between header, footer and body.</param>
    /// <param name="resourceLoader">Service to resolve an image.</param>
    /// <param name="processingMode">Specifies how images should be processed (embed, link, or data URI only).</param>
    public ImagePrefetcher(T hostingPart, IWebRequest resourceLoader, ImageProcessingMode processingMode = ImageProcessingMode.Embed)
    {
        this.hostingPart = hostingPart;
        this.resourceLoader = resourceLoader;
        this.processingMode = processingMode;
        this.prefetchedImages = [];
    }

    //____________________________________________________________________
    //
    // Public Functionality

    /// <summary>
    /// Download the remote or local image located at the specified url.
    /// </summary>
    public async Task<HtmlImageInfo?> Download(string imageUri, CancellationToken cancellationToken)
    {
        // Check if image is already cached using thread-safe operation
        if (prefetchedImages.TryGetValue(imageUri, out var iinfo))
            return iinfo;

        if (DataUri.IsWellFormed(imageUri)) // data inline, encoded in base64
        {
            iinfo = ReadDataUri(imageUri);
        }
        else
        {
            // Handle external images based on processing mode
            if (processingMode == ImageProcessingMode.EmbedDataUriOnly)
            {
                // Skip external images entirely
                return null;
            }
            else if (processingMode == ImageProcessingMode.LinkExternal)
            {
                // Create external link without downloading
                iinfo = CreateExternalImageLink(imageUri);
            }
            else
            {
                // Default: Download and embed
                iinfo = await DownloadRemoteImage(imageUri, cancellationToken).ConfigureAwait(false);
            }
        }

        // Add to cache using thread-safe operation
        if (iinfo != null)
        {
            // Double-check pattern to prevent duplicate adds during concurrent access
            prefetchedImages.TryAdd(imageUri, iinfo);
        }

        return iinfo;
    }

    /// <summary>
    /// Download the image and try to find its format type.
    /// </summary>
    private async Task<HtmlImageInfo?> DownloadRemoteImage(string src, CancellationToken cancellationToken)
    {
        Uri imageUri = new(src, UriKind.RelativeOrAbsolute);
        if (imageUri.IsAbsoluteUri && !resourceLoader.SupportsProtocol(imageUri.Scheme))
            return null;

        using var response = await resourceLoader.FetchAsync(imageUri, cancellationToken).ConfigureAwait(false);
        if (response?.Content == null || !response.Content.CanRead)
            return null;

        // For requested url with no filename, we need to read the media mime type if provided
        response.Headers.TryGetValue("Content-Type", out var mime);
        if (!TryInspectMimeType(mime, out PartTypeInfo type)
            && !TryGuessTypeFromUri(imageUri, out type)
            && !TryGuessTypeFromStream(response.Content, out type)
            )
        {
            return null;
        }

        return SaveImageAssert(type, response.Content.CopyTo);
    }

    /// <summary>
    /// Create an external relationship to an image without downloading it.
    /// </summary>
    private HtmlImageInfo? CreateExternalImageLink(string src)
    {
        Uri imageUri = new(src, UriKind.RelativeOrAbsolute);

        // Resolve relative URIs if possible (only for DefaultWebRequest which has BaseImageUrl)
        if (!imageUri.IsAbsoluteUri && resourceLoader is DefaultWebRequest defaultWebRequest
            && defaultWebRequest.BaseImageUrl != null)
        {
            string url1 = defaultWebRequest.BaseImageUrl.AbsoluteUri.TrimEnd('/', '\\');
            string path = src.TrimStart('/', '\\');
            imageUri = new Uri(string.Format("{0}/{1}", url1, path), UriKind.Absolute);
        }

        // Only create external links for absolute URIs with supported protocols
        if (!imageUri.IsAbsoluteUri || !resourceLoader.SupportsProtocol(imageUri.Scheme))
            return null;

        // Generate a unique GUID-based relationship ID for the external relationship
        string relationshipId = "imgext_" + Guid.NewGuid().ToString("N");

        // Create external relationship
        lock (lockObject)
        {
            hostingPart.AddExternalRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                imageUri,
                relationshipId);
        }

        // Return image info with external flag set
        // Note: Size will be empty as we don't download the image
        return new HtmlImageInfo(relationshipId) {
            IsExternal = true,
            Size = Size.Empty,
            TypeInfo = ImagePartType.Png // Default type, actual type doesn't matter for external links
        };
    }

    /// <summary>
    /// Parse the Data inline image.
    /// </summary>
    private HtmlImageInfo? ReadDataUri(string src)
    {
        if (DataUri.TryCreate(src, out var dataUri))
        {
            TryInspectMimeType(dataUri!.Mime, out PartTypeInfo type);

            return SaveImageAssert(type, stream => stream.Write(dataUri.Data, 0, dataUri.Data.Length));
        }

        return null;
    }

    private HtmlImageInfo SaveImageAssert(PartTypeInfo type, Action<Stream> writeImage)
    {
        ImagePart ipart;
        string relationshipId = "img_" + Guid.NewGuid().ToString("N");
        lock (lockObject)
        {
            ipart = hostingPart.AddImagePart(type, relationshipId);
        }

        Size originalSize;
        using (var outputStream = ipart.GetStream(FileMode.Create))
        {
            writeImage(outputStream);
            outputStream.Seek(0L, SeekOrigin.Begin);
            originalSize = GetImageSize(outputStream);
        }

        string partId = hostingPart.GetIdOfPart(ipart);
        return new HtmlImageInfo(partId)
        {
            TypeInfo = type,
            Size = originalSize
        };
    }
}
