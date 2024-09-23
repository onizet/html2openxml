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
using System.Threading;
using System.Threading.Tasks;
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
sealed class ImagePrefetcher<T> : IImageLoader
    where T: OpenXmlPartContainer, ISupportedRelationship<ImagePart>
{
    // Map extension to PartTypeInfo
    private static readonly Dictionary<string, PartTypeInfo> knownExtensions = new(StringComparer.OrdinalIgnoreCase) {
        { ".gif", ImagePartType.Gif },
        { ".bmp", ImagePartType.Bmp },
        { ".emf", ImagePartType.Emf },
        { ".ico", ImagePartType.Icon },
        { ".jp2", ImagePartType.Jp2 },
        { ".jpeg", ImagePartType.Jpeg },
        { ".jpg", ImagePartType.Jpeg },
        { ".jpe", ImagePartType.Jpeg },
        { ".pcx", ImagePartType.Pcx },
        { ".png", ImagePartType.Png },
        { ".svg", ImagePartType.Svg },
        { ".tif", ImagePartType.Tif },
        { ".tiff", ImagePartType.Tiff },
        { ".wmf", ImagePartType.Wmf }
    };
    private readonly T hostingPart;
    private readonly IWebRequest resourceLoader;
    private readonly HtmlImageInfoCollection prefetchedImages;


    /// <summary>
    /// Constructor.
    /// </summary>
    /// <param name="hostingPart">The image will be linked to that hosting part.
    /// Images are not shared between header, footer and body.</param>
    /// <param name="resourceLoader">Service to resolve an image.</param>
    public ImagePrefetcher(T hostingPart, IWebRequest resourceLoader)
    {
        this.hostingPart = hostingPart;
        this.resourceLoader = resourceLoader;
        this.prefetchedImages = new HtmlImageInfoCollection();
    }

    //____________________________________________________________________
    //
    // Public Functionality

    /// <summary>
    /// Download the remote or local image located at the specified url.
    /// </summary>
    public async Task<HtmlImageInfo?> Download(string imageUri, CancellationToken cancellationToken)
    {
        if (prefetchedImages.Contains(imageUri))
            return prefetchedImages[imageUri];

        HtmlImageInfo? iinfo;
        if (DataUri.IsWellFormed(imageUri)) // data inline, encoded in base64
        {
            iinfo = ReadDataUri(imageUri);
        }
        else
        {
            iinfo = await DownloadRemoteImage(imageUri, cancellationToken);
        }

        if (iinfo != null)
            prefetchedImages.Add(iinfo);

        return iinfo;
    }

    /// <summary>
    /// Download the image and try to find its format type.
    /// </summary>
    private async Task<HtmlImageInfo?> DownloadRemoteImage(string src, CancellationToken cancellationToken)
    {
        Uri imageUri = new Uri(src, UriKind.RelativeOrAbsolute);
        if (imageUri.IsAbsoluteUri && !resourceLoader.SupportsProtocol(imageUri.Scheme))
            return null;

        Resource? response;

        response = await resourceLoader.FetchAsync(imageUri, cancellationToken).ConfigureAwait(false);
        if (response?.Content == null)
            return null;

        using (response)
        {
            // For requested url with no filename, we need to read the media mime type if provided
            response.Headers.TryGetValue("Content-Type", out var mime);
            if (!TryInspectMimeType(mime, out PartTypeInfo type)
                && !TryGuessTypeFromUri(imageUri, out type)
                && !TryGuessTypeFromStream(response.Content, out type))
            {
                return null;
            }

            var ipart = hostingPart.AddImagePart(type);
            Size originalSize;
            using (var outputStream = ipart.GetStream(FileMode.Create))
            {
                response.Content.CopyTo(outputStream);

                outputStream.Seek(0L, SeekOrigin.Begin);
                originalSize = GetImageSize(outputStream);
            }

            return new HtmlImageInfo(src, hostingPart.GetIdOfPart(ipart)) {
                TypeInfo = type,
                Size = originalSize
            };
        }
    }

    /// <summary>
    /// Parse the Data inline image.
    /// </summary>
    private HtmlImageInfo? ReadDataUri(string src)
    {
        if (DataUri.TryCreate(src, out var dataUri))
        {
            Size originalSize;
            knownContentType.TryGetValue(dataUri!.Mime, out PartTypeInfo type);
            var ipart = hostingPart.AddImagePart(type);
            using (var outputStream = ipart.GetStream(FileMode.Create))
            {
                outputStream.Write(dataUri.Data, 0, dataUri.Data.Length);

                outputStream.Seek(0L, SeekOrigin.Begin);
                originalSize = GetImageSize(outputStream);
            }

            return new HtmlImageInfo(src, hostingPart.GetIdOfPart(ipart)) {
                TypeInfo = type,
                Size = originalSize
            };
        }

        return null;
    }

    //____________________________________________________________________
    //
    // Private Implementation

    // http://stackoverflow.com/questions/58510/using-net-how-can-you-find-the-mime-type-of-a-file-based-on-the-file-signature
    private static readonly Dictionary<string, PartTypeInfo> knownContentType = new(StringComparer.OrdinalIgnoreCase) {
        { "image/gif", ImagePartType.Gif },
        { "image/pjpeg", ImagePartType.Jpeg },
        { "image/jp2", ImagePartType.Jp2 },
        { "image/jpg", ImagePartType.Jpeg },
        { "image/jpeg", ImagePartType.Jpeg },
        { "image/x-png", ImagePartType.Png },
        { "image/png", ImagePartType.Png },
        { "image/tiff", ImagePartType.Tiff },
        { "image/vnd.microsoft.icon", ImagePartType.Icon },
        // these icons mime type are wrong but we should nevertheless take care (http://en.wikipedia.org/wiki/ICO_%28file_format%29#MIME_type)
        { "image/x-icon", ImagePartType.Icon },
        { "image/icon", ImagePartType.Icon },
        { "image/ico", ImagePartType.Icon },
        { "text/ico", ImagePartType.Icon },
        { "text/application-ico", ImagePartType.Icon },
        { "image/bmp", ImagePartType.Bmp },
        { "image/svg+xml", ImagePartType.Svg },
    };

    /// <summary>
    /// Inspect the response headers of a web request and decode the mime type if provided
    /// </summary>
    /// <returns>Returns the extension of the image if provideds.</returns>
    private static bool TryInspectMimeType(string? contentType, out PartTypeInfo type)
    {
        // can be null when the protocol used doesn't allow response headers
        if (contentType != null &&
            knownContentType.TryGetValue(contentType, out type))
            return true;

        type = default;
        return false;
    }

    /// <summary>
    /// Gets the OpenXml PartTypeInfo associated to an image.
    /// </summary>
    private static bool TryGuessTypeFromUri(Uri uri, out PartTypeInfo type)
    {
        string extension = Path.GetExtension(uri.IsAbsoluteUri ? uri.Segments[uri.Segments.Length - 1] : uri.OriginalString);
        if (knownExtensions.TryGetValue(extension, out type)) return true;

        // extension not recognized, try with checking the query string. Expecting to resolve something like:
        // ./image.axd?picture=img1.jpg
        extension = Path.GetExtension(uri.IsAbsoluteUri ? uri.AbsoluteUri : uri.ToString());
        if (knownExtensions.TryGetValue(extension, out type)) return true;

        return false;
    }

    /// <summary>
    /// Gets the OpenXml PartTypeInfo associated to an image.
    /// </summary>
    private static bool TryGuessTypeFromStream(Stream stream, out PartTypeInfo type)
    {
        if (ImageHeader.TryDetectFileType(stream, out ImageHeader.FileType guessType))
        {
            switch (guessType)
            {
                case ImageHeader.FileType.Bitmap: type = ImagePartType.Bmp; return true;
                case ImageHeader.FileType.Emf: type = ImagePartType.Emf; return true;
                case ImageHeader.FileType.Gif: type = ImagePartType.Gif; return true;
                case ImageHeader.FileType.Jpeg: type = ImagePartType.Jpeg; return true;
                case ImageHeader.FileType.Png: type = ImagePartType.Png; return true;
            }
        }
        type = ImagePartType.Bmp;
        return false;
    }

    /// <summary>
    /// Loads an image from a stream and grab its size.
    /// </summary>
    private static Size GetImageSize(Stream imageStream)
    {
        // Read only the size of the image
        try
        {
            return ImageHeader.GetDimensions(imageStream);
        }
        catch (ArgumentException)
        {
            return Size.Empty;
        }
    }
}
