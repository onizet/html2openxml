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
using DocumentFormat.OpenXml.Packaging;

namespace HtmlToOpenXml
{
	/// <summary>
	/// Download and provison the metadata of a requested image.
	/// </summary>
	sealed class ImageProvisioningProvider
	{
		// Map extension to ImagePartType
		private static readonly Dictionary<String, ImagePartType> knownExtensions = new Dictionary<String, ImagePartType>(StringComparer.OrdinalIgnoreCase) {
			{ ".gif", ImagePartType.Gif },
			{ ".bmp", ImagePartType.Bmp },
			{ ".emf", ImagePartType.Emf },
			{ ".ico", ImagePartType.Icon },
			{ ".jpeg", ImagePartType.Jpeg },
			{ ".jpg", ImagePartType.Jpeg },
			{ ".jpe", ImagePartType.Jpeg },
			{ ".pcx", ImagePartType.Pcx },
			{ ".png", ImagePartType.Png },
			{ ".tiff", ImagePartType.Tiff },
			{ ".wmf", ImagePartType.Wmf }
		};

		private readonly WebProxy proxy;
        private readonly Size preferredSize;
		private HtmlImageInfo imageInfo = new HtmlImageInfo();


		public ImageProvisioningProvider(WebProxy proxy, Size preferredSize)
		{
			this.proxy = proxy;
			this.preferredSize = preferredSize;
		}

		//____________________________________________________________________
		//
		// Public Functionality

		#region DownloadData

		/// <summary>
		/// Download the remote or local image located at the specified url.
		/// </summary>
		public HtmlImageInfo DownloadData(Uri imageUrl)
		{
			// is it a local path?
			if (imageUrl.IsFile)
			{
				// replace string %20 in LocalPath by daviderapicavoli (patch #15938)
				String localPath = Uri.UnescapeDataString(imageUrl.LocalPath);

				try
				{
                    // just read the picture from the file system
                    imageInfo.RawData = File.ReadAllBytes(localPath);
				}
				catch (Exception exc)
				{
					if (Logging.On) Logging.PrintError("ImageDownloader.DownloadData(\"" + localPath + "\")", exc);

					if (exc is IOException || exc is UnauthorizedAccessException || exc is System.Security.SecurityException || exc is NotSupportedException)
						return null;
					throw;
				}

				return imageInfo;
			}

			// data inline, encoded in base64
			if (imageUrl.Scheme == "data")
			{
				DataUri dataUri = DataUri.Parse(imageUrl.OriginalString);
				return DownloadData(dataUri);
			}

            var response = BackChannels.CreateWebRequest(imageUrl, proxy);
            if (response != null)
            {
                imageInfo.RawData = response.Body;

                // For requested url with no filename, we need to read the media mime type if provided
                imageInfo.Type = InspectMimeType(response.ContentType);
            }

            return imageInfo;
        }

		#endregion

		#region DownloadData

		/// <summary>
		/// Decrypt the given inline DataUri.
		/// </summary>
		public HtmlImageInfo DownloadData(DataUri dataUri)
		{
            if (dataUri == null)
                return null;

			ImagePartType type;
			if (knownContentType.TryGetValue(dataUri.Mime, out type))
				imageInfo.Type = type;

			imageInfo.RawData = dataUri.Data;
            return imageInfo;
		}

		#endregion

		#region Provision

		/// <summary>
		/// Discover the metadata of the image.
		/// </summary>
		public static bool Provision(HtmlImageInfo imageInfo, Uri imageUrl)
		{
			if (imageInfo?.RawData == null) return false;

			if (!imageInfo.Type.HasValue)
				imageInfo.Type = GetImagePartTypeForImageUrl(imageUrl);

			if (!imageInfo.Type.HasValue)
				return false;

			if (imageInfo.Size.Width == 0 || imageInfo.Size.Height == 0)
			{
				using (Stream outputStream = new MemoryStream(imageInfo.RawData))
					imageInfo.Size = GetImageSize(outputStream);
			}

			return true;
		}

		#endregion

		//____________________________________________________________________
		//
		// Private Implementation

		#region InspectMimeType

		// http://stackoverflow.com/questions/58510/using-net-how-can-you-find-the-mime-type-of-a-file-based-on-the-file-signature
		private static Dictionary<String, ImagePartType> knownContentType = new Dictionary<String, ImagePartType>(StringComparer.OrdinalIgnoreCase) {
			{ "image/gif", ImagePartType.Gif },
            { "image/pjpeg", ImagePartType.Jpeg },
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
			{ "image/bmp", ImagePartType.Bmp }
		};

		/// <summary>
		/// Inspect the response headers of a web request and decode the mime type if provided
		/// </summary>
		/// <returns>Returns the extension of the image if provideds.</returns>
		private static ImagePartType? InspectMimeType(string contentType)
		{
    		// can be null when the protocol used doesn't allow response headers
			if (contentType == null) return null;

			ImagePartType type;
			if (knownContentType.TryGetValue(contentType, out type))
				return type;

			return null;
		}

		#endregion

		#region GetImagePartTypeForImageUrl

		/// <summary>
		/// Gets the OpenXml ImagePartType associated to an image.
		/// </summary>
		public static ImagePartType? GetImagePartTypeForImageUrl(Uri uri)
		{
            ImagePartType type;
            String extension = System.IO.Path.GetExtension(uri.IsAbsoluteUri ? uri.Segments[uri.Segments.Length - 1] : uri.OriginalString);
            if (knownExtensions.TryGetValue(extension, out type)) return type;

            // extension not recognized, try with checking the query string. Expecting to resolve something like:
            // ./image.axd?picture=img1.jpg
            extension = System.IO.Path.GetExtension(uri.IsAbsoluteUri ? uri.AbsoluteUri : uri.ToString());
            if (knownExtensions.TryGetValue(extension, out type)) return type;

            // so, match text of the form: data:image/yyy;base64,zzzzzzzzzzzz...
            // where yyy is the MIME type, zzz is the base64 encoded data
            DataUri dataUri = DataUri.Parse(uri.ToString());
            if (dataUri != null)
            {
                if (knownContentType.TryGetValue(dataUri.Mime, out type)) return type;
            }

            return null;
		}

		#endregion

		#region GetImageSize

		/// <summary>
		/// Loads an image from a stream and grab its size.
		/// </summary>
		public static Size GetImageSize(Stream imageStream)
		{
            // Read only the size of the image using a little API found on codeproject.
            try
            {
                return ImageHeader.GetDimensions(imageStream);
            }
            catch (ArgumentException)
            {
                return Size.Empty;
            }
        }

        #endregion
	}
}