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
using DocumentFormat.OpenXml.Packaging;

namespace HtmlToOpenXml
{
	/// <summary>
	/// Represents an image and its metadata.
	/// </summary>
	sealed class HtmlImageInfo
	{
		/// <summary>
		/// Gets or sets the size of the image
		/// </summary>
		public Size Size { get; set; }

		/// <summary>
		/// Gets or sets the binary data of the image could read.
		/// </summary>
		public byte[] RawData { get; set; }

		/// <summary>
		/// Gets or sets the format of the image.
		/// </summary>
		public ImagePartType? Type { get; set; }
	}
}