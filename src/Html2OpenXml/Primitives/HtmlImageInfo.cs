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
using DocumentFormat.OpenXml.Packaging;

namespace HtmlToOpenXml
{
	/// <summary>
	/// Represents an image and its metadata.
	/// </summary>
	sealed class HtmlImageInfo
	{
		/// <summary>
		/// The URI identifying this cached image information.
		/// </summary>
		public string Source { get; set; }

		/// <summary>
		/// The Unique identifier of the ImagePart in the <see cref="MainDocumentPart"/>.
		/// </summary>
		public string ImagePartId { get; set; }

		/// <summary>
		/// Gets or sets the size of the image
		/// </summary>
		public Size Size { get; set; }
	}

	/// <summary>
	/// Typed dictionary of <see cref="HtmlImageInfo"/> where the Source URI is the identifier.
	/// </summary>
	sealed class HtmlImageInfoCollection : System.Collections.ObjectModel.KeyedCollection<string, HtmlImageInfo>
	{
		protected override string GetKeyForItem(HtmlImageInfo item)
		{
			return item.Source;
		}
	}
}