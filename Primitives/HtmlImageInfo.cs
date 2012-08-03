using System;
using System.Drawing;
using DocumentFormat.OpenXml.Packaging;

namespace NotesFor.HtmlToOpenXml
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
		/// Gets or sets the binary data of the image (something <see cref="System.Drawing.Image"/> could read.
		/// </summary>
		public byte[] RawData { get; set; }

		/// <summary>
		/// Gets or sets the format of the image.
		/// </summary>
		public ImagePartType? Type { get; set; }
	}
}