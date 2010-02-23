using System;
using DocumentFormat.OpenXml.Packaging;

namespace NotesFor.HtmlToOpenXml
{
	/// <summary>
	/// The event arguments used for a ProvisionImage event.
	/// </summary>
	public class ProvisionImageEventArgs : EventArgs
	{
		internal ProvisionImageEventArgs(Uri uri)
		{
			this.ImageUrl = uri;
		}

		/// <summary>
		/// Gets the value of the href tag.
		/// </summary>
		public Uri ImageUrl { get; private set; }

		/// <summary>
		/// Gets the styles definition part located inside MainDocumentPart.
		/// </summary>
		public byte[] Data { get; set; }

		/// <summary>
		/// Gets or sets the format of the image.
		/// </summary>
		public ImagePartType? ImageExtension { get; set; }

		/// <summary>
		/// Gets or sets the width and height (in pixels) of the image as it should be displayed in the document.
		/// </summary>
		public System.Drawing.Size ImageSize { get; set; }
	}
}
