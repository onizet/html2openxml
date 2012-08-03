using System;
using DocumentFormat.OpenXml.Packaging;

namespace NotesFor.HtmlToOpenXml
{
	/// <summary>
	/// The event arguments used for a ProvisionImage event.
	/// </summary>
	public class ProvisionImageEventArgs : EventArgs
	{
		private HtmlImageInfo info;


		internal ProvisionImageEventArgs(Uri uri, HtmlImageInfo info)
		{
			this.ImageUrl = uri;
			this.info = info;
		}

		/// <summary>
		/// Gets the value of the href tag.
		/// </summary>
		public Uri ImageUrl { get; private set; }

		/// <summary>
		/// Gets the styles definition part located inside MainDocumentPart.
		/// </summary>
		public byte[] Data
		{
			get { return info.RawData; }
			set { info.RawData = value; }
		}

		/// <summary>
		/// Gets or sets the format of the image.
		/// </summary>
		public ImagePartType? ImageExtension
		{
			get { return info.Type; }
			set { info.Type = value; }
		}

		/// <summary>
		/// Gets or sets the width and height (in pixels) of the image as it should be displayed in the document.
		/// </summary>
		public System.Drawing.Size ImageSize
		{
			get { return info.Size; }
			set { info.Size = value; }
		}
	}
}