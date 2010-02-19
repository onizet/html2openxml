using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NotesFor.HtmlToOpenXml
{
	/// <summary>
	/// Specifies the position of an acronym or abbreviation in the resulting conversion.
	/// </summary>
	public enum AcronymPosition
	{
		/// <summary>
		/// Position at the end of the page.
		/// </summary>
		PageEnd = 0,
		/// <summary>
		/// Position at the end of the document.
		/// </summary>
		DocumentEnd = 1,
	}

	/// <summary>
	/// Specifies how the &lt;img&gt; tag will be handled during the conversion.
	/// </summary>
	public enum ImageProcessing
	{
		/// <summary>
		/// Image tag are not processed.
		/// </summary>
		Ignore,
		/// <summary>
		/// The image will be downloaded using a classic <see cref="System.Net.WebClient"/>. The src attribute should
		/// point on an absolute uri.
		/// </summary>
		AutomaticDownload,
		/// <summary>
		/// The image data will be provided by calling the <see cref="HtmlConverter.ProvisionImage"/> event.
		/// </summary>
		ManualProvisioning
	}
}
