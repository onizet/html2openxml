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

    /// <summary>
    /// Predefined quote style as defined by the browser (used for the &lt;q&gt; tag).
    /// </summary>
    public sealed class QuoteChars
    {
        /// <summary>Internet Explorer style: « abc » </summary>
        public static readonly QuoteChars IE = new QuoteChars("« ", " »");
        /// <summary>Firefox style: “abc”</summary>
        public static readonly QuoteChars Gecko = new QuoteChars("“", "”");
        /// <summary>Chrome/Safari/Opera style: "abc"</summary>
        public static readonly QuoteChars WebKit = new QuoteChars("\"", "\"");
        internal readonly String[] chars;

		/// <summary>
		/// Initializes a new instance of <see cref="QuoteChars"/> class.
		/// </summary>
		/// <param name="begin">The characters at the beginning of a quote.</param>
		/// <param name="end">The characters at the end of a quote.</param>
        public QuoteChars(String begin, String end)
	    {
            this.chars = new String[] { begin, end };
	    }
    }
}