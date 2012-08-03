using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NotesFor.HtmlToOpenXml
{
	/// <summary>
	/// The event arguments used for a StyleMissing event.
	/// </summary>
	public class StyleEventArgs : EventArgs
	{
		internal StyleEventArgs(String styleId, MainDocumentPart mainPart, StyleValues type)
		{
			this.Name = styleId;
			this.StyleDefinitionsPart = mainPart.StyleDefinitionsPart;
			this.Type = type;
		}

		/// <summary>
		/// Gets the invariant name of the style.
		/// </summary>
		public String Name { get; private set; }

		/// <summary>
		/// Gets the styles definition part located inside MainDocumentPart.
		/// </summary>
		public StyleDefinitionsPart StyleDefinitionsPart { get; private set; }

		/// <summary>
		/// Gets the type of style seeked (character or paragraph).
		/// </summary>
		public StyleValues Type { get; private set; }
	}
}