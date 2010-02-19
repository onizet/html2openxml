using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NotesFor.HtmlToOpenXml
{
	using TagsAtSameLevel = System.ArraySegment<DocumentFormat.OpenXml.OpenXmlElement>;


	sealed class ParagraphStyleCollection : OpenXmlStyleCollection
	{
		/// <summary>
		/// Apply all the current Html tag (Paragraph properties) to the specified paragrah.
		/// </summary>
		public override void ApplyTags(OpenXmlCompositeElement paragraph)
		{
			if (tags.Count == 0) return;

			ParagraphProperties properties = paragraph.GetFirstChild<ParagraphProperties>();
			if (properties == null) paragraph.PrependChild<ParagraphProperties>(properties = new ParagraphProperties());

			var en = tags.GetEnumerator();
			while (en.MoveNext())
			{
				TagsAtSameLevel tagsOfSameLevel = en.Current.Value.Peek();
				foreach (OpenXmlElement tag in tagsOfSameLevel.Array)
					properties.Append(tag.CloneNode(true));
			}
		}

		/// <summary>
		/// Gets the default StyleId to apply on the any new paragraph.
		/// </summary>
		internal String DefaultParagraphStyle { get; set; }

		public Paragraph NewParagraph()
		{
			Paragraph p = new Paragraph();
			if (this.DefaultParagraphStyle != null)
				p.InsertInProperties(new ParagraphStyleId() { Val = this.DefaultParagraphStyle });
			return p;
		}
	}
}
