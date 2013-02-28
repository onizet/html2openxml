using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NotesFor.HtmlToOpenXml
{
	using TagsAtSameLevel = System.ArraySegment<DocumentFormat.OpenXml.OpenXmlElement>;


	sealed class ParagraphStyleCollection : OpenXmlStyleCollection
	{
		private HtmlDocumentStyle documentStyle;


		internal ParagraphStyleCollection(HtmlDocumentStyle documentStyle)
		{
			this.documentStyle = documentStyle;
		}

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

		/// <summary>
		/// There is a few attributes shared by a large number of tags. This method will check them for a limited
		/// number of tags (&lt;p&gt;, &lt;pre&gt;, &lt;div&gt;, &lt;span&gt; and &lt;body&gt;).
		/// </summary>
		/// <returns>Returns true if the processing of this tag should generate a new paragraph.</returns>
		public bool ProcessCommonAttributes(HtmlEnumerator en, IList<OpenXmlElement> styleAttributes)
		{
			if (en.Attributes.Count == 0) return false;

			bool newParagraph = false;
			List<OpenXmlElement> containerStyleAttributes = new List<OpenXmlElement>();

			string attrValue = en.Attributes["lang"];
			if (attrValue != null && attrValue.Length > 0)
			{
				try
				{
					var ci = System.Globalization.CultureInfo.GetCultureInfo(attrValue);
					bool rtl = ci.TextInfo.IsRightToLeft;

					Languages lang = new Languages() { Val = ci.TwoLetterISOLanguageName };
					if (rtl)
					{
						lang.Bidi = ci.Name;
						styleAttributes.Add(new Languages() { Bidi = ci.Name });
					}

					containerStyleAttributes.Add(
						new ParagraphMarkRunProperties(lang));

					containerStyleAttributes.Add(new BiDi() { Val = OnOffValue.FromBoolean(rtl) });
				}
				catch (ArgumentException)
				{
					// lang not valid, ignore it
				}
			}


			attrValue = en.StyleAttributes["text-align"];
			if (attrValue != null && en.CurrentTag != "<font>")
			{
				JustificationValues? align = ConverterUtility.FormatParagraphAlign(attrValue);
				if (align.HasValue)
				{
					containerStyleAttributes.Add(new Justification { Val = align });
				}
			}

			// according to w3c, dir should be used in conjonction with lang. But whatever happens, we'll apply the RTL layout
			attrValue = en.Attributes["dir"];
			if (attrValue != null && attrValue.Equals("rtl", StringComparison.InvariantCultureIgnoreCase))
			{
				styleAttributes.Add(new RightToLeftText());
				containerStyleAttributes.Add(new Justification() { Val = JustificationValues.Right });
			}

			// <span> and <font> are considered as semi-container attribute. When converted to OpenXml, there are Runs but not Paragraphs
			if (en.CurrentTag == "<p>" || en.CurrentTag == "<div>" || en.CurrentTag == "<pre>")
			{
				var border = en.StyleAttributes.GetAsBorder("border");
				if (!border.IsEmpty)
				{
					ParagraphBorders borders = new ParagraphBorders();
					if (border.Top.IsValid) borders.Append(
						new TopBorder() { Val = border.Top.Style, Color = border.Top.Color.ToHexString(), Size = (uint) border.Top.Width.ValueInPx * 4, Space = 1U });
					if (border.Right.IsValid) borders.Append(
						new RightBorder() { Val = border.Right.Style, Color = border.Right.Color.ToHexString(), Size = (uint) border.Right.Width.ValueInPx * 4, Space = 1U });
					if (border.Bottom.IsValid) borders.Append(
						new BottomBorder() { Val = border.Bottom.Style, Color = border.Bottom.Color.ToHexString(), Size = (uint) border.Bottom.Width.ValueInPx * 4, Space = 1U });
					if (border.Left.IsValid) borders.Append(
						new LeftBorder() { Val = border.Left.Style, Color = border.Left.Color.ToHexString(), Size = (uint) border.Left.Width.ValueInPx * 4, Space = 1U });

					containerStyleAttributes.Add(borders);
					newParagraph = true;
				}
			}
			else if (en.CurrentTag == "<span>" || en.CurrentTag == "<font>")
			{
				// OpenXml limits the border to 4-side of the same color and style.
				SideBorder border = en.StyleAttributes.GetAsSideBorder("border");
				if (border.IsValid)
				{
					styleAttributes.Add(new DocumentFormat.OpenXml.Wordprocessing.Border() {
						Val = border.Style,
						Color = border.Color.ToHexString(),
						Size = (uint) border.Width.ValueInPx * 4,
						Space = 1U
					});
				}
			}

			String[] classes = en.Attributes.GetAsClass();
			if (classes != null)
			{
				for (int i = 0; i < classes.Length; i++)
				{
					string className = documentStyle.GetStyle(classes[i], StyleValues.Paragraph, ignoreCase: true);
					if (className != null)
					{
						styleAttributes.Add(new ParagraphStyleId() { Val = className });
						break;
					}
				}
			}

			Margin margin = en.StyleAttributes.GetAsMargin("margin");
			if (!margin.IsEmpty)
			{
				if (margin.Top.IsValid || margin.Bottom.IsValid)
				{
					SpacingBetweenLines spacing = new SpacingBetweenLines();
					if (margin.Top.IsValid) spacing.Before = margin.Top.ValueInDxa.ToString(CultureInfo.InvariantCulture);
					if (margin.Bottom.IsValid) spacing.After = margin.Bottom.ValueInDxa.ToString(CultureInfo.InvariantCulture);
					containerStyleAttributes.Add(spacing);
				}
				if (margin.Left.IsValid || margin.Right.IsValid)
				{
					Indentation indentation = new Indentation();
					if (margin.Left.IsValid) indentation.Left = margin.Left.ValueInDxa.ToString(CultureInfo.InvariantCulture);
					if (margin.Right.IsValid) indentation.Right = margin.Right.ValueInDxa.ToString(CultureInfo.InvariantCulture);
					containerStyleAttributes.Add(indentation);
				}
			}

			this.BeginTag(en.CurrentTag, containerStyleAttributes);

			// Process general run styles
			documentStyle.Runs.ProcessCommonAttributes(en, styleAttributes);

			return newParagraph;
		}
	}
}