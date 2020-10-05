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
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml
{
	using TagsAtSameLevel = System.ArraySegment<DocumentFormat.OpenXml.OpenXmlElement>;


	sealed class ParagraphStyleCollection : OpenXmlStyleCollectionBase
	{
		private readonly HtmlDocumentStyle documentStyle;

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
					SetProperties(properties, tag.CloneNode(true));
			}
		}

		#region NewParagraph

		/// <summary>
		/// Factor method to create a new Paragraph with its default style already defined.
		/// </summary>
		public Paragraph NewParagraph()
		{
			Paragraph p = new Paragraph();
			if (this.DefaultParagraphStyle != null)
				p.InsertInProperties(prop => prop.ParagraphStyleId = new ParagraphStyleId() { Val = this.DefaultParagraphStyle });
			return p;
		}

		#endregion

		#region ProcessCommonAttributes

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
#if !NETSTANDARD1_3
                    var ci = System.Globalization.CultureInfo.GetCultureInfo(attrValue);
#else
                    var ci = new System.Globalization.CultureInfo(attrValue);
#endif
                    bool rtl = ci.TextInfo.IsRightToLeft;

					Languages lang = new Languages() { Val = ci.TwoLetterISOLanguageName };
					if (rtl)
					{
						lang.Bidi = ci.Name;
						styleAttributes.Add(new Languages() { Bidi = ci.Name });

						// notify table
						documentStyle.Tables.BeginTag(en.CurrentTag, new TableJustification() { Val = TableRowAlignmentValues.Right });
					}

					containerStyleAttributes.Add(new ParagraphMarkRunProperties(lang));
					containerStyleAttributes.Add(new BiDi() { Val = OnOffValue.FromBoolean(rtl) });
				}
				catch (ArgumentException exc)
				{
                    // lang not valid, ignore it
                    if (Logging.On) Logging.PrintError($"lang attribute {attrValue} not recognized: " + exc.Message, exc);
				}
			}


			attrValue = en.StyleAttributes["text-align"];
			if (attrValue != null && en.CurrentTag != "<font>")
			{
				JustificationValues? align = Converter.ToParagraphAlign(attrValue);
				if (align.HasValue)
				{
					containerStyleAttributes.Add(new Justification { Val = align });
				}
			}

			// according to w3c, dir should be used in conjonction with lang. But whatever happens, we'll apply the RTL layout
			attrValue = en.Attributes["dir"];
			if (attrValue != null)
			{
				if (attrValue.Equals("rtl", StringComparison.OrdinalIgnoreCase))
				{
					styleAttributes.Add(new RightToLeftText());
					containerStyleAttributes.Add(new Justification() { Val = JustificationValues.Right });
				}
				else if (attrValue.Equals("ltr", StringComparison.OrdinalIgnoreCase))
				{
					containerStyleAttributes.Add(new Justification() { Val = JustificationValues.Left });
				}
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
                    if (border.Left.IsValid) borders.Append(
                        new LeftBorder() { Val = border.Left.Style, Color = border.Left.Color.ToHexString(), Size = (uint) border.Left.Width.ValueInPx * 4, Space = 1U });
                    if (border.Bottom.IsValid) borders.Append(
                        new BottomBorder() { Val = border.Bottom.Style, Color = border.Bottom.Color.ToHexString(), Size = (uint) border.Bottom.Width.ValueInPx * 4, Space = 1U });
                    if (border.Right.IsValid) borders.Append(
						new RightBorder() { Val = border.Right.Style, Color = border.Right.Color.ToHexString(), Size = (uint) border.Right.Width.ValueInPx * 4, Space = 1U });

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
						containerStyleAttributes.Add(new ParagraphStyleId() { Val = className });
						break;
					}
				}
			}

			Margin margin = en.StyleAttributes.GetAsMargin("margin");
            Indentation indentation = null;
            if (!margin.IsEmpty)
			{
                if (margin.Top.IsFixed || margin.Bottom.IsFixed)
				{
					SpacingBetweenLines spacing = new SpacingBetweenLines();
                    if (margin.Top.IsFixed) spacing.Before = margin.Top.ValueInDxa.ToString(CultureInfo.InvariantCulture);
                    if (margin.Bottom.IsFixed) spacing.After = margin.Bottom.ValueInDxa.ToString(CultureInfo.InvariantCulture);
					containerStyleAttributes.Add(spacing);
				}
                if (margin.Left.IsFixed || margin.Right.IsFixed)
				{
					indentation = new Indentation();
                    if (margin.Left.IsFixed) indentation.Left = margin.Left.ValueInDxa.ToString(CultureInfo.InvariantCulture);
                    if (margin.Right.IsFixed) indentation.Right = margin.Right.ValueInDxa.ToString(CultureInfo.InvariantCulture);
					containerStyleAttributes.Add(indentation);
				}
			}

            // implemented by giorand (feature #13787)
            Unit textIndent = en.StyleAttributes.GetAsUnit("text-indent");
            if (textIndent.IsValid && (en.CurrentTag == "<p>" || en.CurrentTag == "<div>"))
            {
                if (indentation == null) indentation = new Indentation();
                indentation.FirstLine = textIndent.ValueInDxa.ToString(CultureInfo.InvariantCulture);
                containerStyleAttributes.Add(indentation);
            }

			this.BeginTag(en.CurrentTag, containerStyleAttributes);

			// Process general run styles
			documentStyle.Runs.ProcessCommonAttributes(en, styleAttributes);

			return newParagraph;
		}

        #endregion

        //____________________________________________________________________
        //
        // Properties

        /// <summary>
        /// Gets the default StyleId to apply on the any new paragraph.
        /// </summary>
        internal String DefaultParagraphStyle { get; set; }
	}
}