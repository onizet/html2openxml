using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NotesFor.HtmlToOpenXml
{
	using TagsAtSameLevel = System.ArraySegment<DocumentFormat.OpenXml.OpenXmlElement>;


	sealed class RunStyleCollection : OpenXmlStyleCollection
	{
		private HtmlDocumentStyle documentStyle;


		internal RunStyleCollection(HtmlDocumentStyle documentStyle)
		{
			this.documentStyle = documentStyle;
		}

		/// <summary>
		/// Apply all the current Html tag (Run properties) to the specified run.
		/// </summary>
		public override void ApplyTags(OpenXmlCompositeElement run)
		{
			if (tags.Count == 0 && DefaultRunStyle == null) return;

			RunProperties properties = run.GetFirstChild<RunProperties>();
			if (properties == null) run.PrependChild<RunProperties>(properties = new RunProperties());

			var en = tags.GetEnumerator();
			while (en.MoveNext())
			{
				TagsAtSameLevel tagsOfSameLevel = en.Current.Value.Peek();
				foreach (OpenXmlElement tag in tagsOfSameLevel.Array)
					properties.Append(tag.CloneNode(true));
			}

			if (this.DefaultRunStyle != null)
				properties.Append(new RunStyle() { Val = this.DefaultRunStyle });
		}

		/// <summary>
		/// Gets the default StyleId to apply on the any new runs.
		/// </summary>
		internal String DefaultRunStyle { get; set; }

		/// <summary>
		/// Move inside the current tag related to table (td, thead, tr, ...) and converts some common
		/// attributes to their OpenXml equivalence.
		/// </summary>
		/// <param name="styleAttributes">The collection of attributes where to store new discovered attributes.</param>
		public void ProcessCommonAttributes(HtmlEnumerator en, IList<OpenXmlElement> styleAttributes)
		{
			if (en.Attributes.Count == 0) return;

			var colorValue = en.StyleAttributes.GetAsColor("color");
			if (colorValue.IsEmpty) colorValue = en.Attributes.GetAsColor("color");
			if (!colorValue.IsEmpty)
				styleAttributes.Add(new Color { Val = colorValue.ToHexString() });

			colorValue = en.StyleAttributes.GetAsColor("background-color");
			if (!colorValue.IsEmpty)
			{
				// change the way the background-color renders. It now uses Shading instead of Highlight.
				// Changes brought by Wude on http://notesforhtml2openxml.codeplex.com/discussions/277570
				styleAttributes.Add(new Shading { Val = ShadingPatternValues.Clear, Fill = colorValue.ToHexString() });
			}

			string attrValue = en.StyleAttributes["text-decoration"];
			if (attrValue == "underline")
			{
				styleAttributes.Add(new Underline { Val = UnderlineValues.Single });
			}
			else if (attrValue == "line-through")
			{
				styleAttributes.Add(new Strike());
			}

			attrValue = en.StyleAttributes["font-style"];
			if (attrValue == "italic" || attrValue == "oblique")
			{
				styleAttributes.Add(new Italic());
			}

			attrValue = en.StyleAttributes["font-weight"];
			if (attrValue == "bold" || attrValue == "bolder")
			{
				styleAttributes.Add(new Bold());
			}

			String[] classes = en.Attributes.GetAsClass();
			if (classes != null)
			{
				for (int i = 0; i < classes.Length; i++)
				{
					string className = documentStyle.GetStyle(classes[i], StyleValues.Character, ignoreCase: true);
					if (className != null) // only one Style can be applied in OpenXml and dealing with inheritance is out of scope
					{
						styleAttributes.Add(new RunStyle() { Val = className });
						break;
					}
				}
			}

			// We ignore font-family and font-size voluntarily because the user oftenly copy-paste from web pages
			// but don't want to see these font in the report.
		}
	}
}