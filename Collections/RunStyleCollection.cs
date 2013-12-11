using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace NotesFor.HtmlToOpenXml
{
	using TagsAtSameLevel = System.ArraySegment<DocumentFormat.OpenXml.OpenXmlElement>;


	sealed class RunStyleCollection : OpenXmlStyleCollectionBase
	{
		private HtmlDocumentStyle documentStyle;
		private static GetSequenceNumberHandler getTagOrderHandler;


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
					SetProperties(properties, tag.CloneNode(true));
			}

			if (this.DefaultRunStyle != null)
				properties.RunStyle = new RunStyle() { Val = this.DefaultRunStyle };
		}

		#region ProcessCommonAttributes

		/// <summary>
		/// Converts some common styling attributes to their OpenXml equivalence.
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

			HtmlFont font = en.StyleAttributes.GetAsFont("font");
			if (!font.IsEmpty)
			{
				if (font.Style == FontStyle.Italic)
					styleAttributes.Add(new Italic());

				if (font.Weight == FontWeight.Bold || font.Weight == FontWeight.Bolder)
					styleAttributes.Add(new Bold());

				if (font.Variant == FontVariant.SmallCaps)
					styleAttributes.Add(new SmallCaps());

				if (font.Family != null)
					styleAttributes.Add(new RunFonts() { Ascii = font.Family.Name, HighAnsi = font.Family.Name });

				// size are half-point font size
				if (font.Size.IsValid)
					styleAttributes.Add(new FontSize() { Val = (font.Size.ValueInPoint * 2).ToString(CultureInfo.InvariantCulture) });
			}
		}

		#endregion

		#region GetTagOrder
		
		protected override int GetTagOrder(OpenXmlElement element)
		{
			// I don't want to hard-code the sequence number of the child elements of a RunProperties.
			// I prefer relying on the OpenXml API and use a bit Reflection.
			if (getTagOrderHandler == null)
			{
				var mi = typeof(OpenXmlCompositeElement)
					.GetMethod("GetSequenceNumber", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);

				// We use a dummy new RunProperties instance
				getTagOrderHandler = (GetSequenceNumberHandler)
					Delegate.CreateDelegate(typeof(GetSequenceNumberHandler), new RunProperties(), mi, true);
			}

			// Create a delegate to speed up the invocation to the GetSequenceNumber method
			return (int) getTagOrderHandler.DynamicInvoke(element);
		}

		#endregion

		//____________________________________________________________________
		//
		// Properties

		/// <summary>
		/// Gets the default StyleId to apply on the any new runs.
		/// </summary>
		internal String DefaultRunStyle { get; set; }
	}
}