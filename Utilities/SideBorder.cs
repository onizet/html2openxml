using System;
using System.ComponentModel;
using System.Globalization;
using System.Drawing;

namespace NotesFor.HtmlToOpenXml
{
	using w = DocumentFormat.OpenXml.Wordprocessing;
	using System.Collections.Generic;


	/// <summary>
	/// Represents a Html Unit (ie: 120px, 10em, ...).
	/// </summary>
	[System.Diagnostics.DebuggerDisplay("{DebuggerDisplay,nq}")]
	struct SideBorder
	{
		/// <summary>Represents an empty border (not defined).</summary>
		public static readonly SideBorder Empty = new SideBorder();

		private w.BorderValues style;
		private Color color;
		private Unit size;


		public SideBorder(w.BorderValues style, Color color, Unit size)
		{
			this.style = style;
			this.color = color;
			this.size = size;
		}

		public static SideBorder Parse(String str)
		{
			if (str == null) return new SideBorder();

			// The properties of a border that can be set, are (in order): border-width, border-style, and border-color.
			// It does not matter if one of the values above are missing, e.g. border:solid #ff0000; is allowed.
			// The main problem for parsing this attribute is that the browsers allow any permutation of the values... meaning more coding :(
			// http://www.w3schools.com/cssref/pr_border.asp

			List<String> borderParts = new List<String>(str.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));
			if (borderParts.Count == 0) return SideBorder.Empty;

			// Initialize default values
			Unit borderWidth = Unit.Empty;
			Color borderColor = Color.Empty;
			w.BorderValues borderStyle = w.BorderValues.Nil;

			// Now try to guess the values with their permutation

			// handle border style
			for (int i = 0; i < borderParts.Count; i++)
			{
				borderStyle = ConverterUtility.ConvertToBorderStyle(borderParts[i]);
				if (borderStyle != w.BorderValues.Nil)
				{
					borderParts.RemoveAt(i); // no need to process this part anymore
					break;
				}
			}

			for (int i = 0; i < borderParts.Count; i++)
			{
				borderWidth = ParseWidth(borderParts[i]);
				if (borderWidth.IsValid)
				{
					borderParts.RemoveAt(i); // no need to process this part anymore
					break;
				}
			}

			// find width
			if(borderParts.Count > 0)
				borderColor = ConverterUtility.ConvertToForeColor(borderParts[0]);

			// returns the instance with default value if needed.
			// These value are the ones used by the browser, i.e: solid 3px black
			return new SideBorder(
				borderStyle == w.BorderValues.Nil? w.BorderValues.Single : borderStyle,
				borderColor.IsEmpty? Color.Black : borderColor,
				borderWidth.IsValid? borderWidth : new Unit(UnitMetric.Pixel, 4));
		}

		internal static Unit ParseWidth(String borderWidth)
		{
			Unit bu = Unit.Parse(borderWidth);
			if (bu.IsValid)
			{
				if (bu.Value > 0 && bu.Type == UnitMetric.Pixel)
					return bu;
			}
			else
			{
				switch (borderWidth)
				{
					case "thin": return new Unit(UnitMetric.Pixel, 1);
					case "medium": return new Unit(UnitMetric.Pixel, 3);
					case "thick": return new Unit(UnitMetric.Pixel, 5);
				}
			}

			return Unit.Empty;
		}

		//____________________________________________________________________
		//

		/// <summary>
		/// Gets or sets the type of border (solid, dashed, dotted, ...)
		/// </summary>
		public w.BorderValues Style
		{
			get { return style; }
			set { style = value; }
		}

		/// <summary>
		/// Gets or sets the color of the border.
		/// </summary>
		public Color Color
		{
			get { return color; }
			set { color = value; }
		}

		/// <summary>
		/// Gets or sets the size of the border expressed with its unit.
		/// </summary>
		public Unit Width
		{
			get { return size; }
			set { size = value; }
		}

		/// <summary>
		/// Gets whether the border is well formed and not empty.
		/// </summary>
		public bool IsValid
		{
			get { return this.Style != w.BorderValues.Nil; }
		}

		private string DebuggerDisplay
		{
			get { return String.Format("{{Border={0} {1} {2}}}", Style, Width.DebuggerDisplay, Color); }
		}
	}
}