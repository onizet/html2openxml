using System;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Drawing;
using System.Globalization;

namespace NotesFor.HtmlToOpenXml
{
	/// <summary>
	/// Provides some utilies methods for translating Http attributes to OpenXml elements.
	/// </summary>
	static class ConverterUtility
	{
		#region FormatParagraphAlign

		/// <summary>
		/// Convert the Html text align attribute (horizontal alignement) to its corresponding OpenXml value.
		/// </summary>
		public static JustificationValues? FormatParagraphAlign(string htmlAlign)
		{
			if (htmlAlign == null) return null;
			switch (htmlAlign.ToLowerInvariant())
			{
				case "left": return JustificationValues.Left;
				case "right": return JustificationValues.Right;
				case "center": return JustificationValues.Center;
				case "justify": return JustificationValues.Both;
			}

			return null;
		}

		#endregion

		#region FormatVAlign

		/// <summary>
		/// Convert the Html vertical-align attribute to its corresponding OpenXml value.
		/// </summary>
		public static TableVerticalAlignmentValues? FormatVAlign(string htmlAlign)
		{
			if (htmlAlign == null) return null;
			switch (htmlAlign.ToLowerInvariant())
			{
				case "top": return TableVerticalAlignmentValues.Top;
				case "middle": return TableVerticalAlignmentValues.Center;
				case "bottom": return TableVerticalAlignmentValues.Bottom;
			}

			return null;
		}

		#endregion

		#region ConvertToFontSize

		/// <summary>
		/// Convert Html regular font-size to OpenXml font value (expressed in point).
		/// </summary>
		public static Unit ConvertToFontSize(string htmlSize)
		{
			if (htmlSize == null) return Unit.Empty;
			switch (htmlSize.ToLowerInvariant())
			{
				case "1":
				case "xx-small": return new Unit(UnitMetric.Point, 10);
				case "2":
				case "x-small": return new Unit(UnitMetric.Point, 15);
				case "3":
				case "small": return new Unit(UnitMetric.Point, 20);
				case "4":
				case "medium": return new Unit(UnitMetric.Point, 27);
				case "5":
				case "large": return new Unit(UnitMetric.Point, 36);
				case "6":
				case "x-large": return new Unit(UnitMetric.Point, 48);
				case "7":
				case "xx-large": return new Unit(UnitMetric.Point, 72);
				default:
					// the font-size is specified in positive half-points
					Unit unit = Unit.Parse(htmlSize);
					if (!unit.IsValid || unit.Value <= 0)
						return Unit.Empty;

					// this is a rough conversion to support some percent size, considering 100% = 11 pt
					if (unit.Type == UnitMetric.Percent) unit = new Unit(UnitMetric.Point, unit.Value * 0.11);
					return unit;
			}
		}

		#endregion

		#region ConvertToFontVariant

		public static FontVariant? ConvertToFontVariant(string html)
		{
			if (html == null) return null;

			switch (html.ToLowerInvariant())
			{
				case "small-caps": return FontVariant.SmallCaps;
				case "normal": return FontVariant.Normal;
				default: return null;
			}
		}

		#endregion

		#region ConvertToFontStyle

		public static FontStyle? ConvertToFontStyle(string html)
		{
			if (html == null) return null;
			switch (html.ToLowerInvariant())
			{
				case "italic":
				case "oblique": return FontStyle.Italic;
				case "normal": return FontStyle.Normal;
				default: return null;
			}
		}

		#endregion

		#region ConvertToFontWeight

		public static FontWeight? ConvertToFontWeight(string html)
		{
			if (html == null) return null;
			switch (html.ToLowerInvariant())
			{
				case "bold": return FontWeight.Bold;
				case "bolder": return FontWeight.Bolder;
				case "normal": return FontWeight.Normal;
				default: return null;
			}
		}

		#endregion

		#region ConvertToForeColor

		public static System.Drawing.Color ConvertToForeColor(string htmlColor)
		{
			System.Drawing.Color color;

			// Bug fixed by jairoXXX to support rgb(r,g,b) format
			if (htmlColor.StartsWith("rgb(", StringComparison.InvariantCultureIgnoreCase))
			{
				var colorStringArray = htmlColor.Substring(4, htmlColor.LastIndexOf(')') - 4).Split(',');

				return System.Drawing.Color.FromArgb(
					Int32.Parse(colorStringArray[0], NumberStyles.Integer, CultureInfo.InvariantCulture),
					Int32.Parse(colorStringArray[1], NumberStyles.Integer, CultureInfo.InvariantCulture),
					Int32.Parse(colorStringArray[2], NumberStyles.Integer, CultureInfo.InvariantCulture));
			}

			// The Html allows to write color in hexa without the preceding '#'
			// I just ensure it's a correct hexadecimal value (length=6 and first character should be
			// a digit or an hexa letter)
			if (htmlColor.Length == 6 && (Char.IsDigit(htmlColor[0]) || (htmlColor[0] >= 'a' && htmlColor[0] <= 'f')
				|| (htmlColor[0] >= 'A' && htmlColor[0] <= 'F')))
			{
				try
				{
					color = System.Drawing.Color.FromArgb(
						Convert.ToInt32(htmlColor.Substring(0, 2), 16),
						Convert.ToInt32(htmlColor.Substring(2, 2), 16),
						Convert.ToInt32(htmlColor.Substring(4, 2), 16));
				}
				catch (System.FormatException)
				{
					// If the conversion failed, that should be a named color
					// Let the framework dealing with it
					color = System.Drawing.ColorTranslator.FromHtml(htmlColor);
				}
			}
			else
			{
				color = System.Drawing.ColorTranslator.FromHtml(htmlColor);
			}

			return color;
		}

		#endregion

		#region ConvertToBorderStyle

		public static BorderValues ConvertToBorderStyle(string borderStyle)
		{
			if (borderStyle == null) return BorderValues.Nil;
			switch (borderStyle.ToLowerInvariant())
			{
				case "dotted": return BorderValues.Dotted;
				case "dashed": return BorderValues.Dashed;
				case "solid": return BorderValues.Single;
				case "double": return BorderValues.Double;
				case "inset": return BorderValues.Inset;
				case "outset": return BorderValues.Outset;
				case "none": return BorderValues.None;
				default: return BorderValues.Nil;
			}
		}

		#endregion

		#region ConvertToUnitMetric

		public static UnitMetric ConvertToUnitMetric(String type)
		{
			if (type == null) return UnitMetric.Unknown;
			switch (type.ToLowerInvariant())
			{
				case "%": return UnitMetric.Percent;
				case "in": return UnitMetric.Inch;
				case "cm": return UnitMetric.Centimeter;
				case "mm": return UnitMetric.Millimeter;
				case "em": return UnitMetric.EM;
				case "ex": return UnitMetric.Ex;
				case "pt": return UnitMetric.Point;
				case "pc": return UnitMetric.Pica;
				case "px": return UnitMetric.Pixel;
				default: return UnitMetric.Unknown;
			}
		}

		#endregion

		//____________________________________________________________________
		//
		// Private Implementation

		static readonly char[] hexDigits = {
         '0', '1', '2', '3', '4', '5', '6', '7',
         '8', '9', 'A', 'B', 'C', 'D', 'E', 'F'};


		#region Color ToHexString

		/// <summary>
		/// Convert a .Net Color to a hex string.
		/// </summary>
		public static string ToHexString(this System.Drawing.Color color)
		{
			// http://www.cambiaresearch.com/c4/24c09e15-2941-4ad2-8695-00b1b4029f4d/Convert-dotnet-Color-to-Hex-String.aspx

			byte[] bytes = new byte[3];
			bytes[0] = color.R;
			bytes[1] = color.G;
			bytes[2] = color.B;
			char[] chars = new char[bytes.Length * 2];
			for (int i = 0; i < bytes.Length; i++)
			{
				int b = bytes[i];
				chars[i * 2] = hexDigits[b >> 4];
				chars[i * 2 + 1] = hexDigits[b & 0xF];
			}
			return new string(chars);
		}

		#endregion
	}
}