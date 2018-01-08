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
using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml
{
	/// <summary>
	/// Provides some utilies methods for translating Http attributes to OpenXml elements.
	/// </summary>
	static class Converter
	{
		#region ToParagraphAlign

		/// <summary>
		/// Convert the Html text align attribute (horizontal alignement) to its corresponding OpenXml value.
		/// </summary>
		public static JustificationValues? ToParagraphAlign(string htmlAlign)
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

		#region ToVAlign

		/// <summary>
		/// Convert the Html vertical-align attribute to its corresponding OpenXml value.
		/// </summary>
		public static TableVerticalAlignmentValues? ToVAlign(string htmlAlign)
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

		#region ToFontSize

		/// <summary>
		/// Convert Html regular font-size to OpenXml font value (expressed in point).
		/// </summary>
		public static Unit ToFontSize(string htmlSize)
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

		#region ToFontVariant

		public static FontVariant? ToFontVariant(string html)
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

		#region ToFontStyle

		public static FontStyle? ToFontStyle(string html)
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

		#region ToFontWeight

		public static FontWeight? ToFontWeight(string html)
		{
			if (html == null) return null;
			switch (html.ToLowerInvariant())
			{
                case "700":
				case "bold": return FontWeight.Bold;
				case "bolder": return FontWeight.Bolder;
                case "400":
                case "normal": return FontWeight.Normal;
				default: return null;
			}
		}

		#endregion

		#region ToFontFamily

		public static string ToFontFamily(string str)
		{
			String[] names = str.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
			for (int i = 0; i < names.Length; i++)
			{
                String fontName = names[i];
				try
				{
                    if (fontName[0] == '\'' && fontName[fontName.Length-1] == '\'') fontName = fontName.Substring(1, fontName.Length - 2);
					return fontName;
				}
				catch (ArgumentException)
				{
					// the name is not a TrueType font or is not a font installed on this computer
				}
			}

			return null;
		}

		#endregion

		#region ToBorderStyle

		public static BorderValues ToBorderStyle(string borderStyle)
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

		#region ToUnitMetric

		public static UnitMetric ToUnitMetric(String type)
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

		#region ToPageOrientation

		public static PageOrientationValues ToPageOrientation(String orientation)
		{
			if (String.Equals(orientation, "landscape", StringComparison.OrdinalIgnoreCase))
				return PageOrientationValues.Landscape;

			return PageOrientationValues.Portrait;
		}

		#endregion

		#region ToTextDecoration

		public static TextDecoration ToTextDecoration(string html)
        {
			// this style could take multiple values separated by a space
			// ex: text-decoration: blink underline;

			TextDecoration decoration = TextDecoration.None;

			if (html == null) return decoration;
			foreach (string part in html.ToLowerInvariant().Split(' '))
			{
				switch (part)
				{
					case "underline": decoration |= TextDecoration.Underline; break;
					case "line-through": decoration |= TextDecoration.LineThrough; break;
					default: break; // blink and overline are not supported
				}
			}
			return decoration;
		}

		#endregion
	}
}