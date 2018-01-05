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
using System.Text.RegularExpressions;

namespace HtmlToOpenXml
{
	using w = DocumentFormat.OpenXml.Wordprocessing;


	/// <summary>
	/// Represents the collection of attributes present in the current html tag.
	/// </summary>
	sealed class HtmlAttributeCollection
	{
		// This regex split the attributes. This line is valid and all the attributes are well discovered:
		// <table border="1" contenteditable style="text-align: center; color: #ff00e6" cellpadding=0 cellspacing='0' align="center">
		// RegexOptions.Singleline stands for dealing with attributes that contain newline (typically for base64 image, see issue #8)
		private static Regex stripAttributesRegex = new Regex(@"
#tag and its value surrounded by "" or '
((?<tag>\w+)=(?<sep>""|')\s*(?<val>\#?.*?)(\k<sep>|>))
|
# tag whereas the value is not delimited: cellspacing=0
(?<tag>\w+)=(?<val>\w+)
|
# single tag (with no value): contenteditable
\b(?<tag>\w+)\b", RegexOptions.IgnorePatternWhitespace| RegexOptions.Singleline);

        private static Regex stripStyleAttributesRegex = new Regex(@"(?<name>.+?):\s*(?<val>[^;]+);*\s*");

		private Dictionary<string, string> attributes;



		private HtmlAttributeCollection()
		{
			this.attributes = new Dictionary<string, string>();
		}

		public static HtmlAttributeCollection Parse(String htmlTag)
		{
			HtmlAttributeCollection collection = new HtmlAttributeCollection();
			if (String.IsNullOrEmpty(htmlTag)) return collection;

			// We remove the name of the tag (due to our regex) and ensure there are at least one parameter
			int startIndex;
			for (startIndex = 0; startIndex < htmlTag.Length; startIndex++)
			{
				if (Char.IsWhiteSpace(htmlTag[startIndex]))
				{
					startIndex++;
					break;
				}
				else if (htmlTag[startIndex] == '>' || htmlTag[startIndex] == '/')
				{
					// no attribute in this tag
					return collection;
				}
			}

			MatchCollection matches = stripAttributesRegex.Matches(htmlTag, startIndex);
			foreach (Match m in matches)
			{
				collection.attributes[m.Groups["tag"].Value] = m.Groups["val"].Value;
			}

			return collection;
		}

		public static HtmlAttributeCollection ParseStyle(String htmlTag)
		{
			HtmlAttributeCollection collection = new HtmlAttributeCollection();
			if (String.IsNullOrEmpty(htmlTag)) return collection;

            // Encoded ':' and ';' characters are valid for browser but not handled by the regex (bug #13812 reported by robin391)
            // ex= <span style="text-decoration&#58;underline&#59;color:red">
			MatchCollection matches = stripStyleAttributesRegex.Matches(HttpUtility.HtmlDecode(htmlTag));
			foreach (Match m in matches)
				collection.attributes[m.Groups["name"].Value] = m.Groups["val"].Value;

			return collection;
		}

		/// <summary>
		/// Gets the number of attributes for this tag.
		/// </summary>
		public int Count
		{
			get { return attributes.Count; }
		}

		/// <summary>
		/// Gets the named attribute.
		/// </summary>
		public String this[String name]
		{
			get
            {
                string value;
                return attributes.TryGetValue(name, out value)? value : null;
            }
		}

		/// <summary>
		/// Gets an attribute representing an integer.
		/// </summary>
		public Int32? GetAsInt(String name)
		{
			string attrValue = this[name];
			int val;
			if (attrValue != null && Int32.TryParse(attrValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out val))
				return val;

			return null;
		}

		/// <summary>
		/// Gets an attribute representing a color (named color, hexadecimal or hexadecimal 
		/// without the preceding # character).
		/// </summary>
		public HtmlColor GetAsColor(String name)
		{
			return HtmlColor.Parse(this[name]);
		}

		/// <summary>
		/// Gets an attribute representing an unit: 120px, 10pt, 5em, 20%, ...
		/// </summary>
		/// <returns>If the attribute is misformed, the <see cref="Unit.IsValid"/> property is set to false.</returns>
		public Unit GetAsUnit(String name)
		{
			return Unit.Parse(this[name]);
		}

		/// <summary>
		/// Gets an attribute representing the 4 unit sides.
		/// If a side has been specified individually, it will override the grouped definition.
		/// </summary>
		/// <returns>If the attribute is misformed, the <see cref="Margin.IsValid"/> property is set to false.</returns>
		public Margin GetAsMargin(String name)
		{
			Margin margin = Margin.Parse(this[name]);
			Unit u;

			u = GetAsUnit(name + "-top");
			if (u.IsValid) margin.Top = u;
			u = GetAsUnit(name + "-right");
			if (u.IsValid) margin.Right = u;
			u = GetAsUnit(name + "-bottom");
			if (u.IsValid) margin.Bottom = u;
			u = GetAsUnit(name + "-left");
			if (u.IsValid) margin.Left = u;

			return margin;
		}

        /// <summary>
        /// Gets an attribute representing the 4 border sides.
        /// If a border style/color/width has been specified individually, it will override the grouped definition.
        /// </summary>
        /// <returns>If the attribute is misformed, the <see cref="HtmlBorder.IsEmpty"/> property is set to false.</returns>
        public HtmlBorder GetAsBorder(String name)
		{
			HtmlBorder border = new HtmlBorder(GetAsSideBorder(name));
			SideBorder sb;

			sb = GetAsSideBorder(name + "-top");
			if (sb.IsValid) border.Top = sb;
			sb = GetAsSideBorder(name + "-right");
			if (sb.IsValid) border.Right = sb;
			sb = GetAsSideBorder(name + "-bottom");
			if (sb.IsValid) border.Bottom = sb;
			sb = GetAsSideBorder(name + "-left");
			if (sb.IsValid) border.Left = sb;

			return border;
		}

        /// <summary>
        /// Gets an attribute representing a single border side.
        /// If a border style/color/width has been specified individually, it will override the grouped definition.
        /// </summary>
        /// <returns>If the attribute is misformed, the <see cref="HtmlBorder.IsEmpty"/> property is set to false.</returns>
        public SideBorder GetAsSideBorder(String name)
		{
			string attrValue = this[name];
			SideBorder border = SideBorder.Parse(attrValue);

			// handle attributes specified individually.
			Unit width = SideBorder.ParseWidth(this[name + "-width"]);
			if (width.IsValid) border.Width = width;

			var color = GetAsColor(name + "-color");
			if (!color.IsEmpty) border.Color = color;

			var style = Converter.ToBorderStyle(this[name + "-style"]);
			if (style != w.BorderValues.Nil) border.Style = style;

			return border;
		}

		/// <summary>
		/// Gets the class attribute that specify one or more classnames.
		/// </summary>
		public String[] GetAsClass()
		{
			string attrValue = this["class"];
			if (attrValue == null) return null;
			return attrValue.Split(HttpUtility.WhiteSpaces, StringSplitOptions.RemoveEmptyEntries);
		}

		/// <summary>
		/// Gets the font attribute and combine with the style, size and family.
		/// </summary>
		public HtmlFont GetAsFont(String name)
		{
			HtmlFont font = HtmlFont.Parse(this[name]);
			string attrValue = this[name + "-style"];
			if (attrValue != null)
			{
				var style = Converter.ToFontStyle(attrValue);
				if (style.HasValue) font.Style = style.Value;
			}
			attrValue = this[name + "-variant"];
			if (attrValue != null)
			{
				var variant = Converter.ToFontVariant(attrValue);
				if (variant.HasValue) font.Variant = variant.Value;
			}
			attrValue = this[name + "-weight"];
			if (attrValue != null)
			{
				var weight = Converter.ToFontWeight(attrValue);
				if (weight.HasValue) font.Weight = weight.Value;
			}
			attrValue = this[name + "-family"];
			if (attrValue != null)
			{
				font.Family = Converter.ToFontFamily(attrValue);
			}
			Unit unit = this.GetAsUnit(name + "-size");
			if (unit.IsValid) font.Size = unit;
			return font;
		}
	}
}