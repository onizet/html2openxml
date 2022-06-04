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

namespace HtmlToOpenXml
{
    /// <summary>
    /// Represents a Html font (15px arial,sans-serif).
    /// </summary>
    struct HtmlFont
	{
		/// <summary>Represents an empty font (not defined).</summary>
		public static readonly HtmlFont Empty = new HtmlFont(FontStyle.Normal, FontVariant.Normal, FontWeight.Normal, Unit.Empty, null);

		private FontStyle _style;
		private FontVariant _variant;
		private string _family;
		private FontWeight _weight;
		private Unit _size;


		public HtmlFont(FontStyle style, FontVariant variant, FontWeight weight, Unit size, string family)
		{
			this._style = style;
			this._variant = variant;
			this._family = family;
			this._weight = weight;
			this._size = size;
		}

		public static HtmlFont Parse(String str)
		{
			if (str == null) return HtmlFont.Empty;

			// The font shorthand property sets all the font properties in one declaration.
			// The properties that can be set, are (in order):
			// "font-style font-variant font-weight font-size/line-height font-family"
			// The font-size and font-family values are required.
			// If one of the other values are missing, the default values will be inserted, if any.
			// http://www.w3schools.com/cssref/pr_font_font.asp


			// in order to split by white spaces, we remove any white spaces between 2 family names (ex: Verdana, Arial -> Verdana,Arial)
			str = System.Text.RegularExpressions.Regex.Replace(str, @",\s+?", ",");

			String[] fontParts = str.Split(HttpUtility.WhiteSpaces, StringSplitOptions.RemoveEmptyEntries);
			if (fontParts.Length < 2) return HtmlFont.Empty;
			HtmlFont font = HtmlFont.Empty;

			if (fontParts.Length == 2) // 2=the minimal set of required parameters
			{
				// should be the size and the family (in that order). Others are set to their default values
				font._size = ReadFontSize(fontParts[0]);
				if (!font._size.IsValid) return HtmlFont.Empty;
				font._family = Converter.ToFontFamily(fontParts[1]);
				return font;
			}

			int index = 0;

			FontStyle? style = Converter.ToFontStyle(fontParts[index]);
			if (style.HasValue) { font._style = style.Value; index++; }

			if (index + 2 > fontParts.Length) return HtmlFont.Empty;
			FontVariant? variant = Converter.ToFontVariant(fontParts[index]);
			if (variant.HasValue) { font._variant = variant.Value; index++; }

			if (index + 2 > fontParts.Length) return HtmlFont.Empty;
			FontWeight? weight = Converter.ToFontWeight(fontParts[index]);
			if (weight.HasValue) { font._weight = weight.Value; index++; }

			if (fontParts.Length - index < 2) return HtmlFont.Empty;
			font._size = ReadFontSize(fontParts[fontParts.Length - 2]);
			if (!font._size.IsValid) return HtmlFont.Empty;

			font._family = Converter.ToFontFamily(fontParts[fontParts.Length - 1]);

			return font;
		}

		private static Unit ReadFontSize(string str)
		{
			Unit size = Converter.ToFontSize(str);
			return size; // % and ratio font-size/line-height are not supported
		}

		//____________________________________________________________________
		//

		/// <summary>
		/// Gets or sets the name of this font.
		/// </summary>
		public string Family
		{
			get { return _family; }
			set { _family = value; }
		}

		/// <summary>
		/// Gest or sets the style for the text.
		/// </summary>
		public FontStyle Style
		{
			get { return _style; }
			set { _style = value; }
		}

		/// <summary>
		/// Gets or sets the variation of the characters.
		/// </summary>
		public FontVariant Variant
		{
			get { return _variant; }
			set { _variant = value; }
		}

		/// <summary>
		/// Gets or sets the size of the font, expressed in half points.
		/// </summary>
		public Unit Size
		{
			get { return _size; }
			set { _size = value; }
		}

		/// <summary>
		/// Gets or sets the weight of the characters (thin or thick).
		/// </summary>
		public FontWeight Weight
		{
			get { return _weight; }
			set { _weight = value; }
		}

		/// <summary>
		/// Gets whether the border is well formed and not empty.
		/// </summary>
		public bool IsEmpty
		{
			get { return _family == null && !_size.IsValid && _weight == FontWeight.Normal && _style == FontStyle.Normal && _variant == FontVariant.Normal; }
		}
	}
}