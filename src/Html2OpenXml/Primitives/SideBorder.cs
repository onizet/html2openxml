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
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml
{
    /// <summary>
    /// Represents a Html Unit (ie: 120px, 10em, ...).
    /// </summary>
    readonly struct SideBorder(BorderValues style, HtmlColor color, Unit size)
    {
        /// <summary>Represents an empty border (not defined).</summary>
        public static readonly SideBorder Empty = new(BorderValues.Nil, HtmlColor.Empty, Unit.Empty);

        private readonly BorderValues style = style;
        private readonly HtmlColor color = color;
        private readonly Unit size = size;

        public static SideBorder Parse(string? str)
        {
            if (str == null) return SideBorder.Empty;

            // The properties of a border that can be set, are (in order): border-width, border-style, and border-color.
            // It does not matter if one of the values above are missing, e.g. border:solid #ff0000; is allowed.
            // The main problem for parsing this attribute is that the browsers allow any permutation of the values... meaning more coding :(
            // http://www.w3schools.com/cssref/pr_border.asp

            // Remove the spaces that could appear in the color parameter: rgb(233, 233, 233) -> rgb(233,233,233)
            str = Regex.Replace(str, @",\s+?", ",");
            var borderParts = new List<string>(str.Split(HttpUtility.WhiteSpaces, StringSplitOptions.RemoveEmptyEntries));
            if (borderParts.Count == 0) return SideBorder.Empty;

            // Initialize default values
            Unit borderWidth = Unit.Empty;
            HtmlColor borderColor = HtmlColor.Empty;
            BorderValues borderStyle = BorderValues.Nil;

            // Now try to guess the values with their permutation

            // handle border style
            for (int i = 0; i < borderParts.Count; i++)
            {
                borderStyle = Converter.ToBorderStyle(borderParts[i]);
                if (borderStyle != BorderValues.Nil)
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
                borderColor = HtmlColor.Parse(borderParts[0]);

            if (borderColor.IsEmpty && !borderWidth.IsValid && borderStyle == BorderValues.Nil)
                return SideBorder.Empty;

            // returns the instance with default value if needed.
            // These value are the ones used by the browser, i.e: solid 3px black
            return new SideBorder(
                borderStyle == BorderValues.Nil? BorderValues.Single : borderStyle,
                borderColor.IsEmpty? HtmlColor.Black : borderColor,
                borderWidth.IsFixed? borderWidth : new Unit(UnitMetric.Pixel, 4));
        }

        internal static Unit ParseWidth(string? borderWidth)
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
        public BorderValues Style
        {
            get { return style; }
        }

        /// <summary>
        /// Gets or sets the color of the border.
        /// </summary>
        public HtmlColor Color
        {
            get { return color; }
        }

        /// <summary>
        /// Gets or sets the size of the border expressed with its unit.
        /// </summary>
        public Unit Width
        {
            get { return size; }
        }

        /// <summary>
        /// Gets whether the border is well formed and not empty.
        /// </summary>
        public bool IsValid
        {
            get { return !BorderValues.Nil.Equals(Style); }
        }
    }
}