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

namespace HtmlToOpenXml
{
    /// <summary>
    /// Represents an ARGB color.
    /// </summary>
    struct HtmlColor
    {
        private static readonly char[] hexDigits = {
         '0', '1', '2', '3', '4', '5', '6', '7',
         '8', '9', 'A', 'B', 'C', 'D', 'E', 'F'};

        /// <summary>
        /// Represents a color that is null.
        /// </summary>
        public static readonly HtmlColor Empty = new HtmlColor();
        /// <summary>
        /// Gets a system-defined color that has an ARGB value of #FF000000.
        /// </summary>
        public static readonly HtmlColor Black = FromArgb(0, 0, 0);


        public static HtmlColor Parse(string htmlColor)
        {
            HtmlColor color = new HtmlColor();

            // Bug fixed by jairoXXX to support rgb(r,g,b) format
            if (htmlColor.StartsWith("rgb", StringComparison.OrdinalIgnoreCase))
            {
                int startIndex = htmlColor.IndexOf('(', 3), endIndex = htmlColor.LastIndexOf(')');
                if (startIndex >= 3 && endIndex > -1)
                {
                    var colorStringArray = htmlColor.Substring(startIndex + 1, endIndex - startIndex - 1).Split(',');
                    if (colorStringArray.Length >= 3)
                    {
                        color = FromArgb(
                            colorStringArray.Length == 3 ? (byte)255 : Byte.Parse(colorStringArray[0], NumberStyles.Integer, CultureInfo.InvariantCulture),
                            Byte.Parse(colorStringArray[colorStringArray.Length - 3], NumberStyles.Integer, CultureInfo.InvariantCulture),
                            Byte.Parse(colorStringArray[colorStringArray.Length - 2], NumberStyles.Integer, CultureInfo.InvariantCulture),
                            Byte.Parse(colorStringArray[colorStringArray.Length - 1], NumberStyles.Integer, CultureInfo.InvariantCulture)
                        );
                        return color;
                    }
                }
            }

            try
            {
                // The Html allows to write color in hexa without the preceding '#'
                // I just ensure it's a correct hexadecimal value (length=6 and first character should be
                // a digit or an hexa letter)
                if (htmlColor.Length == 6 && (Char.IsDigit(htmlColor[0]) || (htmlColor[0] >= 'a' && htmlColor[0] <= 'f')
                    || (htmlColor[0] >= 'A' && htmlColor[0] <= 'F')))
                {
                    try
                    {
                        color = FromArgb(255,
                            Convert.ToByte(htmlColor.Substring(0, 2), 16),
                            Convert.ToByte(htmlColor.Substring(2, 2), 16),
                            Convert.ToByte(htmlColor.Substring(4, 2), 16));
                    }
                    catch (System.FormatException)
                    {
                        // If the conversion failed, that should be a named color
                        // Let the framework dealing with it
                        color = HtmlColor.Empty;
                    }
                }

                if (color.IsEmpty)
                {
#if FEATURE_DRAWING
                    var nativeColor = System.Drawing.ColorTranslator.FromHtml(htmlColor);
                    color = FromArgb(nativeColor.A, nativeColor.B, nativeColor.G, nativeColor.B);
#else
                    color = HtmlColorTranslator.FromHtml(htmlColor);
#endif
                }
            }
            catch (Exception exc)
            {
                if (exc.InnerException is System.FormatException)
                    return HtmlColor.Empty;
                throw;
            }

            return color;
        }

        /// <summary>
        /// Creates a <see cref="HtmlColor"/> structure from the four RGB component values.
        /// </summary>
        /// <param name="red">The red component.</param>
        /// <param name="green">The green component.</param>
        /// <param name="blue">The blue component.</param>
        public static HtmlColor FromArgb(byte red, byte green, byte blue)
        {
            return FromArgb(255, red, green, blue);
        }

        /// <summary>
        /// Creates a <see cref="HtmlColor"/> structure from the four ARGB component values.
        /// </summary>
        /// <param name="alpha">The alpha component.</param>
        /// <param name="red">The red component.</param>
        /// <param name="green">The green component.</param>
        /// <param name="blue">The blue component.</param>
        public static HtmlColor FromArgb(byte alpha, byte red, byte green, byte blue)
        {
            return new HtmlColor() {
                A = alpha, R = red, G = green, B = blue
            };
        }

        /// <summary>
        /// Tests whether the specified object is a HtmlColor structure and is equivalent to this color structure.
        /// </summary>
        public bool Equals(HtmlColor color)
        {
            return color.A == A && color.R == A && color.G == A && color.B == A;
        }

        /// <summary>
        /// Convert a .Net Color to a hex string.
        /// </summary>
        public string ToHexString()
        {
            // http://www.cambiaresearch.com/c4/24c09e15-2941-4ad2-8695-00b1b4029f4d/Convert-dotnet-Color-to-Hex-String.aspx

            byte[] bytes = new byte[3];
            bytes[0] = this.R;
            bytes[1] = this.G;
            bytes[2] = this.B;
            char[] chars = new char[bytes.Length * 2];
            for (int i = 0; i < bytes.Length; i++)
            {
                int b = bytes[i];
                chars[i * 2] = hexDigits[b >> 4];
                chars[i * 2 + 1] = hexDigits[b & 0xF];
            }
            return new string(chars);
        }

        //____________________________________________________________________
        //

        /// <summary>Gets the alpha component value of this color structure.</summary>
        public byte A { get; set; }
        /// <summary>Gets the red component value of this cColor structure.</summary>
        public byte R { get; set; }
        /// <summary>Gets the green component value of this color structure.</summary>
        public byte G { get; set; }
        /// <summary>Gets the blue component value of this color structure.</summary>
        public byte B { get; set; }

        /// <summary>
        /// Specifies whether this HtmlColor structure is uninitialized.
        /// </summary>
        public bool IsEmpty { get { return this.Equals(Empty); } }
    }
}