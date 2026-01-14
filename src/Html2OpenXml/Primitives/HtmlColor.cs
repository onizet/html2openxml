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
using System.Globalization;

namespace HtmlToOpenXml;

/// <summary>
/// Represents an ARGB color.
/// </summary>
readonly partial struct HtmlColor : IEquatable<HtmlColor>
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


    public HtmlColor(double alpha, byte red, byte green, byte blue) : this()
    {
        A = alpha;
        R = red;
        G = green;
        B = blue;
    }

    /// <summary>
    /// Try to parse a value (RGB(A) or HSL(A), hexadecimal, or named color) to its RGB representation.
    /// </summary>
    /// <param name="span">The color to parse.</param>
    /// <returns>Returns <see cref="HtmlColor.Empty"/> if parsing failed.</returns>
    public static HtmlColor Parse(ReadOnlySpan<char> span)
    {
        span = span.Trim();
        if (span.Length < 3)
            return Empty;
 
        try
        {
            // Is it in hexa? Note: we no more accept hexa value without preceding the '#'
            if (span[0] == '#')
            {
                return ParseHexa(span);
            }

            // RGB or RGBA
            if (span.StartsWith(['r','g','b'], StringComparison.OrdinalIgnoreCase))
            {
                return ParseRgb(span);
            }

            // HSL or HSLA
            if (span.StartsWith(['h','s','l'], StringComparison.OrdinalIgnoreCase))
            {
                return ParseHsl(span);
            }
        }
        catch (Exception exc)
        {
            if (exc is FormatException || exc is OverflowException || exc is ArgumentOutOfRangeException)
                return Empty;
            throw;
        }

        return GetNamedColor(span);
    }

    private static HtmlColor ParseHexa(ReadOnlySpan<char> span)
    {
        if (span.Length == 7)
        {
            return FromArgb(
                span.Slice(1, 2).AsByte(NumberStyles.HexNumber),
                span.Slice(3, 2).AsByte(NumberStyles.HexNumber),
                span.Slice(5, 2).AsByte(NumberStyles.HexNumber));
        }
        if (span.Length == 4)
        {
            // #0FF --> #00FFFF
            ReadOnlySpan<char> r = [span[1], span[1]];
            ReadOnlySpan<char> g = [span[2], span[2]];
            ReadOnlySpan<char> b = [span[3], span[3]];
            return FromArgb(
                    r.AsByte(NumberStyles.HexNumber),
                    g.AsByte(NumberStyles.HexNumber),
                    b.AsByte(NumberStyles.HexNumber));
        }
        return Empty;
    }

    private static HtmlColor ParseRgb(ReadOnlySpan<char> span)
    {
        int startIndex = span.IndexOf('('), endIndex = span.LastIndexOf(')');
        if (startIndex < 3 || endIndex == -1)
            return Empty;

        span = span.Slice(startIndex + 1, endIndex - startIndex - 1);
        Span<Range> tokens = stackalloc Range[5];
        var sep = span.IndexOf(',') > -1? ',' : ' ';
        return span.Split(tokens, sep, StringSplitOptions.RemoveEmptyEntries) switch
        {
            3 => FromArgb(1.0,
                span.Slice(tokens[0]).AsByte(NumberStyles.Integer),
                span.Slice(tokens[1]).AsByte(NumberStyles.Integer),
                span.Slice(tokens[2]).AsByte(NumberStyles.Integer)),
            4 => FromArgb(span.Slice(tokens[3]).AsDouble(),
                span.Slice(tokens[0]).AsByte(NumberStyles.Integer),
                span.Slice(tokens[1]).AsByte(NumberStyles.Integer),
                span.Slice(tokens[2]).AsByte(NumberStyles.Integer)),
            // r g b / a
            5 => FromArgb(span.Slice(tokens[4]).AsDouble(),
                span.Slice(tokens[0]).AsByte(NumberStyles.Integer),
                span.Slice(tokens[1]).AsByte(NumberStyles.Integer),
                span.Slice(tokens[2]).AsByte(NumberStyles.Integer)),
            _ => Empty
        };
    }

    private static HtmlColor ParseHsl(ReadOnlySpan<char> span)
    {
        int startIndex = span.IndexOf('('), endIndex = span.LastIndexOf(')');
        if (startIndex < 3 || endIndex == -1)
            return Empty;

        span = span.Slice(startIndex + 1, endIndex - startIndex - 1);
        Span<Range> tokens = stackalloc Range[5];
        var sep = span.IndexOf(',') > -1? ',' : ' ';
        return span.Split(tokens, sep, StringSplitOptions.RemoveEmptyEntries) switch
        {
            3 => FromHsl(1.0,
                span.Slice(tokens[0]).AsDouble(),
                span.Slice(tokens[1]).AsPercent(),
                span.Slice(tokens[2]).AsPercent()),
            4 => FromHsl(span.Slice(tokens[3]).AsDouble(),
                span.Slice(tokens[0]).AsDouble(),
                span.Slice(tokens[1]).AsPercent(),
                span.Slice(tokens[2]).AsPercent()),
            _ => Empty
        };
    }

    /// <summary>
    /// Convert a potential percentage value to its numeric representation.
    /// Saturation and Lightness can contains both a percentage value or a value comprised between 0.0 and 1.0. 
    /// </summary>
    private static double ParsePercent (string value)
    {
        double parsedValue;
        if (value.IndexOf('%') > -1)
            parsedValue = double.Parse(value.Replace('%', ' '), CultureInfo.InvariantCulture) / 100d;
        else
            parsedValue = double.Parse(value, CultureInfo.InvariantCulture);

        return Math.Min(1, Math.Max(0, parsedValue));
    }

    /// <summary>
    /// Creates a <see cref="HtmlColor"/> structure from the four RGB component values.
    /// </summary>
    /// <param name="red">The red component.</param>
    /// <param name="green">The green component.</param>
    /// <param name="blue">The blue component.</param>
    public static HtmlColor FromArgb(byte red, byte green, byte blue)
    {
        return FromArgb(1d, red, green, blue);
    }

    /// <summary>
    /// Creates a <see cref="HtmlColor"/> structure from the four ARGB component values.
    /// </summary>
    /// <param name="alpha">The alpha component (0.0-1.0).</param>
    /// <param name="red">The red component (0-255).</param>
    /// <param name="green">The green component (0-255).</param>
    /// <param name="blue">The blue component (0-255).</param>
    public static HtmlColor FromArgb(double alpha, byte red, byte green, byte blue)
    {
        if (alpha < 0.0 || alpha > 1.0)
            throw new ArgumentOutOfRangeException(nameof(alpha), alpha, "Alpha should be comprised between 0.0 and 1.0");

        return new HtmlColor(alpha, red, green, blue);
    }

    /// <summary>
    /// Convert a color using the HSL to RGB.
    /// </summary>
    /// <param name="alpha">The alpha component (0.0-1.0).</param>
    /// <param name="hue">The Hue component (0.0 - 360.0).</param>
    /// <param name="saturation">The saturation component (0.0 - 1.0).</param>
    /// <param name="luminosity">The luminosity component (0.0 - 1.0).</param>
    public static HtmlColor FromHsl(double alpha, double hue, double saturation, double luminosity)
    {
        if (alpha < 0.0 || alpha > 1.0)
            throw new ArgumentOutOfRangeException(nameof(alpha), alpha, "Alpha should be comprised between 0.0 and 1.0");

        if (hue < 0 || hue > 360)
            throw new ArgumentOutOfRangeException(nameof(hue), hue, "Hue should be comprised between 0° and 360°");

        if (saturation < 0 || saturation > 1)
            throw new ArgumentOutOfRangeException(nameof(saturation), saturation, "Saturation should be comprised between 0.0 and 1.0");

        if (luminosity < 0 || luminosity > 1)
            throw new ArgumentOutOfRangeException(nameof(luminosity), luminosity, "Brightness should be comprised between 0.0 and 1.0");

        if (0 == saturation)
        {
            return FromArgb(alpha, Convert.ToByte(luminosity * 255),
                Convert.ToByte(luminosity * 255), Convert.ToByte(luminosity * 255));
        }

        double fMax, fMid, fMin;
        int iSextant;

        if (0.5 < luminosity)
        {
            fMax = luminosity - (luminosity * saturation) + saturation;
            fMin = luminosity + (luminosity * saturation) - saturation;
        }
        else
        {
            fMax = luminosity + (luminosity * saturation);
            fMin = luminosity - (luminosity * saturation);
        }

        iSextant = (int) Math.Floor(hue / 60f);
        if (300f <= hue)
        {
            hue -= 360f;
        }
        hue /= 60f;
        hue -= 2f * (float) Math.Floor((iSextant + 1f) % 6f / 2f);
        if (0 == iSextant % 2)
        {
            fMid = hue * (fMax - fMin) + fMin;
        }
        else
        {
            fMid = fMin - hue * (fMax - fMin);
        }

        byte iMax = Convert.ToByte(fMax * 255);
        byte iMid = Convert.ToByte(fMid * 255);
        byte iMin = Convert.ToByte(fMin * 255);

        return iSextant switch
        {
            1 => FromArgb(alpha, iMid, iMax, iMin),
            2 => FromArgb(alpha, iMin, iMax, iMid),
            3 => FromArgb(alpha, iMin, iMid, iMax),
            4 => FromArgb(alpha, iMid, iMin, iMax),
            5 => FromArgb(alpha, iMax, iMin, iMid),
            _ => FromArgb(alpha, iMax, iMid, iMin),
        };
    }

    /// <summary>
    /// Tests whether the specified object is a HtmlColor structure and is equivalent to this color structure.
    /// </summary>
    public bool Equals(HtmlColor color)
    {
        return color.A == A && color.R == R && color.G == G && color.B == B;
    }

    /// <summary>
    /// Convert a .Net Color to a hex string.
    /// </summary>
    public string ToHexString()
    {
        // http://www.cambiaresearch.com/c4/24c09e15-2941-4ad2-8695-00b1b4029f4d/Convert-dotnet-Color-to-Hex-String.aspx

        byte[] bytes = [R, G, B];
        char[] chars = new char[bytes.Length * 2];
        for (int i = 0; i < bytes.Length; i++)
        {
            int b = bytes[i];
            chars[i * 2] = hexDigits[b >> 4];
            chars[i * 2 + 1] = hexDigits[b & 0xF];
        }
        return new string(chars);
    }

    /// <summary>
    /// Gets a representation of this color expressed in ARGB.
    /// </summary>
    public override string ToString()
    {
        return string.Format("A: {0:#0.##} R: {1:#0##} G: {2:#0##} B: {3:#0##}", A, R, G, B);
    }

    //____________________________________________________________________
    //

    /// <summary>Gets the alpha component value of this color structure.</summary>
    public double A { get; }
    /// <summary>Gets the red component value of this cColor structure.</summary>
    public byte R { get; }
    /// <summary>Gets the green component value of this color structure.</summary>
    public byte G { get; }
    /// <summary>Gets the blue component value of this color structure.</summary>
    public byte B { get; }

    /// <summary>
    /// Specifies whether this HtmlColor structure is uninitialized.
    /// </summary>
    public bool IsEmpty { get => this.Equals(Empty); }
}
