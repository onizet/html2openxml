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

namespace HtmlToOpenXml;

/// <summary>
/// Represents a Html font (15px arial,sans-serif).
/// </summary>
readonly struct HtmlFont(Unit size, string? family, FontStyle? style,
    FontVariant? variant, FontWeight? weight, Unit lineHeight)
{
    /// <summary>Represents an empty font (not defined).</summary>
    public static readonly HtmlFont Empty = new ();

    private readonly FontStyle? style = style;
    private readonly FontVariant? variant = variant;
    private readonly string? family = family;
    private readonly FontWeight? weight = weight;
    private readonly Unit size = size;
    private readonly Unit lineHeight = lineHeight;


    /// <summary>
    /// Parse the font style attribute.
    /// </summary>
    /// <remarks>
    /// The font shorthand property sets all the font properties in one declaration.
    /// The properties that can be set, are (in order):
    /// "font-style font-variant font-weight font-size/line-height font-family"
    /// The font-size and font-family values are required.
    /// If one of the other values are missing, the default values will be inserted, if any.
    /// /// </remarks>
    public static HtmlFont Parse(ReadOnlySpan<char> span)
    {
        // http://www.w3schools.com/cssref/pr_font_font.asp

        if (span.IsEmpty || span.Length < 2) return Empty;

        Span<Range> tokens = stackalloc Range[6];
        var tokenCount = span.SplitCompositeAttribute(tokens, ' ', skipSeparatorIfPrecededBy: ',');
        if (tokenCount == 0)
            return Empty;

        // Initialize default values
        FontStyle? style = null;
        FontVariant? variant = null;
        FontWeight? weight = null;
        // % and ratio font-size/line-height are not supported
        Unit fontSize = Unit.Empty, lineHeight = Unit.Empty;
        string? fontFamily = null;

        if (tokenCount == 2) // 2=the minimal set of required parameters
        {
            // should be the size and the family (in that order). Others are set to their default values
            fontSize = Converter.ToFontSize(span.Slice(tokens[0]));
            if (!fontSize.IsValid) return Empty;
            fontFamily = Converter.ToFontFamily(span.Slice(tokens[1]));
            return new HtmlFont(fontSize, fontFamily, style, variant, weight, lineHeight);
        }
        else if (tokenCount > 10)
        {
            // safety check to avoid overflow with stackalloc in a loop
            return Empty;
        }

        Span<char> loweredValue = stackalloc char[128];
        for (int i = 0; i < tokenCount; i++)
        {
            var token = span.Slice(tokens[i]).Trim();
            token.ToLowerInvariant(loweredValue);

            switch (loweredValue.Slice(0, token.Length))
            {
                case "italic" or "oblique": style = FontStyle.Italic; break;
                case "normal":
                    style ??= FontStyle.Normal;
                    variant ??= FontVariant.Normal; 
                    weight ??= FontWeight.Normal;
                    break;
                case "small-caps": variant = FontVariant.SmallCaps; break;
                case "700" or "bold": weight = FontWeight.Bold; break;
                case "bolder": weight = FontWeight.Bolder; break;
                case "400": weight = FontWeight.Normal; break;
                case "xx-small": fontSize = new Unit(UnitMetric.Point, 10); break;
                case "x-small": fontSize = new Unit(UnitMetric.Point, 15); break;
                case "small": fontSize = new Unit(UnitMetric.Point, 20); break;
                case "medium": fontSize = new Unit(UnitMetric.Point, 27); break;
                case "large": fontSize = new Unit(UnitMetric.Point, 36); break;
                case "x-large": fontSize = new Unit(UnitMetric.Point, 48); break;
                case "xx-large": fontSize = new Unit(UnitMetric.Point, 72); break;
                default:
                {
                    if (fontSize.IsValid || !TryParseFontSize (token, out fontSize, out lineHeight))
                    {
                        fontFamily ??= Converter.ToFontFamily(token);
                    }

                    break;
                }
            }
        }

        return new HtmlFont(fontSize, fontFamily, style, variant, weight, lineHeight);
    }

    private static bool TryParseFontSize(ReadOnlySpan<char> token, out Unit fontSize, out Unit lineHeight)
    {
        // Handle font-size/line-height
        var slash = token.IndexOf('/');
        if (slash > 0)
        {
            fontSize = Unit.Parse(token.Slice(0, slash));
            lineHeight = Unit.Parse(token.Slice(slash + 1));
            return fontSize.IsValid;
        }

        fontSize = Unit.Parse(token);
        lineHeight = Unit.Empty;
        return fontSize.IsValid;
    }

    //____________________________________________________________________
    //

    /// <summary>
    /// Gets the name of this font.
    /// </summary>
    public string? Family
    {
        get { return family; }
    }

    /// <summary>
    /// Gest the style for the text.
    /// </summary>
    public FontStyle? Style
    {
        get { return style; }
    }

    /// <summary>
    /// Gets the variation of the characters.
    /// </summary>
    public FontVariant? Variant
    {
        get { return variant; }
    }

    /// <summary>
    /// Gets the size of the font, expressed in half points.
    /// </summary>
    public Unit Size
    {
        get { return size; }
    }

    /// <summary>
    /// Gets the weight of the characters (thin or thick).
    /// </summary>
    public FontWeight? Weight
    {
        get { return weight; }
    }

    /// <summary>
    /// Gets the height of a line.
    /// </summary>
    public Unit LineHeight
    {
        get { return lineHeight; }
    }
}
