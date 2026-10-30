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

        if (span.IsEmpty || span.Length < 2)
            return Empty;

        Span<Range> tokens = stackalloc Range[10];
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

        // fast path: 2=the minimal set of required parameters, e.g.: '12px Arial'
        if (tokenCount == 2)
        {
            // should be the size and the family (in that order). Others are set to their default values
            if (!Converter.TryParseFontSize(span[tokens[0]], out fontSize))
                return Empty;

            _ = Converter.TryParseFontFamily(span[tokens[1]], out fontFamily);
            return new HtmlFont(fontSize, fontFamily, style, variant, weight, lineHeight);
        }

        for (int i = 0; i < tokenCount; i++)
        {
            var token = span[tokens[i]].Trim();

            // normal could be used for multiple cases
            if (token.Is("normal"))
            {
                style ??= FontStyle.Normal;
                variant ??= FontVariant.Normal;
                weight ??= FontWeight.Normal;
                continue;
            }

            if (!weight.HasValue && Converter.TryParseFontWeight(token, out weight))
            {
                continue;
            }
            if (!style.HasValue && Converter.TryParseFontStyle(token, out style))
            {
                continue;
            }
            if (!variant.HasValue && Converter.TryParseFontVariant(token, out variant))
            {
                continue;
            }
            if (!fontSize.IsValid && TryParseFontSize(token, out fontSize, out lineHeight))
            {
                continue;
            }
            else if (fontFamily is null)
            {
                Converter.TryParseFontFamily(token, out fontFamily);
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

        lineHeight = Unit.Empty;
        return Converter.TryParseFontSize(token, out fontSize);
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
