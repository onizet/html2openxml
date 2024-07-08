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
readonly struct HtmlFont(FontStyle? style, FontVariant? variant, FontWeight? weight, Unit? size, string? family)
{
    /// <summary>Represents an empty font (not defined).</summary>
    public static readonly HtmlFont Empty = new ();

    private readonly FontStyle? style = style;
    private readonly FontVariant? variant = variant;
    private readonly string? family = family;
    private readonly FontWeight? weight = weight;
    private readonly Unit size = size ?? Unit.Empty;

    public static HtmlFont Parse(string? str)
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

        var fontParts = str.Split(HttpUtility.WhiteSpaces, StringSplitOptions.RemoveEmptyEntries);
        if (fontParts.Length < 2) return HtmlFont.Empty;
        
        FontStyle? style = null;
        FontVariant? variant = null;
        FontWeight? weight = null;
        // % and ratio font-size/line-height are not supported
        Unit fontSize;
        string? family;

        if (fontParts.Length == 2) // 2=the minimal set of required parameters
        {
            // should be the size and the family (in that order). Others are set to their default values
            fontSize = Converter.ToFontSize(fontParts[0]);
            if (!fontSize.IsValid) fontSize = Unit.Empty;
            family = Converter.ToFontFamily(fontParts[1]);
            return new HtmlFont(style, variant, weight, fontSize, family);
        }

        int index = 0;

        style = Converter.ToFontStyle(fontParts[index]);
        if (style.HasValue) { index++; }

        if (index + 2 > fontParts.Length) return HtmlFont.Empty;
        variant = Converter.ToFontVariant(fontParts[index]);
        if (variant.HasValue) { index++; }

        if (index + 2 > fontParts.Length) return HtmlFont.Empty;
        weight = Converter.ToFontWeight(fontParts[index]);
        if (weight.HasValue) { index++; }

        if (fontParts.Length - index < 2) return HtmlFont.Empty;
        fontSize = Converter.ToFontSize(fontParts[fontParts.Length - 2]);
        if (!fontSize.IsValid) return HtmlFont.Empty;

        family = Converter.ToFontFamily(fontParts[fontParts.Length - 1]);

        return new HtmlFont(style, variant, weight, fontSize, family);
    }

    //____________________________________________________________________
    //

    /// <summary>
    /// Gets or sets the name of this font.
    /// </summary>
    public string? Family
    {
        get { return family; }
    }

    /// <summary>
    /// Gest or sets the style for the text.
    /// </summary>
    public FontStyle? Style
    {
        get { return style; }
    }

    /// <summary>
    /// Gets or sets the variation of the characters.
    /// </summary>
    public FontVariant? Variant
    {
        get { return variant; }
    }

    /// <summary>
    /// Gets or sets the size of the font, expressed in half points.
    /// </summary>
    public Unit Size
    {
        get { return size; }
    }

    /// <summary>
    /// Gets or sets the weight of the characters (thin or thick).
    /// </summary>
    public FontWeight? Weight
    {
        get { return weight; }
    }
}
