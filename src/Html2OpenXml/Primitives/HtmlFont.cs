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
using System.Linq;

namespace HtmlToOpenXml;

/// <summary>
/// Represents a Html font (15px arial,sans-serif).
/// </summary>
readonly struct HtmlFont(Unit size, string? family, FontStyle? style, FontVariant? variant, FontWeight? weight)
{
    /// <summary>Represents an empty font (not defined).</summary>
    public static readonly HtmlFont Empty = new ();

    private readonly FontStyle? style = style;
    private readonly FontVariant? variant = variant;
    private readonly string? family = family;
    private readonly FontWeight? weight = weight;
    private readonly Unit size = size;

    /// <inheritdoc cref="Parse(ReadOnlySpan{char})"/>
    public static HtmlFont Parse(string? str)
    {
        if (str == null)
            return Empty;
        return Parse(str.AsSpan());
    }

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

        // in order to split by white spaces, we remove any white spaces between 2 family names (ex: Verdana, Arial -> Verdana,Arial)
        //str = System.Text.RegularExpressions.Regex.Replace(str, @",\s+?", ",");

        Span<Range> tokens = stackalloc Range[6];
        var tokenCount = span.SplitCompositeAttribute(tokens, ' ', skipSeparatorIfPrecededBy: ',');
        if (tokenCount == 0)
            return Empty;

        // Initialize default values
        FontStyle? style = null;
        FontVariant? variant = null;
        FontWeight? weight = null;
        // % and ratio font-size/line-height are not supported
        Unit fontSize = Unit.Empty;
        string? family = null;

        if (tokenCount == 2) // 2=the minimal set of required parameters
        {
            // should be the size and the family (in that order). Others are set to their default values
            fontSize = Converter.ToFontSize(span.Slice(tokens[0]));
            if (!fontSize.IsValid) fontSize = Unit.Empty;
            family = Converter.ToFontFamily(span.Slice(tokens[1]));
            return new HtmlFont(fontSize, family, style, variant, weight);
        }

        // Now try to guess the values with their permutation
        var tokenIndexes = new List<int>(Enumerable.Range(0, tokenCount));

        // handle border style
        for (int i = 0; i < tokenIndexes.Count; i++)
        {
            style = Converter.ToFontStyle(span.Slice(tokens[tokenIndexes[i]]));
            if (style != null)
            {
                tokenIndexes.RemoveAt(i); // no need to process this part anymore
                break;
            }
        }

        for (int i = 0; i < tokenIndexes.Count; i++)
        {
            variant = Converter.ToFontVariant(span.Slice(tokens[tokenIndexes[i]]));
            if (variant != null)
            {
                tokenIndexes.RemoveAt(i); // no need to process this part anymore
                break;
            }
        }

        for (int i = 0; i < tokenIndexes.Count; i++)
        {
            weight = Converter.ToFontWeight(span.Slice(tokens[tokenIndexes[i]]));
            if (weight != null)
            {
                tokenIndexes.RemoveAt(i); // no need to process this part anymore
                break;
            }
        }

        for (int i = 0; i < tokenIndexes.Count; i++)
        {
            fontSize = Unit.Parse(span.Slice(tokens[tokenIndexes[i]]));
            if (fontSize.IsValid)
            {
                tokenIndexes.RemoveAt(i); // no need to process this part anymore
                break;
            }
        }
        if (!fontSize.IsValid) fontSize = Unit.Empty;

        // keep font family as the latest because it is the most permissive
        if(tokenIndexes.Count > 0)
            family = Converter.ToFontFamily(span.Slice(tokens[tokenIndexes[0]]));

        return new HtmlFont(fontSize, family, style, variant, weight);
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
