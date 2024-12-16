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
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml;

/// <summary>
/// Represents a Html border (ie: 1.2px solid blue...).
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
        if (str == null) return Empty;
        return Parse(str.AsSpan());
    }

    public static SideBorder Parse(ReadOnlySpan<char> span)
    {
        // The properties of a border that can be set, are (in order): border-width, border-style, and border-color.
        // It does not matter if one of the values above are missing, e.g. border:solid #ff0000; is allowed.
        // The main problem for parsing this attribute is that the browsers allow any permutation of the values... meaning more coding :(
        // http://www.w3schools.com/cssref/pr_border.asp

        if (span.Length < 2)
            return Empty;

        Span<Range> tokens = stackalloc Range[6];
        var tokenCount = span.SplitCompositeAttribute(tokens);
        if (tokenCount == 0)
            return Empty;

        // Initialize default values
        Unit borderWidth = Unit.Empty;
        HtmlColor borderColor = HtmlColor.Empty;
        BorderValues borderStyle = BorderValues.Nil;

        // Now try to guess the values with their permutation
        var tokenIndexes = new List<int>(Enumerable.Range(0, tokenCount));

        // handle border style
        for (int i = 0; i < tokenIndexes.Count; i++)
        {
            borderStyle = Converter.ToBorderStyle(span.Slice(tokens[tokenIndexes[i]]));
            if (borderStyle != BorderValues.Nil)
            {
                tokenIndexes.RemoveAt(i); // no need to process this part anymore
                break;
            }
        }

        for (int i = 0; i < tokenIndexes.Count; i++)
        {
            borderWidth = ParseWidth(span.Slice(tokens[tokenIndexes[i]]));
            if (borderWidth.IsValid)
            {
                tokenIndexes.RemoveAt(i); // no need to process this part anymore
                break;
            }
        }

        // find width
        if(tokenIndexes.Count > 0)
            borderColor = HtmlColor.Parse(span.Slice(tokens[tokenIndexes[0]]));

        if (borderColor.IsEmpty && !borderWidth.IsValid && borderStyle == BorderValues.Nil)
            return Empty;

        // returns the instance with default value if needed.
        // These value are the ones used by the browser, i.e: solid 3px black
        return new SideBorder(
            borderStyle == BorderValues.Nil? BorderValues.Single : borderStyle,
            borderColor.IsEmpty? HtmlColor.Black : borderColor,
            borderWidth.IsFixed? borderWidth : new Unit(UnitMetric.Pixel, 4));
    }

    internal static Unit ParseWidth(ReadOnlySpan<char> borderWidth)
    {
        Unit bu = Unit.Parse(borderWidth, UnitMetric.Pixel);
        if (bu.IsValid)
        {
            if (bu.Value > 0 && bu.Metric == UnitMetric.Pixel)
                return bu;
            return Unit.Empty;
        }
        else
        {
            Span<char> loweredValue = borderWidth.Length <= 128 ? stackalloc char[borderWidth.Length] : new char[borderWidth.Length];
            borderWidth.ToLowerInvariant(loweredValue);

            return loweredValue switch {
               "thin" => new Unit(UnitMetric.Pixel, 1),
                "medium" => new Unit(UnitMetric.Pixel, 3),
                "thick" => new Unit(UnitMetric.Pixel, 5),
                _ => Unit.Empty,
            };
        }
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
