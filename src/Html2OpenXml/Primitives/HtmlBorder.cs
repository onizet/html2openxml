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

using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml;

/// <summary>
/// Represents a Html Border with the 4 sides.
/// </summary>
/// <remarks>Apply to all four sides.</remarks>
struct HtmlBorder(SideBorder all)
{
    public static readonly HtmlBorder Empty = new();

    private SideBorder top = all, right = all, bottom = all, left = all;


    public void ParseWidths(ReadOnlySpan<char> span)
    {
        var widths = BoxThickness.Parse(span);
        if (!widths.IsValid) return;

        Top = Top.Apply(widths.Top);
        Left = Left.Apply(widths.Left);
        Right = Right.Apply(widths.Right);
        Bottom = Bottom.Apply(widths.Bottom);
    }

    public void ParseColors(ReadOnlySpan<char> span)
    {
        if (span.IsEmpty || span.IsWhiteSpace())
            return;

        if (!CssShorthand.TryParse(span,
            s => HtmlColor.Parse(s), c => !c.IsEmpty, out var sides))
        {
            return;
        }

        top = top.Apply(sides.Top);
        right = right.Apply(sides.Right);
        bottom = bottom.Apply(sides.Bottom);
        left = left.Apply(sides.Left);
    }

    public void ParseStyles(ReadOnlySpan<char> span)
    {
        if (span.IsEmpty || span.IsWhiteSpace())
            return;

        if (!CssShorthand.TryParse(span,
            s => Converter.ToBorderStyle(s), s => s != BorderValues.Nil, out var sides))
        {
            return;
        }

        top = top.Apply(sides.Top);
        right = right.Apply(sides.Right);
        bottom = bottom.Apply(sides.Bottom);
        left = left.Apply(sides.Left);
    }

    //____________________________________________________________________
    //

    /// <summary>
    /// Gets or sets the unit of the bottom side.
    /// </summary>
    public SideBorder Bottom
    {
        readonly get => bottom;
        set => bottom = value;
    }

    /// <summary>
    /// Gets or sets the unit of the left side.
    /// </summary>
    public SideBorder Left
    {
        readonly get => left;
        set => left = value;
    }

    /// <summary>
    /// Gets or sets the unit of the top side.
    /// </summary>
    public SideBorder Top
    {
        readonly get => top;
        set => top = value;
    }

    /// <summary>
    /// Gets or sets the unit of the right side.
    /// </summary>
    public SideBorder Right
    {
        readonly get => right;
        set => right = value;
    }

    /// <summary>
    /// Gets whether at least one side has been specified.
    /// </summary>
    public readonly bool IsEmpty
    {
        get => !(Left.IsValid || Right.IsValid || Bottom.IsValid || Top.IsValid);
    }
}