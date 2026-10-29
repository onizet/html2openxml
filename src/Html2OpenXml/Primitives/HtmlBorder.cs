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

namespace HtmlToOpenXml;

/// <summary>
/// Represents a Html Border with the 4 sides.
/// </summary>
struct HtmlBorder
{
    public static readonly HtmlBorder Empty = new();

    private SideBorder top;
    private SideBorder right;
    private SideBorder bottom;
    private SideBorder left;


    /// <summary>Apply to all four sides.</summary>
    public HtmlBorder(SideBorder all)
    {
        this.top = all;
        this.right = all;
        this.bottom = all;
        this.left = all;
    }

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

        Span<Range> tokens = stackalloc Range[4];
        HtmlColor[] colors;
        var tokenCount = span.SplitCompositeAttribute(tokens, separator: ' ');
        if (tokenCount < 1 || tokenCount > 4)
            return;

        if (tokenCount == 1)
        {
            var c = HtmlColor.Parse(span[tokens[0]]);
            if (c.IsEmpty) return;
            colors = [c, c, c, c];
        }
        else if (tokenCount == 2)
        {
            var c1 = HtmlColor.Parse(span[tokens[0]]);
            var c2 = HtmlColor.Parse(span[tokens[1]]);
            if (c1.IsEmpty || c2.IsEmpty) return;

            colors = [c1, c2, c1, c2];
        }
        else if (tokenCount == 3)
        {
            var c1 = HtmlColor.Parse(span[tokens[0]]);
            var c2 = HtmlColor.Parse(span[tokens[1]]);
            var c3 = HtmlColor.Parse(span[tokens[2]]);
            if (c1.IsEmpty || c2.IsEmpty || c3.IsEmpty) return;

            colors = [c1, c2, c3, c2];
        }
        else
        {
            var c1 = HtmlColor.Parse(span[tokens[0]]);
            var c2 = HtmlColor.Parse(span[tokens[1]]);
            var c3 = HtmlColor.Parse(span[tokens[2]]);
            var c4 = HtmlColor.Parse(span[tokens[2]]);
            if (c1.IsEmpty || c2.IsEmpty || c3.IsEmpty || c4.IsEmpty) return;

            colors = [c1, c2, c3, c4];
        }

        top = top.Apply(colors[0]);
        right = right.Apply(colors[1]);
        bottom = bottom.Apply(colors[2]);
        left = left.Apply(colors[3]);
    }

    public void ParseStyles(ReadOnlySpan<char> span)
    {
        
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