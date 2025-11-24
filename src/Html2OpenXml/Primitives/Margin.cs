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
/// Represents a Html Margin.
/// </summary>
struct Margin
{
    /// <summary>Represents an empty margin (not defined).</summary>
    public static readonly Margin Empty = new();

    private Unit top;
    private Unit right;
    private Unit bottom;
    private Unit left;


    /// <summary>Apply to all four sides.</summary>
    public Margin(Unit all)
    {
        this.top = all;
        this.right = all;
        this.bottom = all;
        this.left = all;
    }

    /// <summary>Top and bottom | left and right.</summary>
    public Margin(Unit topAndBottom, Unit leftAndRight)
    {
        this.top = topAndBottom;
        this.bottom = topAndBottom;
        this.left = leftAndRight;
        this.right = leftAndRight;
    }

    /// <summary>Top | left and right | bottom.</summary>
    public Margin(Unit top, Unit leftAndRight, Unit bottom)
    {
        this.top = top;
        this.right = leftAndRight;
        this.bottom = bottom;
        this.left = leftAndRight;
    }

    /// <summary>Top | right | bottom | left.</summary>
    public Margin(Unit top, Unit right, Unit bottom, Unit left)
    {
        this.top = top;
        this.right = right;
        this.bottom = bottom;
        this.left = left;
    }

    /// <summary>
    /// Parse the margin style attribute.
    /// </summary>
    /// <remarks>
    /// The margin property can have from one to four values.
    /// <b>margin:25px 50px 75px 100px;</b>
    /// top margin is 25px
    /// right margin is 50px
    /// bottom margin is 75px
    /// left margin is 100px
    /// 
    /// <b>margin:25px 50px 75px;</b>
    /// top margin is 25px
    /// right and left margins are 50px
    /// bottom margin is 75px
    /// 
    /// <b>margin:25px 50px;</b>
    /// top and bottom margins are 25px
    /// right and left margins are 50px
    /// 
    /// <b>margin:25px;</b>
    /// all four margins are 25px
    /// </remarks>
    public static Margin Parse(ReadOnlySpan<char> span)
    {
        if (span.IsEmpty || span.IsWhiteSpace())
            return Empty;

        Span<Range> tokens = stackalloc Range[5];
        return span.SplitCompositeAttribute(tokens) switch
        {
            1 => new Margin(Unit.Parse(span.Slice(tokens[0]), UnitMetric.Pixel)),
            2 => new Margin(
                Unit.Parse(span.Slice(tokens[0]), UnitMetric.Pixel),
                Unit.Parse(span.Slice(tokens[1]), UnitMetric.Pixel)),
            3 => new Margin(
                Unit.Parse(span.Slice(tokens[0]), UnitMetric.Pixel),
                Unit.Parse(span.Slice(tokens[1]), UnitMetric.Pixel),
                Unit.Parse(span.Slice(tokens[2]), UnitMetric.Pixel)),
            4 => new Margin(
                Unit.Parse(span.Slice(tokens[0]), UnitMetric.Pixel),
                Unit.Parse(span.Slice(tokens[1]), UnitMetric.Pixel),
                Unit.Parse(span.Slice(tokens[2]), UnitMetric.Pixel),
                Unit.Parse(span.Slice(tokens[3]), UnitMetric.Pixel)),
            _ => Empty
        };
    }

    //____________________________________________________________________
    //

    /// <summary>
    /// Gets or sets the unit of the bottom side.
    /// </summary>
    public Unit Bottom
    {
        readonly get => bottom;
        set => bottom = value;
    }

    /// <summary>
    /// Gets or sets the unit of the left side.
    /// </summary>
    public Unit Left
    {
        readonly get => left;
        set => left = value;
    }

    /// <summary>
    /// Gets or sets the unit of the top side.
    /// </summary>
    public Unit Top
    {
        readonly get => top;
        set => top = value;
    }

    /// <summary>
    /// Gets or sets the unit of the right side.
    /// </summary>
    public Unit Right
    {
        readonly get => right;
        set => right = value;
    }

    public bool IsValid
    {
        get => Left.IsValid && Right.IsValid && Bottom.IsValid && Top.IsValid;
    }

    /// <summary>
    /// Gets whether at least one side has been specified.
    /// </summary>
    public bool IsEmpty
    {
        get => !(Left.IsValid || Right.IsValid || Bottom.IsValid || Top.IsValid);
    }
}
