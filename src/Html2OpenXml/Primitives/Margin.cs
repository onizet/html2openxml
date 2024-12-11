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
    public static readonly Margin Empty = new() { sides = new Unit[4] };
    private Unit[] sides;


    /// <summary>Apply to all four sides.</summary>
    public Margin(Unit all)
    {
        this.sides = [all, all, all, all];
    }

    /// <summary>Top and bottom | left and right.</summary>
    public Margin(Unit topAndBottom, Unit leftAndRight)
    {
        this.sides = [topAndBottom, leftAndRight, topAndBottom, leftAndRight];
    }

    /// <summary>Top | left and right | bottom.</summary>
    public Margin(Unit top, Unit leftAndRight, Unit bottom)
    {
        this.sides = [top, leftAndRight, bottom, leftAndRight];
    }

    /// <summary>Top | right | bottom | left.</summary>
    public Margin(Unit top, Unit right, Unit bottom, Unit left)
    {
        this.sides = [top, right, bottom, left];
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
    public static Margin Parse(string? str)
    {
        if (str == null)
            return Empty;
        return Parse(str.AsSpan());
    }

    public static Margin Parse(ReadOnlySpan<char> span)
    {
        if (span.Length == 0 || span.IsWhiteSpace())
            return Empty;

        Span<Range> tokens = stackalloc Range[5];
        return span.SplitHtmlCompositeAttribute(tokens) switch
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
        readonly get => sides[2];
        set { sides[2] = value; }
    }

    /// <summary>
    /// Gets or sets the unit of the left side.
    /// </summary>
    public Unit Left
    {
        readonly get => sides[3];
        set { sides[3] = value; }
    }

    /// <summary>
    /// Gets or sets the unit of the top side.
    /// </summary>
    public Unit Top
    {
        readonly get => sides[0];
        set { sides[0] = value; }
    }

    /// <summary>
    /// Gets or sets the unit of the right side.
    /// </summary>
    public Unit Right
    {
        readonly get => sides[1];
        set { sides[1] = value; }
    }

    public readonly bool IsValid
    {
        get => Left.IsValid && Right.IsValid && Bottom.IsValid && Top.IsValid;
    }

    /// <summary>
    /// Gets whether at least one side has been specified.
    /// </summary>
    public readonly bool IsEmpty
    {
        get => !(Left.IsValid || Right.IsValid || Bottom.IsValid || Top.IsValid);
    }
}
