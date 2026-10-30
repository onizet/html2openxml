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
/// Represents a Html margin or padding.
/// </summary>
struct BoxThickness
{
    /// <summary>Represents an empty box model (not defined).</summary>
    public static readonly BoxThickness Empty = new();

    private Unit top, right, bottom, left;


    /// <summary>Top | right | bottom | left.</summary>
    public BoxThickness(Unit top, Unit right, Unit bottom, Unit left)
    {
        this.top = top;
        this.right = right;
        this.bottom = bottom;
        this.left = left;
    }

    /// <summary>
    /// Parse the <c>margin</c> or <c>padding</c> style attribute.
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
    public static BoxThickness Parse(ReadOnlySpan<char> span)
    {
        if (span.IsEmpty || span.IsWhiteSpace())
            return Empty;


        if (!CssShorthand.TryParse(span,
            s => Unit.Parse(s, UnitMetric.Pixel), c => c.IsValid, out var sides))
        {
            return Empty;
        }

        return new(sides.Top, sides.Right, sides.Bottom, sides.Left);
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
