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
    private SideBorder[]? sides;


    public HtmlBorder(SideBorder all)
    {
        if (!all.IsValid) sides = null;
        else this.sides = [all, all, all, all];
    }

    private void EnsureSides()
    {
        if(this.sides == null) sides = new SideBorder[4];
    }

    //____________________________________________________________________
    //

    /// <summary>
    /// Gets or sets the border of the bottom side.
    /// </summary>
    public SideBorder Bottom
    {
        readonly get { return sides == null ? SideBorder.Empty : sides[2]; }
        set { EnsureSides(); sides![2] = value; }
    }

    /// <summary>
    /// Gets or sets the border of the left side.
    /// </summary>
    public SideBorder Left
    {
        readonly get { return sides == null ? SideBorder.Empty : sides[3]; }
        set { EnsureSides(); sides![3] = value; }
    }

    /// <summary>
    /// Gets or sets the border of the top side.
    /// </summary>
    public SideBorder Top
    {
        readonly get { return sides == null ? SideBorder.Empty : sides[0]; }
        set { EnsureSides(); sides![0] = value; }
    }

    /// <summary>
    /// Gets or sets the border of the right side.
    /// </summary>
    public SideBorder Right
    {
        readonly get { return sides == null ? SideBorder.Empty : sides[1]; }
        set { EnsureSides(); sides![1] = value; }
    }

    /// <summary>
    /// Gets whether at least one side has been specified.
    /// </summary>
    public readonly bool IsEmpty
    {
        get => sides == null || !(Left.IsValid || Right.IsValid || Bottom.IsValid || Top.IsValid);
    }
}