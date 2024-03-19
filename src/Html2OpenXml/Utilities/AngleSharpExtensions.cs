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
using AngleSharp.Html.Dom;
using HtmlToOpenXml;
using wp = DocumentFormat.OpenXml.Wordprocessing;

namespace AngleSharp.Dom;

/// <summary>
/// Helper class that provide some extension methods to HtmlAgilityPack SDK.
/// </summary>
[System.Diagnostics.DebuggerStepThrough]
static class AngleSharpExtension
{
    /*//// <summary>
    /// Gets an attribute representing an integer.
    /// </summary>
    public static Int32? GetAsInt(this ICssStyleDeclaration style, string name)
    {
        string attrValue = style.getp.GetAttribute(name);
        int val;
        if (attrValue != null && Int32.TryParse(attrValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out val))
            return val;

        return null;
    }*/

    /// <summary>
    /// Gets an attribute representing a color (named color, hexadecimal or hexadecimal 
    /// without the preceding # character).
    /// </summary>
    public static HtmlColor AttributeAsColor(this IHtmlElement node, string name)
    {
        return HtmlColor.Parse(node.GetAttribute(name));
    }

    /*/// <summary>
    /// Gets an attribute representing an unit: 120px, 10pt, 5em, 20%, ...
    /// </summary>
    /// <returns>If the attribute is misformed, the <see cref="Unit.IsValid"/> property is set to false.</returns>
    public static Unit GetAsUnit(this ICssStyleDeclaration style, string name)
    {
        return Unit.Parse(style.GetPropertyValue(name));
    }

    /// <summary>
    /// Gets an attribute representing the 4 unit sides.
    /// If a side has been specified individually, it will override the grouped definition.
    /// </summary>
    /// <returns>If the attribute is misformed, the <see cref="Margin.IsValid"/> property is set to false.</returns>
    public static Margin GetAsMargin(this ICssStyleDeclaration style, string name)
    {
        Margin margin = Margin.Parse(style.GetPropertyValue(name));
        Unit u;

        u = style.GetAsUnit(name + "-top");
        if (u.IsValid) margin.Top = u;
        u = style.GetAsUnit(name + "-right");
        if (u.IsValid) margin.Right = u;
        u = style.GetAsUnit(name + "-bottom");
        if (u.IsValid) margin.Bottom = u;
        u = style.GetAsUnit(name + "-left");
        if (u.IsValid) margin.Left = u;

        return margin;
    }

    /// <summary>
    /// Gets an attribute representing the 4 border sides.
    /// If a border style/color/width has been specified individually, it will override the grouped definition.
    /// </summary>
    /// <returns>If the attribute is misformed, the <see cref="HtmlBorder.IsEmpty"/> property is set to false.</returns>
    public static HtmlBorder GetAsBorder(this ICssStyleDeclaration style, string name)
    {
        HtmlBorder border = new HtmlBorder(style.GetAsSideBorder(name));
        SideBorder sb;

        sb = style.GetAsSideBorder(name + "-top");
        if (sb.IsValid) border.Top = sb;
        sb = style.GetAsSideBorder(name + "-right");
        if (sb.IsValid) border.Right = sb;
        sb = style.GetAsSideBorder(name + "-bottom");
        if (sb.IsValid) border.Bottom = sb;
        sb = style.GetAsSideBorder(name + "-left");
        if (sb.IsValid) border.Left = sb;

        return border;
    }

    /// <summary>
    /// Gets an attribute representing a single border side.
    /// If a border style/color/width has been specified individually, it will override the grouped definition.
    /// </summary>
    /// <returns>If the attribute is misformed, the <see cref="HtmlBorder.IsEmpty"/> property is set to false.</returns>
    public static SideBorder GetAsSideBorder(this ICssStyleDeclaration style, string name)
    {
        string attrValue = style.GetPropertyValue(name);
        SideBorder border = SideBorder.Parse(attrValue);

        // handle attributes specified individually.
        Unit width = SideBorder.ParseWidth(style.GetPropertyValue(name + "-width"));
        if (width.IsValid) border.Width = width;

        var color = style.GetAsColor(name + "-color");
        if (!color.IsEmpty) border.Color = color;

        var borderStyle = Converter.ToBorderStyle(style.GetPropertyValue(name + "-style"));
        if (borderStyle != wp.BorderValues.Nil) border.Style = borderStyle;

        return border;
    }

    /// <summary>
    /// Gets the font attribute and combine with the style, size and family.
    /// </summary>
    public static HtmlFont GetAttributeAsFont(this ICssStyleDeclaration style, string name)
    {
        string attrValue = style.GetPropertyValue(name);
        HtmlFont font = HtmlFont.Parse(attrValue);
        attrValue = style.GetPropertyValue(name + "-style");
        if (attrValue != null)
        {
            var fontStyle = Converter.ToFontStyle(attrValue);
            if (fontStyle.HasValue) font.Style = fontStyle.Value;
        }
        attrValue = style.GetPropertyValue(name + "-variant");
        if (attrValue != null)
        {
            var variant = Converter.ToFontVariant(attrValue);
            if (variant.HasValue) font.Variant = variant.Value;
        }
        attrValue = style.GetPropertyValue(name + "-weight");
        if (attrValue != null)
        {
            var weight = Converter.ToFontWeight(attrValue);
            if (weight.HasValue) font.Weight = weight.Value;
        }
        attrValue = style.GetPropertyValue(name + "-family");
        if (attrValue != null)
        {
            font.Family = Converter.ToFontFamily(attrValue);
        }
        Unit unit = style.GetAsUnit(name + "-size");
        if (unit.IsValid) font.Size = unit;
        return font;
    }*/
}