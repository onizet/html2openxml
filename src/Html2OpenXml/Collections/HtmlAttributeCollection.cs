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
using System.Collections.Generic;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml;

/// <summary>
/// Represents the collection of attributes present in the current html tag.
/// </summary>
sealed class HtmlAttributeCollection
{
    private static readonly Regex stripStyleAttributesRegex = new(@"(?<name>[^;\s]+)\s?(&\#58;|:)\s?(?<val>[^;&]+)\s?(;|&\#59;)*");
    private readonly Dictionary<string, string> attributes = [];



    private HtmlAttributeCollection()
    {
    }

    public static HtmlAttributeCollection ParseStyle(string? htmlTag)
    {
        var collection = new HtmlAttributeCollection();
        if (string.IsNullOrEmpty(htmlTag)) return collection;

        // Encoded ':' and ';' characters are valid for browser but not handled by the regex (bug #13812 reported by robin391)
        // ex= <span style="text-decoration&#58;underline&#59;color:red">
        MatchCollection matches = stripStyleAttributesRegex.Matches(htmlTag);
        foreach (Match m in matches)
            collection.attributes[m.Groups["name"].Value] = m.Groups["val"].Value;

        return collection;
    }

    /// <summary>
    /// Gets the named attribute.
    /// </summary>
    public string? this[string name]
    {
        get => attributes.TryGetValue(name, out var value)? value : null;
    }

    /// <summary>
    /// Gets an attribute representing a color (named color, hexadecimal or hexadecimal 
    /// without the preceding # character).
    /// </summary>
    public HtmlColor GetColor(string name)
    {
        return HtmlColor.Parse(this[name]);
    }

    /// <summary>
    /// Gets an attribute representing an unit: 120px, 10pt, 5em, 20%, ...
    /// </summary>
    /// <returns>If the attribute is misformed, the <see cref="Unit.IsValid"/> property is set to false.</returns>
    public Unit GetUnit(string name, UnitMetric defaultMetric = UnitMetric.Unitless)
    {
        return Unit.Parse(this[name], defaultMetric);
    }

    /// <summary>
    /// Gets an attribute representing the 4 unit sides.
    /// If a side has been specified individually, it will override the grouped definition.
    /// </summary>
    /// <returns>If the attribute is misformed, the <see cref="Margin.IsValid"/> property is set to false.</returns>
    public Margin GetMargin(string name)
    {
        Margin margin = Margin.Parse(this[name]);
        Unit u;

        u = GetUnit(name + "-top", UnitMetric.Pixel);
        if (u.IsValid) margin.Top = u;
        u = GetUnit(name + "-right", UnitMetric.Pixel);
        if (u.IsValid) margin.Right = u;
        u = GetUnit(name + "-bottom", UnitMetric.Pixel);
        if (u.IsValid) margin.Bottom = u;
        u = GetUnit(name + "-left", UnitMetric.Pixel);
        if (u.IsValid) margin.Left = u;

        return margin;
    }

    /// <summary>
    /// Gets an attribute representing the 4 border sides.
    /// If a border style/color/width has been specified individually, it will override the grouped definition.
    /// </summary>
    /// <returns>If the attribute is misformed, the <see cref="HtmlBorder.IsEmpty"/> property is set to false.</returns>
    public HtmlBorder GetBorders()
    {
        HtmlBorder border = new(GetSideBorder("border"));
        SideBorder sb;

        sb = GetSideBorder("border-top");
        if (sb.IsValid) border.Top = sb;
        sb = GetSideBorder("border-right");
        if (sb.IsValid) border.Right = sb;
        sb = GetSideBorder("border-bottom");
        if (sb.IsValid) border.Bottom = sb;
        sb = GetSideBorder("border-left");
        if (sb.IsValid) border.Left = sb;

        return border;
    }

    /// <summary>
    /// Gets an attribute representing a single border side.
    /// If a border style/color/width has been specified individually, it will override the grouped definition.
    /// </summary>
    /// <returns>If the attribute is misformed, the <see cref="HtmlBorder.IsEmpty"/> property is set to false.</returns>
    public SideBorder GetSideBorder(string name)
    {
        var attrValue = this[name];
        SideBorder border = SideBorder.Parse(attrValue);

        // handle attributes specified individually.
        Unit width = SideBorder.ParseWidth(this[name + "-width"]);
        if (!width.IsValid) width = border.Width;

        var color = GetColor(name + "-color");
        if (color.IsEmpty) color = border.Color;

        var style = Converter.ToBorderStyle(this[name + "-style"]);
        if (style == BorderValues.Nil) style = border.Style;

        return new SideBorder(style, color, width);
    }

    /// <summary>
    /// Gets the font attribute and combine with the style, size and family.
    /// </summary>
    public HtmlFont GetFont(string name)
    {
        HtmlFont font = HtmlFont.Parse(this[name]);
        FontStyle? fontStyle = font.Style;
        FontVariant? variant = font.Variant;
        FontWeight? weight = font.Weight;
        Unit fontSize = font.Size;
        string? family = font.Family;

        var attrValue = this[name + "-style"];
        if (attrValue != null)
        {
            fontStyle = Converter.ToFontStyle(attrValue) ?? font.Style;
        }
        attrValue = this[name + "-variant"];
        if (attrValue != null)
        {
            variant = Converter.ToFontVariant(attrValue) ?? font.Variant;
        }
        attrValue = this[name + "-weight"];
        if (attrValue != null)
        {
            weight = Converter.ToFontWeight(attrValue) ?? font.Weight;
        }
        attrValue = this[name + "-family"];
        if (attrValue != null)
        {
            family = Converter.ToFontFamily(attrValue) ?? font.Family;
        }

        Unit unit = this.GetUnit(name + "-size");
        if (unit.IsValid) fontSize = unit;

        return new HtmlFont(fontStyle, variant, weight, fontSize, family);
    }
}
