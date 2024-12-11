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
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml;

/// <summary>
/// Represents the collection of attributes present in the current html tag.
/// </summary>
sealed partial class HtmlAttributeCollection
{
    private readonly string rawValue;
    // Style key associated with a pointer to rawValue.
    private readonly Dictionary<string, Range> attributes = [];

    private HtmlAttributeCollection(string htmlStyles)
    {
        rawValue = htmlStyles;
    }

    /// <summary>
    /// Gets a value that indicates whether this collection is empty.
    /// </summary>
    public bool IsEmpty => attributes.Count == 0;

    public static HtmlAttributeCollection ParseStyle(string? htmlStyles)
    {
        var collection = new HtmlAttributeCollection(htmlStyles!);
        if (string.IsNullOrWhiteSpace(htmlStyles)) return collection;

        var span = htmlStyles.AsSpan();

        // Encoded ':' and ';' characters are valid for browser
        // <span style="text-decoration&#58;underline&#59;color:red">

        int startIndex = 0;
        bool foundKey = false;
        string? key = null;
        while (span.Length > 0)
        {
            int index = span.IndexOfAny(';', '&', ':');
            if (index == -1)
            {
                if (!foundKey) break;
                index = span.Length;
            }

            int separatorSize = 0;
            if (!foundKey)
            {
                if (span.Slice(index).StartsWith(['&','#','5','8',';']))
                {
                    separatorSize = 5;
                }
                else if (span[index] == ':')
                {
                    separatorSize = 1;
                }
                else
                {
                    // unexpected semicolon (ie, key with no value) -> ignore this style
                    separatorSize = -1;
                }

                if (separatorSize > 0 && index > 0)
                {
                    key = span.Slice(0, index).ToString().Trim();
                    foundKey = true;
                }
            }
            else
            {
                if (index < span.Length)
                {
                    if (span.Slice(index).StartsWith(['&','#','5','9',';']))
                    {
                        separatorSize = 5;
                    }
                    else if (span[index] == ';')
                    {
                        separatorSize = 1;
                    }
                    else if (span[index] == ':')
                    {
                        // unexpected colon (ie, key:value:value) -> ignore this style
                        separatorSize = -1;
                        foundKey = false;
                    }
                }

                if (foundKey)
                {
                    if (index > 0)
                        collection.attributes[key!] = new Range(startIndex, startIndex + index);
                    foundKey = false;
                }
            }

            separatorSize = Math.Abs(separatorSize);
            startIndex += index + separatorSize;
            span = span.Slice(index + separatorSize);
        }

        return collection;
    }

    /// <summary>
    /// Gets the named attribute.
    /// </summary>
    public string? this[string name]
    {
        get 
        {
            if (attributes.TryGetValue(name, out var range))
                return rawValue.AsSpan().Slice(range).ToString().Trim();
            return null;
        }
    }

    /// <summary>
    /// Gets an attribute representing a color (named color, hexadecimal or hexadecimal 
    /// without the preceding # character).
    /// </summary>
    public HtmlColor GetColor(string name)
    {
        if (attributes.TryGetValue(name, out var range))
            return HtmlColor.Parse(rawValue.AsSpan().Slice(range));
        return HtmlColor.Empty;
    }

    /// <summary>
    /// Gets an attribute representing an unit: 120px, 10pt, 5em, 20%, ...
    /// </summary>
    /// <returns>If the attribute is misformed, the <see cref="Unit.IsValid"/> property is set to false.</returns>
    public Unit GetUnit(string name, UnitMetric defaultMetric = UnitMetric.Unitless)
    {
        if (attributes.TryGetValue(name, out var range))
            return Unit.Parse(rawValue.AsSpan().Slice(range), defaultMetric);
        return Unit.Empty;
    }

    /// <summary>
    /// Gets an attribute representing the 4 unit sides.
    /// If a side has been specified individually, it will override the grouped definition.
    /// </summary>
    /// <returns>If the attribute is misformed, the <see cref="Margin.IsValid"/> property is set to false.</returns>
    public Margin GetMargin(string name)
    {
        Margin margin = Margin.Empty;
        if (attributes.TryGetValue(name, out var range))
            margin = Margin.Parse(rawValue.AsSpan().Slice(range));

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
        SideBorder border = SideBorder.Empty;
        if (attributes.TryGetValue(name, out Range range))
            border = SideBorder.Parse(rawValue.AsSpan().Slice(range));

        // handle attributes specified individually.
        Unit width = border.Width;
        if (attributes.TryGetValue(name + "-width", out range))
        {
            var w = SideBorder.ParseWidth(rawValue.AsSpan().Slice(range));
            if (width.IsValid) width = w;
        }

        var color = GetColor(name + "-color");
        if (color.IsEmpty) color = border.Color;

        BorderValues style = border.Style;
        if (attributes.TryGetValue(name + "-style", out range))
        {
            var s = Converter.ToBorderStyle(rawValue.AsSpan().Slice(range));
            if (s != BorderValues.Nil) style = s;
        }

        return new SideBorder(style, color, width);
    }

    /// <summary>
    /// Gets the `font` attribute and combine with the style, size and family.
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

    /// <summary>
    /// Gets the composite `text-decoration` style.
    /// </summary>
    public IEnumerable<TextDecoration> GetTextDecorations(string name)
    {
        if (attributes.TryGetValue(name, out Range range))
            return Converter.ToTextDecoration(rawValue.AsSpan().Slice(range));
        return [];
    }
}
