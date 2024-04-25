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
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml;

/// <summary>
/// Provides some utilies methods for translating Http attributes to OpenXml elements.
/// </summary>
static class Converter
{
    /// <summary>
    /// Convert the Html text align attribute (horizontal alignement) to its corresponding OpenXml value.
    /// </summary>
    public static JustificationValues? ToParagraphAlign(string? htmlAlign)
    {
        if (htmlAlign == null) return null;
        return htmlAlign.ToLowerInvariant() switch
        {
            "left" => JustificationValues.Left,
            "right" => JustificationValues.Right,
            "center" => JustificationValues.Center,
            "justify" => JustificationValues.Both,
            _ => null,
        };
    }

    /// <summary>
    /// Convert the Html vertical-align attribute to its corresponding OpenXml value.
    /// </summary>
    public static TableVerticalAlignmentValues? ToVAlign(string? htmlAlign)
    {
        if (htmlAlign == null) return null;
        return htmlAlign.ToLowerInvariant() switch
        {
            "top" => TableVerticalAlignmentValues.Top,
            "middle" => TableVerticalAlignmentValues.Center,
            "bottom" => TableVerticalAlignmentValues.Bottom,
            _ => null,
        };
    }

    /// <summary>
    /// Convert Html regular font-size to OpenXml font value (expressed in point).
    /// </summary>
    public static Unit ToFontSize(string? htmlSize)
    {
        if (htmlSize == null) return Unit.Empty;
        switch (htmlSize.ToLowerInvariant())
        {
            case "1":
            case "xx-small": return new Unit(UnitMetric.Point, 10);
            case "2":
            case "x-small": return new Unit(UnitMetric.Point, 15);
            case "3":
            case "small": return new Unit(UnitMetric.Point, 20);
            case "4":
            case "medium": return new Unit(UnitMetric.Point, 27);
            case "5":
            case "large": return new Unit(UnitMetric.Point, 36);
            case "6":
            case "x-large": return new Unit(UnitMetric.Point, 48);
            case "7":
            case "xx-large": return new Unit(UnitMetric.Point, 72);
            default:
                // the font-size is specified in positive half-points
                Unit unit = Unit.Parse(htmlSize);
                if (!unit.IsValid || unit.Value <= 0)
                    return Unit.Empty;

                // this is a rough conversion to support some percent size, considering 100% = 11 pt
                if (unit.Type == UnitMetric.Percent) unit = new Unit(UnitMetric.Point, unit.Value * 0.11);
                return unit;
        }
    }

    public static FontVariant? ToFontVariant(string? html)
    {
        if (html == null) return null;

        return html.ToLowerInvariant() switch
        {
            "small-caps" => FontVariant.SmallCaps,
            "normal" => FontVariant.Normal,
            _ => null,
        };
    }

    public static FontStyle? ToFontStyle(string? html)
    {
        if (html == null) return null;
        return html.ToLowerInvariant() switch
        {
            "italic" or "oblique" => FontStyle.Italic,
            "normal" => FontStyle.Normal,
            _ => null,
        };
    }

    public static FontWeight? ToFontWeight(string? html)
    {
        if (html == null) return null;
        return html.ToLowerInvariant() switch
        {
            "700" or "bold" => FontWeight.Bold,
            "bolder" => FontWeight.Bolder,
            "400" or "normal" => FontWeight.Normal,
            _ => null,
        };
    }

    public static string? ToFontFamily(string? str)
    {
        if (str == null) return null;

        var names = str.Split(',' );
        for (int i = 0; i < names.Length; i++)
        {
            string fontName = names[i];
            if (fontName.Length == 0) continue;
            try
            {
                if (fontName[0] == '\'' && fontName[fontName.Length-1] == '\'') fontName = fontName.Substring(1, fontName.Length - 2);
                return fontName;
            }
            catch (ArgumentException)
            {
                // the name is not a TrueType font or is not a font installed on this computer
            }
        }

        return null;
    }

    public static BorderValues ToBorderStyle(string? borderStyle)
    {
        if (borderStyle == null) return BorderValues.Nil;
        return borderStyle.ToLowerInvariant() switch
        {
            "dotted" => BorderValues.Dotted,
            "dashed" => BorderValues.Dashed,
            "solid" => BorderValues.Single,
            "double" => BorderValues.Double,
            "inset" => BorderValues.Inset,
            "outset" => BorderValues.Outset,
            "none" => BorderValues.None,
            _ => BorderValues.Nil,
        };
    }

    public static UnitMetric ToUnitMetric(string? type)
    {
        if (type == null) return UnitMetric.Unknown;
        return type.ToLowerInvariant() switch
        {
            "%" => UnitMetric.Percent,
            "in" => UnitMetric.Inch,
            "cm" => UnitMetric.Centimeter,
            "mm" => UnitMetric.Millimeter,
            "em" => UnitMetric.EM,
            "ex" => UnitMetric.Ex,
            "pt" => UnitMetric.Point,
            "pc" => UnitMetric.Pica,
            "px" => UnitMetric.Pixel,
            _ => UnitMetric.Unknown,
        };
    }

    public static PageOrientationValues ToPageOrientation(string? orientation)
    {
        if ( "landscape".Equals(orientation,StringComparison.OrdinalIgnoreCase))
            return PageOrientationValues.Landscape;

        return PageOrientationValues.Portrait;
    }

    public static TextDecoration ToTextDecoration(string? html)
    {
        // this style could take multiple values separated by a space
        // ex: text-decoration: blink underline;

        TextDecoration decoration = TextDecoration.None;

        if (html == null) return decoration;
        foreach (string part in html.ToLowerInvariant().Split(' '))
        {
            switch (part)
            {
                case "underline": decoration |= TextDecoration.Underline; break;
                case "line-through": decoration |= TextDecoration.LineThrough; break;
                default: break; // blink and overline are not supported
            }
        }
        return decoration;
    }
}