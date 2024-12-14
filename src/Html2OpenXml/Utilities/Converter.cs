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
using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml;

/// <summary>
/// Provides some utilities methods for translating Http attributes to OpenXml elements.
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
    public static Unit ToFontSize(ReadOnlySpan<char> span)
    {
        if (span.IsEmpty) return Unit.Empty;

        Span<char> loweredValue = span.Length <= 128 ? stackalloc char[span.Length] : new char[span.Length];
        span.ToLowerInvariant(loweredValue);
        var unit = loweredValue switch
        {
            "1" or "xx-small" => new Unit(UnitMetric.Point, 10),
            "2" or "x-small" => new Unit(UnitMetric.Point, 15),
            "3" or "small" => new Unit(UnitMetric.Point, 20),
            "4" or "medium" => new Unit(UnitMetric.Point, 27),
            "5" or "large" => new Unit(UnitMetric.Point, 36),
            "6" or "x-large" => new Unit(UnitMetric.Point, 48),
            "7" or "xx-large" => new Unit(UnitMetric.Point, 72),
            _ => Unit.Empty
        };

        if (!unit.IsValid)
        {
            // the font-size is specified in positive half-points
            unit = Unit.Parse(loweredValue, UnitMetric.Pixel);
            if (!unit.IsValid || unit.Value <= 0)
                return Unit.Empty;

            // this is a rough conversion to support some percent size, considering 100% = 11 pt
            if (unit.Metric == UnitMetric.Percent) unit = new Unit(UnitMetric.Point, unit.Value * 0.11);
        }
        return unit;
    }

    public static FontVariant? ToFontVariant(ReadOnlySpan<char> span)
    {
        if (span.IsEmpty) return null;

        Span<char> loweredValue = span.Length <= 128 ? stackalloc char[span.Length] : new char[span.Length];
        span.ToLowerInvariant(loweredValue);
        return loweredValue switch
        {
            "small-caps" => FontVariant.SmallCaps,
            "normal" => FontVariant.Normal,
            _ => null,
        };
    }

    public static FontStyle? ToFontStyle(ReadOnlySpan<char> span)
    {
        if (span.IsEmpty) return null;

        Span<char> loweredValue = span.Length <= 128 ? stackalloc char[span.Length] : new char[span.Length];
        span.ToLowerInvariant(loweredValue);
        return loweredValue switch
        {
            "italic" or "oblique" => FontStyle.Italic,
            "normal" => FontStyle.Normal,
            _ => null,
        };
    }

    public static FontWeight? ToFontWeight(ReadOnlySpan<char> span)
    {
        if (span.IsEmpty) return null;

        Span<char> loweredValue = span.Length <= 128 ? stackalloc char[span.Length] : new char[span.Length];
        span.ToLowerInvariant(loweredValue);
        return loweredValue switch
        {
            "700" or "bold" => FontWeight.Bold,
            "bolder" => FontWeight.Bolder,
            "400" or "normal" => FontWeight.Normal,
            _ => null,
        };
    }

    public static string? ToFontFamily(ReadOnlySpan<char> span)
    {
        if (span.IsEmpty) return null;

        // return the first font name
        Span<Range> tokens = stackalloc Range[1];
        return span.SplitCompositeAttribute(tokens, ',') switch {
            1 => span.Slice(tokens[0]).ToString(),
            _ => null
        };
    }

    public static BorderValues ToBorderStyle(ReadOnlySpan<char> span)
    {
        if (span.IsEmpty)
            return BorderValues.Nil;

        Span<char> loweredValue = span.Length <= 128 ? stackalloc char[span.Length] : new char[span.Length];
        span.ToLowerInvariant(loweredValue);
        return loweredValue switch
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

    public static PageOrientationValues ToPageOrientation(string? orientation)
    {
        if ( "landscape".Equals(orientation,StringComparison.OrdinalIgnoreCase))
            return PageOrientationValues.Landscape;

        return PageOrientationValues.Portrait;
    }

    public static ICollection<TextDecoration> ToTextDecoration(ReadOnlySpan<char> values)
    {
        // this style could take multiple values separated by a space
        // ex: text-decoration: blink underline;

        var decorations = new List<TextDecoration>();
        if (values.IsEmpty) return decorations;

        Span<char> loweredValue = values.Length <= 128 ? stackalloc char[values.Length] : new char[values.Length];
        values.ToLowerInvariant(loweredValue);

        Span<Range> tokens = stackalloc Range[5];
        ReadOnlySpan<char> span = loweredValue;
        var tokenCount = span.Split(tokens, ' ', StringSplitOptions.RemoveEmptyEntries);
        for (int i = 0; i < tokenCount; i++)
        {
            switch (span.Slice(tokens[i]))
            {
                case "underline": decorations.Add(TextDecoration.Underline); break;
                case "line-through": decorations.Add(TextDecoration.LineThrough); break;
                case "double": decorations.Add(TextDecoration.Double); break;
                case "dotted": decorations.Add(TextDecoration.Dotted); break;
                case "dashed": decorations.Add(TextDecoration.Dashed); break;
                case "wavy": decorations.Add(TextDecoration.Wave); break;
                default: break; // blink and overline are not supported
            }
        }
        return decorations;
    }

    public static T? ToBorder<T>(SideBorder border) where T: BorderType, new()
    {
        if (!border.IsValid)
            return null;

        return new T() { 
            Val = border.Style,
            Color = border.Color.ToHexString(),
            // according to MSDN,  sz=24 = 3 point
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.tablecellborders
            Size = (uint) Math.Round(border.Width.ValueInPoint * 8),
            Space = 1U
        };
    }

    public static CultureInfo? ToLanguage(string language)
    {
        try
        {
            var ci = new CultureInfo(language);
            if (ci.LCID != 4096) // custom unspecified
                return ci;
        }
        catch (ArgumentException)
        {
            // lang not valid, ignore it
        }
        return null;
    }
}