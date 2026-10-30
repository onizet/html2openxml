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
    public static JustificationValues? ToParagraphAlign(ReadOnlySpan<char> span)
    {
        if (span.Equals("left", StringComparison.OrdinalIgnoreCase))
            return JustificationValues.Left;
        if (span.Equals("right", StringComparison.OrdinalIgnoreCase))
            return JustificationValues.Right;
        if (span.Equals("center", StringComparison.OrdinalIgnoreCase))
            return JustificationValues.Center;
        if (span.Equals("justify", StringComparison.OrdinalIgnoreCase))
            return JustificationValues.Both;

        return null;
    }

    /// <summary>
    /// Convert the Html vertical-align attribute to its corresponding OpenXml value.
    /// </summary>
    public static TableVerticalAlignmentValues? ToVAlign(ReadOnlySpan<char> span)
    {
        if (span.Equals("top", StringComparison.OrdinalIgnoreCase))
            return TableVerticalAlignmentValues.Top;
        if (span.Equals("middle", StringComparison.OrdinalIgnoreCase))
            return TableVerticalAlignmentValues.Center;
        if (span.Equals("bottom", StringComparison.OrdinalIgnoreCase))
            return TableVerticalAlignmentValues.Bottom;

        return null;
    }

    /// <summary>
    /// Convert Html regular font-size to OpenXml font value (expressed in point).
    /// </summary>
    public static Unit ToFontSize(ReadOnlySpan<char> span)
    {
        if (span.IsEmpty) return Unit.Empty;

        if (span.Equals("1", StringComparison.OrdinalIgnoreCase)
            || span.Equals("xx-small", StringComparison.OrdinalIgnoreCase))
            return new Unit(UnitMetric.Point, 10);
        if (span.Equals("2", StringComparison.OrdinalIgnoreCase)
            || span.Equals("x-small", StringComparison.OrdinalIgnoreCase))
            return new Unit(UnitMetric.Point, 15);
        if (span.Equals("3", StringComparison.OrdinalIgnoreCase)
            || span.Equals("small", StringComparison.OrdinalIgnoreCase))
            return new Unit(UnitMetric.Point, 20);
        if (span.Equals("4", StringComparison.OrdinalIgnoreCase)
            || span.Equals("medium", StringComparison.OrdinalIgnoreCase))
            return new Unit(UnitMetric.Point, 27);
        if (span.Equals("5", StringComparison.OrdinalIgnoreCase)
            || span.Equals("large", StringComparison.OrdinalIgnoreCase))
            return new Unit(UnitMetric.Point, 36);
        if (span.Equals("6", StringComparison.OrdinalIgnoreCase)
            || span.Equals("x-large", StringComparison.OrdinalIgnoreCase))
            return new Unit(UnitMetric.Point, 48);
        if (span.Equals("7", StringComparison.OrdinalIgnoreCase)
            || span.Equals("xx-large", StringComparison.OrdinalIgnoreCase))
            return new Unit(UnitMetric.Point, 72);

        // the font-size is specified in positive half-points
        var unit = Unit.Parse(span, UnitMetric.Pixel);
        if (!unit.IsValid || unit.Value <= 0)
            return Unit.Empty;

        // this is a rough conversion to support some percent size, considering 100% = 11 pt
        if (unit.Metric == UnitMetric.Percent)
            unit = new Unit(UnitMetric.Point, unit.Value * 0.11);
        return unit;
    }

    public static FontVariant? ToFontVariant(ReadOnlySpan<char> span)
    {
        if (span.IsEmpty) return null;

        if (span.Equals("small-caps", StringComparison.OrdinalIgnoreCase))
            return FontVariant.SmallCaps;
        if (span.Equals("normal", StringComparison.OrdinalIgnoreCase))
            return FontVariant.Normal;

        return null;
    }

    public static FontStyle? ToFontStyle(ReadOnlySpan<char> span)
    {
        if (span.IsEmpty) return null;

        if (span.Equals("italic", StringComparison.OrdinalIgnoreCase)
            || span.Equals("oblique", StringComparison.OrdinalIgnoreCase))
            return FontStyle.Italic;
        if (span.Equals("normal", StringComparison.OrdinalIgnoreCase))
            return FontStyle.Normal;

        return null;
    }

    public static FontWeight? ToFontWeight(ReadOnlySpan<char> span)
    {
        if (span.IsEmpty) return null;

        if (span.Equals("700", StringComparison.OrdinalIgnoreCase)
            || span.Equals("bold", StringComparison.OrdinalIgnoreCase))
            return FontWeight.Bold;
        if (span.Equals("bolder", StringComparison.OrdinalIgnoreCase))
            return FontWeight.Bolder;
        if (span.Equals("400", StringComparison.OrdinalIgnoreCase)
            || span.Equals("normal", StringComparison.OrdinalIgnoreCase))
            return FontWeight.Normal;

        return null;
    }

    public static string? ToFontFamily(ReadOnlySpan<char> span)
    {
        if (span.IsEmpty) return null;

        // return the first font name
        Span<Range> tokens = stackalloc Range[1];
        return span.SplitCompositeAttribute(tokens, ',') switch
        {
            1 => span[tokens[0]].Unquote().Trim().ToString(),
            _ => null
        };
    }

    public static BorderValues ToBorderStyle(ReadOnlySpan<char> span)
    {
        if (span.IsEmpty)
            return BorderValues.Nil;

        if (span.Equals("dotted", StringComparison.OrdinalIgnoreCase))
            return BorderValues.Dotted;
        if (span.Equals("dashed", StringComparison.OrdinalIgnoreCase))
            return BorderValues.Dashed;
        if (span.Equals("solid", StringComparison.OrdinalIgnoreCase))
            return BorderValues.Single;
        if (span.Equals("double", StringComparison.OrdinalIgnoreCase))
            return BorderValues.Double;
        if (span.Equals("inset", StringComparison.OrdinalIgnoreCase))
            return BorderValues.Inset;
        if (span.Equals("outset", StringComparison.OrdinalIgnoreCase))
            return BorderValues.Outset;
        if (span.Equals("none", StringComparison.OrdinalIgnoreCase))
            return BorderValues.None;

        return BorderValues.Nil;
    }

    public static TextDecoration ToTextDecoration(ReadOnlySpan<char> span)
    {
        // this style could take multiple values separated by a space
        // ex: text-decoration: blink underline;

        var decorations = TextDecoration.None;
        if (span.IsEmpty) return decorations;

        Span<Range> tokens = stackalloc Range[5];
        var tokenCount = span.Split(tokens, ' ', StringSplitOptions.RemoveEmptyEntries);
        for (int i = 0; i < tokenCount; i++)
        {
            var token = span[tokens[i]];
            if (token.Equals("underline", StringComparison.OrdinalIgnoreCase))
                decorations |= TextDecoration.Underline;
            else if (token.Equals("line-through", StringComparison.OrdinalIgnoreCase))
                decorations |= TextDecoration.LineThrough;
            else if (token.Equals("double", StringComparison.OrdinalIgnoreCase))
                decorations |= TextDecoration.Double;
            else if (token.Equals("dotted", StringComparison.OrdinalIgnoreCase))
                decorations |= TextDecoration.Dotted;
            else if (token.Equals("dashed", StringComparison.OrdinalIgnoreCase))
                decorations |= TextDecoration.Dashed;
            else if (token.Equals("wavy", StringComparison.OrdinalIgnoreCase))
                decorations |= TextDecoration.Wave;
            // blink and overline are not supported
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
            Size = (uint) Math.Round(border.Width.ValueInPoint * 8)
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