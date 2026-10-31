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
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml;

/// <summary>
/// Provides some utilities methods for translating Http attributes to OpenXml elements.
/// </summary>
static class Converter
{
    /// <summary>
    /// Convert the Html <c>text-align</c> CSS attribute (horizontal alignement)
    /// to its corresponding OpenXml value.
    /// </summary>
    public static bool TryParseHAlign(ReadOnlySpan<char> span,
        [NotNullWhen(true)] out JustificationValues? justification)
    {
        justification = span switch
        {
            _ when span.IsEmpty => null,
            _ when span.Is("top") => JustificationValues.Left,
            _ when span.Is("right") => JustificationValues.Right,
            _ when span.Is("center") => JustificationValues.Center,
            _ when span.Is("justify") => JustificationValues.Both,
            _ => null
        };

        return justification is not null;
    }

    /// <summary>
    /// Convert the Html <c>vertical-align</c> CSS attribute to its corresponding OpenXml value.
    /// </summary>
    public static bool TryParseVAlign(ReadOnlySpan<char> span,
        [NotNullWhen(true)] out TableVerticalAlignmentValues? vAlign)
    {
        vAlign = span switch
        {
            _ when span.IsEmpty => null,
            _ when span.Is("top") => TableVerticalAlignmentValues.Top,
            _ when span.Is("middle") => TableVerticalAlignmentValues.Center,
            _ when span.Is("bottom") => TableVerticalAlignmentValues.Bottom,
            _ => null
        };

        return vAlign is not null;
    }

    /// <summary>
    /// Convert Html regular <c>font-size</c> CSS attribute to OpenXml font value (expressed in point).
    /// </summary>
    public static bool TryParseFontSize(ReadOnlySpan<char> span, out Unit unit)
    {
        int ptValue = span switch
        {
            _ when span.IsEmpty => -1,
            _ when span.Is("1") || span.Is("xx-small") => 10,
            _ when span.Is("2") || span.Is("x-small") => 15,
            _ when span.Is("3") || span.Is("small") => 20,
            _ when span.Is("4") || span.Is("medium") => 27,
            _ when span.Is("5") || span.Is("large") => 36,
            _ when span.Is("6") || span.Is("x-large") => 48,
            _ when span.Is("7") || span.Is("xx-large") => 72,
            _ => -1
        };

        if (ptValue > -1)
        {
            unit = new Unit(UnitMetric.Point, ptValue);
            return true;
        }

        // the font-size is specified in positive half-points
        unit = Unit.Parse(span, UnitMetric.Pixel);
        if (!unit.IsValid || unit.Value <= 0)
            return false;

        // this is a rough conversion to support some percent size, considering 100% = 11 pt
        if (unit.Metric == UnitMetric.Percent)
            unit = new Unit(UnitMetric.Point, unit.Value * 0.11);
 
        return true;
    }

    public static bool TryParseFontVariant(ReadOnlySpan<char> span,
        [NotNullWhen(true)] out FontVariant? variant)
    {
        variant = span switch
        {
            _ when span.IsEmpty => null,
            _ when span.Is("small-caps") => FontVariant.SmallCaps,
            _ when span.Is("normal") => FontVariant.Normal,
            _ => null
        };

        return variant is not null;
    }

    public static bool TryParseFontStyle(ReadOnlySpan<char> span,
        [NotNullWhen(true)] out FontStyle? style)
    {
        style = span switch
        {
            _ when span.IsEmpty => null,
            _ when span.Is("italic") || span.Is("oblique") => FontStyle.Italic,
            _ when span.Is("normal") => FontStyle.Normal,
            _ => null
        };

        return style is not null;
    }

    public static bool TryParseFontWeight(ReadOnlySpan<char> span,
        [NotNullWhen(true)] out FontWeight? weight)
    {
        weight = span switch
        {
            _ when span.IsEmpty => null,
            _ when span.Is("700") || span.Is("bold") => FontWeight.Bold,
            _ when span.Is("bolder") => FontWeight.Bolder,
            _ when span.Is("400") || span.Is("normal") => FontWeight.Normal,
            _ => null
        };

        return weight is not null;
    }

    public static bool TryParseFontFamily(ReadOnlySpan<char> span,
        [NotNullWhen(true)] out string? family)
    {
        // return the first font name
        Span<Range> tokens = stackalloc Range[1];
        if (span.SplitCompositeAttribute(tokens, ',') == 1)
        {
            family = span[tokens[0]].Unquote().Trim().ToString();
            return true;
        }

        family = null;
        return false;
    }

    public static bool TryParseBorderStyle(ReadOnlySpan<char> span,
        [NotNullWhen(true)] out BorderValues? border)
    {
        border = span switch
        {
            _ when span.IsEmpty => null,
            _ when span.Is("dotted") => BorderValues.Dotted,
            _ when span.Is("dashed") => BorderValues.Dashed,
            _ when span.Is("solid") => BorderValues.Single,
            _ when span.Is("double") => BorderValues.Double,
            _ when span.Is("inset") => BorderValues.Inset,
            _ when span.Is("outset") => BorderValues.Outset,
            _ when span.Is("none") => BorderValues.None,
            _ => null
        };

        return border is not null;
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
            if (token.Is("underline"))
                decorations |= TextDecoration.Underline;
            else if (token.Is("line-through"))
                decorations |= TextDecoration.LineThrough;
            else if (token.Is("double"))
                decorations |= TextDecoration.Double;
            else if (token.Is("dotted"))
                decorations |= TextDecoration.Dotted;
            else if (token.Is("dashed"))
                decorations |= TextDecoration.Dashed;
            else if (token.Is("wavy"))
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