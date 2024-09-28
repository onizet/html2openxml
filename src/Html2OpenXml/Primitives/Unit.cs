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
/// Represents a Html Unit (ie: 120px, 10em, ...).
/// </summary>
[System.Diagnostics.DebuggerDisplay("Unit: {Value} {Type}")]
readonly struct Unit
{
    /// <summary>Represents an empty unit (not defined).</summary>
    public static readonly Unit Empty = new Unit();
    /// <summary>Represents an Auto unit.</summary>
    public static readonly Unit Auto = new Unit(UnitMetric.Auto, 0L);

    private readonly UnitMetric metric;
    private readonly double value;
    private readonly long valueInEmus;


    public Unit(UnitMetric metric, double value)
    {
        this.metric = metric;
        this.value = value;
        this.valueInEmus = ComputeInEmus(metric, value);
    }

    public static Unit Parse(ReadOnlySpan<char> span, UnitMetric defaultMetric = UnitMetric.Unitless)
    {
        span.Trim();
        if (span.Length <= 1)
        {
            // either this is invalid or this is a single digit
            if (span.Length == 0 || !char.IsDigit(span[0])) return Empty;
            return new Unit(defaultMetric, span[0] - '0');
        }

        // guess the unit first than use the native Double parsing
        UnitMetric metric;
        int metricSize = 2;
        if (span[span.Length - 1] == '%')
        {
            metric = UnitMetric.Percent;
            metricSize = 1;
        }
        else
        {
            Span<char> loweredValue = span.Length <= 128 ? stackalloc char[span.Length] : new char[span.Length];
            span.ToLowerInvariant(loweredValue);

            var metricSpan = loweredValue.Slice(loweredValue.Length - 2, 2);
            metric = metricSpan switch {
                "in" => UnitMetric.Inch,
                "cm" => UnitMetric.Centimeter,
                "mm" => UnitMetric.Millimeter,
                "em" => UnitMetric.EM,
                "ex" => UnitMetric.Ex,
                "pt" => UnitMetric.Point,
                "pc" => UnitMetric.Pica,
                "px" => UnitMetric.Pixel,
                _ =>  UnitMetric.Unknown,
            };

            // not recognised but maybe this is unitless (only digits) 
            if (metric == UnitMetric.Unknown && char.IsDigit(metricSpan[0]))
            {
                metric = UnitMetric.Unitless;
                metricSize = 0;
            }
        }

        double value;
        try
        {
            value = span.Slice(0, span.Length - metricSize).AsDouble();

            if (value < short.MinValue || value > short.MaxValue)
                return Empty;
        }
        catch (Exception)
        {
            // No digits, we ignore this style
            return span is "auto"? Auto : Empty;
        }

        return new Unit(metric, value);
    }

    public static Unit Parse(string? str, UnitMetric defaultMetric = UnitMetric.Unitless)
    {
        if (string.IsNullOrWhiteSpace(str))
            return Empty;

        return Parse(str.AsSpan(), defaultMetric);
    }

    /// <summary>
    /// Gets the value expressed in the English Metrics Units.
    /// </summary>
    private static long ComputeInEmus(UnitMetric metric, double value)
    {
        /* Compute width and height in English Metrics Units.
         * There are 360000 EMUs per centimeter, 914400 EMUs per inch, 12700 EMUs per point
         * widthInEmus = widthInPixels / HorizontalResolutionInDPI * 914400
         * heightInEmus = heightInPixels / VerticalResolutionInDPI * 914400
         * 
         * According to 1 px ~= 9525 EMU -> 914400 EMU per inch / 9525 EMU = 96 dpi
         * So Word use 96 DPI printing which seems fair.
         * http://hastobe.net/blogs/stevemorgan/archive/2008/09/15/howto-insert-an-image-into-a-word-document-and-display-it-using-openxml.aspx
         * http://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
         *
         * The list of units supported are explained here: http://www.w3schools.com/css/css_units.asp
         */

        switch (metric)
        {
            case UnitMetric.Auto:
            case UnitMetric.Unitless:
            case UnitMetric.Percent: return 0L; // not applicable
            case UnitMetric.Emus: return (long) value;
            case UnitMetric.Inch: return (long) (value * 914400L);
            case UnitMetric.Centimeter: return (long) (value * 360000L);
            case UnitMetric.Millimeter: return (long) (value * 36000L);
            case UnitMetric.EM:
                // well this is a rough conversion but considering 1em = 12pt (http://sureshjain.wordpress.com/2007/07/06/53/)    
                return (long) (value / 72 * 914400L * 12);
            case UnitMetric.Ex:
                return (long) (value / 72 * 914400L * 12) / 2;
            case UnitMetric.Point: return (long) (value * 12700L);
            case UnitMetric.Pica: return (long) (value / 72 * 914400L) * 12;
            case UnitMetric.Pixel: return (long) (value / 96 * 914400L);
            default: goto case UnitMetric.Pixel;
        }
    }

    //____________________________________________________________________
    //

    /// <summary>
    /// Gets the type of unit (pixel, percent, point, ...)
    /// </summary>
    public UnitMetric Metric
    {
        get { return metric; }
    }

    /// <summary>
    /// Gets the value of this unit.
    /// </summary>
    public Double Value
    {
        get { return value; }
    }

    /// <summary>
    /// Gets the value expressed in English Metrics Unit.
    /// </summary>
    public Int64 ValueInEmus
    {
        get { return valueInEmus; }
    }

    /// <summary>
    /// Gets the value expressed in Dxa unit.
    /// </summary>
    public Int64 ValueInDxa
    {
        get { return (long) (((double) valueInEmus / 914400L) * 20 * 72); }
    }

    /// <summary>
    /// Gets the value expressed in Pixel unit.
    /// </summary>
    public int ValueInPx
    {
        get { return (int) (metric == UnitMetric.Pixel ? this.value : (float) valueInEmus / 914400L * 96); }
    }

    /// <summary>
    /// Gets the value expressed in Point unit.
    /// </summary>
    public double ValueInPoint
    {
        get { return (double) (metric == UnitMetric.Point ? this.value : (float) valueInEmus / 12700L); }
    }

    /// <summary>
    /// Gets the value expressed in 1/8 of a Point
    /// IMPORTANT: Use this for borders, as OpenXML expresses Border Width in 1/8 of points,
    /// with a minimum value of 2 (1/4 of a point) and a maximum value of 96 (12 points).
    /// </summary>
    public double ValueInEighthPoint
    {
        get { return ValueInPoint * 8; }
    }

    /// <summary>
    /// Gets whether the unit is well formed and not empty.
    /// </summary>
    public bool IsValid
    {
        get { return this.Metric != UnitMetric.Unknown; }
    }

    /// <summary>
    /// Gets whether the unit is well formed and not absolute nor auto.
    /// </summary>
    public bool IsFixed
    {
        get { return IsValid && Metric != UnitMetric.Percent && Metric != UnitMetric.Auto; }
    }
}
