/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
namespace HtmlToOpenXml;

/// <summary>
/// Specifies the measurement values of a Html Unit.
/// </summary>
enum UnitMetric
{
    Unknown,
    Percent,
    Inch,
    Centimeter,
    Millimeter,
    /// <summary>1em is equal to the current font size.</summary>
    EM,
    /// <summary>one ex is the x-height of a font (x-height is usually about half the font-size)</summary>
    Ex,
    Point,
    Pica,
    Pixel,

    // this value is not parsed but can be used internally
    Emus,
    // this value is not parsed but can be used internally for paper size
    Twips,

    /// <summary>Not convertible to any other units.</summary>
    Auto,
    /// <summary>Raw value, not convertible to any other units</summary>
    Unitless
}
